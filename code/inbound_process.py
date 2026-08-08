"""动态计算在途货件预计到仓日期。

这个脚本替代旧 notebook 中依赖「预计到达.xlsx」的部分：
1. 优先使用物流追踪表中的真实节点时间。
2. 真实节点缺失时，用标准时效参数补齐。
3. 查验未放行货件按登记表决定是否计入在途库存。

输入数据：
- 发货单清单：提供货件维度的状态、物流方式、提货日期等信息。
- 物流追踪更新：提供货件的物流追踪节点时间。
- 标准时效参数表：提供不同国家、物流商、渠道、地区的标准物流时效。
输出数据：
- 在途货件-正常：包含预计到仓日期和预计出运日期的在途货件明细。
- 在途货件-接收中：入库中货件的明细和汇总，方便跟踪接收进度。
- 物流时效缺失详情：记录缺失标准时效参数的货件信息，方便后续补齐参数表。
过程异常输出：
- 重复FBA ID 记录：如果物流追踪表中存在同一 FBA ID 对应多条记录，输出这些重复记录到单独文件，方便检查数据质量。
- 目的国家为US但地区缺失的记录：发货单中目的国家为US但地区信息缺失的记录，输出到单独文件并报错，要求补齐地区信息后重跑。
- 超过预占天数的未出库调拨单：输出到单独文件，提示这些单据暂不纳入正常在途库存。
- 查验货件登记表中存在纳入库存计算但签收日期为空的货件：报错提示检查登记表数据完整性。
"""

from pathlib import Path
import sys
import pandas as pd



# 默认处理日期；命令行传日期时会覆盖它。
DEFAULT_DATE = "20260601"
# 未出库调拨单超过该天数后，暂不纳入正常在途库存。
EXCEPTION_DAYS = 15

# 所有数据文件都按项目根目录定位，避免依赖当前运行目录。
ROOT_DIR = Path(__file__).resolve().parents[1]
INBOUND_DIR = ROOT_DIR / "src_data" / "在途货件"
YUZHAN_DIR = ROOT_DIR / "src_data" / "预占单"
INSPECTION_PATH = ROOT_DIR / "src_data" / "2026 查验货件登记表.xlsx"

# 标准时效表的匹配键、内部参数键，以及三段物流时效。
STANDARD_KEYS = ["国家", "物流商", "物流渠道", "地区"]
PARAM_KEYS = ["参数国家", "参数物流商", "参数物流渠道", "参数地区"]
TIME_SEGMENTS = ["入库-离港", "离港-到港", "到港-签收"]

# 物流追踪表里可能是中文国家名，标准参数表使用站点简码。
COUNTRY_MAP = {
    "美国": "US",
    "加拿大": "CA",
    "英国": "UK",
    "德国": "DE",
    "日本": "JP",
    "欧盟": "DE",
    "EU": "DE",
}

# 发货单里常见的是物流商简称，标准参数表使用中文物流商名称。
VENDOR_MAP = {
    "DS": "德速",
    "DESU_FIRST-1": "德速",
    "YB": "一八",
    "YFH": "原飞航",
    "MT": "美通",
    "SS": "速十",
    "SF": "顺丰",
    "CHENGE": "陈哥",
    "HAIY": "海源",
}


def clean_text(value):
    """把单元格值统一清洗成无首尾空格的字符串。"""
    if pd.isna(value):
        return ""
    return str(value).strip().strip('"').strip()


def clean_series(series):
    """批量清洗文本列，用于 join key 和渠道名。"""
    return series.map(clean_text)


def normalize_country(value):
    """把国家字段统一成标准参数表可匹配的格式。"""
    text = clean_text(value)
    return COUNTRY_MAP.get(text, text)


def normalize_vendor(value):
    """把物流商字段统一成标准参数表可匹配的格式。"""
    text = clean_text(value)
    return VENDOR_MAP.get(text, text)


def parse_date_series(series, actual_only=False):
    """从 Excel 日期列中提取日期。

    物流追踪表里会出现「预计2026-05-17」这类文本。
    这种预计节点还是当作真实物流时间使用，将预计去掉。
   
    """
    text = series.astype("string")
    extracted = text.str.extract(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})")
    date_text = extracted[0] + "-" + extracted[1].str.zfill(2) + "-" + extracted[2].str.zfill(2)
    parsed = pd.to_datetime(date_text, format="%Y-%m-%d", errors="coerce")
    # if actual_only:
    #     parsed = parsed.mask(text.str.contains("预计", na=False))  # 如果只要真实节点，遇到带“预计”的文本就当作缺失处理；但实际测试中发现有些追踪表里节点时间前面带了“预计”但其实是准确的真实节点，所以先不区分，统一提取日期部分作为实际时间。
    return parsed


def read_inspection_rules():
    """读取查验货件规则。

    返回两个结果：
    - excluded：不纳入库存计算的货件号集合。
    - override_eta：纳入库存计算的货件号及指定签收日期。
    """
    if not INSPECTION_PATH.exists():
        return set(), pd.Series(dtype="datetime64[ns]")

    inspection = pd.read_excel(
        INSPECTION_PATH,
        sheet_name="汇总",
        usecols=["FBA ID", "是否纳入库存计算", "纳入库存计算后签收日期"],
    )
    
    # FBA ID 单元格可能夹带备注或运单号，只抽取 FBA 开头的货件号。
    inspection["货件号"] = inspection["FBA ID"].astype(str).str.extract(r"(FBA[0-9A-Z]+)", expand=False)

    # 不纳入库存计算：从发货单主表剔除，后续不会进入在途库存。
    excluded = set(
        inspection.loc[
            inspection["是否纳入库存计算"].astype(str).str.strip().eq("否"),
            "货件号",
        ].dropna()
    )

    # 纳入库存计算：保留货件，但预计到仓日期直接使用登记表指定日期。
    included = inspection[inspection["是否纳入库存计算"].astype(str).str.strip().eq("是")].copy()
    included["查验预计到仓日期"] = parse_date_series(included["纳入库存计算后签收日期"])
    if included["查验预计到仓日期"].isna().any():
        raise ValueError("查验货件登记表中存在纳入库存计算但签收日期为空的货件，请检查")

    override_eta = included.dropna(subset=["货件号"]).drop_duplicates("货件号").set_index("货件号")["查验预计到仓日期"]
    
    # print(excluded)
    # print(override_eta)
    
    return excluded, override_eta


def read_standard_params():
    """读取并清洗标准时效参数表。""" 
    path = INBOUND_DIR / "标准时效参数表.xlsx"
    standard = pd.read_excel(path, sheet_name="标准时效", usecols=STANDARD_KEYS + TIME_SEGMENTS)

    # 标准表本身也做一次清洗，保证和发货单/追踪表的 key 口径一致。
    standard["国家"] = standard["国家"].map(normalize_country)
    standard["物流商"] = standard["物流商"].map(normalize_vendor)
    standard["物流渠道"] = clean_series(standard["物流渠道"])
    standard["地区"] = clean_series(standard["地区"])

    for segment in TIME_SEGMENTS:
        # 时效必须是数字天数，非数字会变成空值并进入缺参明细。
        standard[segment] = pd.to_numeric(standard[segment], errors="coerce")

    return standard.drop_duplicates(subset=STANDARD_KEYS)


def calculate_actual_segments(df):
    """根据物流追踪节点计算三段真实物流时效。"""
    df["入库-离港_实际天数"] = (df["离港时间_实际"] - df["入库时间_实际"]).dt.days
    df["离港-到港_实际天数"] = (df["到港时间_实际"] - df["离港时间_实际"]).dt.days
    df["到港-签收_实际天数"] = (df["签收时间_实际"] - df["到港时间_实际"]).dt.days
    return df


def add_eta_by_standard(df, standard, source_name):
    """合并标准时效，并计算预计到仓日期。

    每段时效的优先级：
    1. 物流追踪真实节点差值。
    2. 标准时效参数表。
    3. 查验表指定签收日期可直接覆盖最终到仓日期。
    """
    # 用国家、物流商、物流渠道、地区匹配标准三段时效。
    df = df.merge(standard, how="left", left_on=PARAM_KEYS, right_on=STANDARD_KEYS)
    has_override = df["查验预计到仓日期"].notna() if "查验预计到仓日期" in df.columns else pd.Series(False, index=df.index)

    for segment in TIME_SEGMENTS:
        actual_col = f"{segment}_实际天数"
        if actual_col not in df.columns:
            df[actual_col] = pd.NA
        # 真实时效优先；真实时效为空时，用标准参数补齐。
        df[f"{segment}_最终天数"] = pd.to_numeric(df[actual_col], errors="coerce").combine_first(df[segment])

    # 某一段既没有真实时效、也没有标准参数，就记录为缺参。
    missing_mask = pd.Series(False, index=df.index)
    for segment in TIME_SEGMENTS:
        missing_mask |= df[f"{segment}_实际天数"].isna() & df[segment].isna()
    # 查验表已经指定到仓日期的货件，不再要求标准时效齐全。
    missing_mask &= ~has_override

    missing = build_missing_detail(df[missing_mask].copy(), source_name)

    # 三段时效都存在时才能计算到仓日期。
    total_days = df[[f"{segment}_最终天数" for segment in TIME_SEGMENTS]].sum(axis=1, min_count=3)
    calculated_eta = df["预计出运日期"] + pd.to_timedelta(total_days, unit="D")
    # 查验表指定日期优先级最高。
    df["预计到仓日期"] = df["查验预计到仓日期"].combine_first(calculated_eta) if "查验预计到仓日期" in df.columns else calculated_eta
    return df, missing


def build_missing_detail(df, source_name):
    """整理标准时效缺失明细，方便后续补参数表。"""
    if df.empty:
        return pd.DataFrame()

    def missing_segments(row):
        """列出当前行缺失的时效段。"""
        segments = []
        for segment in TIME_SEGMENTS:
            if pd.isna(row.get(f"{segment}_实际天数")) and pd.isna(row.get(segment)):
                segments.append(segment)
        return ",".join(segments)

    df["数据来源"] = source_name
    df["缺失标准时效段"] = df.apply(missing_segments, axis=1)

    # 不同来源字段略有差异，只导出实际存在的定位字段。
    export_cols = [
        "数据来源",
        "发货单号",
        "调拨单号",
        "货件号",
        "物流跟踪号",
        "参数国家",
        "参数物流商",
        "参数物流渠道",
        "参数地区",
        "缺失标准时效段",
    ]
    export_cols = [col for col in export_cols if col in df.columns]
    missing = df[export_cols].drop_duplicates()
    return missing.rename(
        columns={
            "参数国家": "国家",
            "参数物流商": "物流商",
            "参数物流渠道": "物流渠道",
            "参数地区": "地区",
        }
    )


def read_tracking(date_parm):
    """读取物流追踪汇总表。"""
    path = INBOUND_DIR / f"{date_parm} 物流追踪更新.xlsx"
    tracking = pd.read_excel(
        path,
        sheet_name="物流追踪汇总（无公式版）",
        usecols=[
            "物流商",
            "国家",
            "物流渠道",
            "货运方式",
            # "物流运单号",
            "FBA ID",
            "入库时间",
            "离港时间",
            "到港时间",
            "签收时间",
        ],
    )

    # 追踪表字段和发货单字段重名较多，先加前缀避免 merge 后含义混乱。
    tracking = tracking.rename(
        columns={
            "物流商": "追踪物流商",
            "国家": "追踪国家",
            "物流渠道": "追踪物流渠道",
            "货运方式": "追踪货运方式",
        }
    )
    # join key 做文本清洗，避免 Excel 空格或引号导致匹配失败。
    tracking["FBA ID_key"] = clean_series(tracking["FBA ID"])
    # 只取FBA ID_key的前12个字符，因为有的追踪表里这个字段夹带了运单号或备注信息。
    tracking["FBA ID_key"] = tracking["FBA ID_key"].str[:12]
    # 保存FBA ID_key去重前的所有重复行记录到表中，方便后续检查追踪表数据质量。
    tracking_duplicates = tracking[tracking.duplicated("FBA ID_key", keep=False)]
    tracking_duplicates.to_excel(INBOUND_DIR / f"{date_parm} 物流追踪重复记录.xlsx", index=False)
    # 根据FBA ID_key去重，保留一条记录，避免重复追踪信息导致匹配到多个节点。
    tracking = tracking.drop_duplicates("FBA ID_key", keep="first")
    
    return tracking


def build_inbound_data(date_parm, run_date, standard):
    """构建发货单来源的在途货件和接收中货件数据。

    date_parm 用于定位输入文件，
    run_date 当天日期，用于计算接收天数，
    standard 标准时效参数表，用于补齐缺失物流追踪时效。

    """
    inbound_path = INBOUND_DIR / f"发货单清单{date_parm}.xlsx"
    path = INBOUND_DIR / "标准时效参数表.xlsx"

    # 读取地区映射表，直接当作参数表的一个匹配字段。
    fba_area = pd.read_excel(path, sheet_name="地区映射表", usecols=["FBA仓库", "地区"])

    # inspection_excluded不纳入库存计算的货件集合
    inspection_excluded, inspection_eta = read_inspection_rules()
    print("不纳入库存计算的FBAid：", inspection_excluded)
    print("纳入库存计算的FBAid及到仓日期：", inspection_eta)

    # 主表负责货件维度的状态、物流方式和物流节点。
    inbound_total = pd.read_excel(
        inbound_path,
        sheet_name="system_发货单",
        usecols=[
            "发货单号",
            "目的国家",
            "目的仓",
            "物流方式",
            "物流商",
            "状态",
            "提货日期",
            "实际到仓日期",
            "货件号",
            "货件目的仓库",
            # "物流跟踪号",
        ],
    )
    # 明细表负责 SKU/MSKU 和发货量。
    inbound_detail = pd.read_excel(
        inbound_path, 
        sheet_name="system_发货单明细",
        usecols=["发货单号", "仓库", "ShipmentId", "SKU", "MSKU", "发货量", "差异量"],
    )
    inbound_detail['仓库'] = inbound_detail['仓库'].str.replace('EU', 'DE')

    # ---------------------只保留 FBA 在途相关状态。---------------------------
    status_mask = inbound_total["状态"].isin(["已出运", "入库中", "提货中"])
    country_mask = inbound_total["目的仓"].astype(str).str.contains("FBA", na=False)
    inbound_total = inbound_total[status_mask & country_mask].copy()

    # 将货件号和货件目的仓库中，如果货件号存在多个货件号（例如一个发货单拆成多个货件），先按分隔符分开成多行，再做清洗和匹配。
    # 将两列zip后一起explode（推荐，保持严格对齐）。
    inbound_total = inbound_total.assign(
        货件号=inbound_total["货件号"].str.split(","),
        货件目的仓库=inbound_total["货件目的仓库"].str.split(",")
    ).explode(["货件号", "货件目的仓库"], ignore_index=True)
    # 以上：这里会将1个发货单拆成多行，说明发货单有重复行，不能作为唯一键去匹配发货单明细表了。解决就是：发货单+FBAid 一起作为唯一键去匹配发货单明细表。

    # 先把US站地区信息合并到发货单主表，作为后续标准时效匹配的一个参数字段。
    inbound_total = inbound_total.merge(
        fba_area.rename(columns={"FBA仓库": "货件目的仓库"}), 
        how="left", 
        on="货件目的仓库"
    )
    # 将目的国家不为US的“地区”列填充为“非美国”。
    inbound_total["地区"] = inbound_total["地区"].where(inbound_total["目的国家"].eq("US"), "非美国")
   

    # 将目的国家为US但地区缺失的行输出到异常表并报错，方便后续检查和补齐地区信息。
    missing_area = inbound_total[inbound_total["目的国家"].eq("US") & inbound_total["地区"].isnull()]
    if not missing_area.empty:
        missing_area.to_excel(INBOUND_DIR / f"{date_parm} 目的国家为US但地区缺失的记录.xlsx", index=False)
        print(f"{date_parm} 目的国家为US但地区缺失的记录已输出到 {INBOUND_DIR / f'{date_parm} 目的国家为US但地区缺失的记录.xlsx'}，请检查并补齐地区信息后重跑")
        raise ValueError(f"{date_parm} 目的国家为US但地区缺失的记录已输出到 {INBOUND_DIR / f'{date_parm} 目的国家为US但地区缺失的记录.xlsx'}，请检查并补齐地区信息后重跑")
    else:
        print(f"{date_parm} 目的国家为US但地区缺失的记录为空，无需输出异常表")

    # 查验表：不纳入库存的货件直接剔除；纳入库存的货件保留并覆盖到仓日期。
    inbound_total["货件号_key"] = clean_series(inbound_total["货件号"])
    # inbound_total = inbound_total[~inbound_total["货件号_key"].isin(inspection_excluded)].copy()
    # inbound_total["查验预计到仓日期"] = inbound_total["货件号_key"].map(inspection_eta)

    # 清洗后续匹配与计算需要的字段。
    inbound_total["物流方式"] = clean_series(inbound_total["物流方式"])
    inbound_total["物流商"] = inbound_total["物流商"].map(normalize_vendor)
    inbound_total["提货日期"] = parse_date_series(inbound_total["提货日期"])
    inbound_total["实际到仓日期"] = parse_date_series(inbound_total["实际到仓日期"])
    inbound_total["接收天数"] = (run_date - inbound_total["实际到仓日期"]).dt.days.fillna(0)

    # 用货件号 关联真实物流追踪信息。
    tracking = read_tracking(date_parm)

    inbound = inbound_total.merge(
        tracking,
        how="left",
        left_on=["货件号_key"],
        right_on=["FBA ID_key"],
        indicator=True,  # 标记 merge 结果，方便后续判断哪些货件有物流追踪信息，“both”表示发货单和追踪表都有匹配记录，“left_only”表示只有发货单有记录但追踪表没有，“right_only”表示只有追踪表有记录但发货单没有（理论上不应该出现）。
    )
    inbound["有物流追踪"] = inbound["_merge"].eq("both") # 只要能匹配到追踪表记录就算有物流追踪，不要求所有节点都存在。

    print(f"发货单总表与物流追踪表合并后-列名: {inbound.columns.tolist()}\n")

    # 物流追踪节点中带“预计”的时间不算真实时间。
    for col in ["入库时间", "离港时间", "到港时间", "签收时间"]:
        inbound[f"{col}_实际"] = parse_date_series(inbound[col], actual_only=True)

    inbound = calculate_actual_segments(inbound)
    # 有追踪时出运日期取入库时间；没有追踪时退回发货单提货日期。
    inbound["预计出运日期"] = inbound["入库时间_实际"].where(inbound["有物流追踪"], inbound["提货日期"])
    inbound["预计出运日期"] = inbound["预计出运日期"].fillna(inbound["提货日期"])

    # 标准时效匹配参数：有追踪时优先用追踪表渠道和地区，否则用发货单字段。
    inbound["参数国家"] = inbound["目的国家"] 
    inbound["参数物流商"] = inbound["追踪物流商"].where(inbound["有物流追踪"], inbound["物流商"]).map(normalize_vendor)  # 有物流追踪时，用追踪表里的物流商作为 参数物流商；没有物流追踪时，退回用发货单的物流商作为 参数物流商。
    inbound["参数物流渠道"] = clean_series(inbound["追踪物流渠道"].where(inbound["有物流追踪"], inbound["物流方式"]))
    inbound["参数地区"] = inbound["地区"]   # 直接用发货单的地区字段作为 参数地区，因为追踪表的地区信息存在不完整。
    # 将“地区”列删除
    inbound = inbound.drop(columns=["地区"])

    inbound, missing = add_eta_by_standard(inbound, standard, "在途货件")

    print(f"在途货件-发货单总表列名: {inbound.columns.tolist()}")

    # "发货单号","货件号" 维度结果再和 SKU 明细合并，形成输出需要的明细粒度。
    inbound_ret = inbound[["发货单号", "货件号", "物流方式", "实际到仓日期", "预计到仓日期", "预计出运日期", "接收天数", "状态", "离港时间", "到港时间", "签收时间"]]
    merge_df = inbound_ret.merge(inbound_detail, left_on=["发货单号","货件号"], right_on=["发货单号","ShipmentId"], how="left")
    # 删除“货件号”列
    merge_df = merge_df.drop(columns=["货件号"])
    print(f"发货单总表与发货单明细表合并后数据列名: {merge_df.columns.tolist()}")

    # 入库中货件单独输出接收中报表；其余进入正常在途库存。
    fba_inbound = merge_df[merge_df["状态"] != "入库中"].copy()

    fba_receive_detail = merge_df[merge_df["状态"] == "入库中"].copy()
    fba_receive_detail = fba_receive_detail.drop(columns=["离港时间", "到港时间", "签收时间"])

    
    # 接收中汇总沿用旧输出口径：仓库去掉 _FBA 后作为店铺站点。
    fba_receive_detail = fba_receive_detail.rename(columns={"仓库": "店铺-站点"})
    fba_receive_detail["店铺-站点"] = fba_receive_detail["店铺-站点"].str.replace("_FBA", "", regex=False)
    fba_receive_detail["已收量"] = fba_receive_detail["发货量"] - fba_receive_detail["差异量"]

    # 筛选出货件接收7天内且差异量>0的行，将该差异量计入为FBA可售库存

    # # 条件
    # condition = (fba_receive_detail['接收天数'] <= 7) & (fba_receive_detail['差异量'] > 0)
    # # 新增列
    # fba_receive_detail['是否计入可售'] = np.where(condition, '是', '否')

    fba_receive_detail['是否已计入可售'] = '否'
    fba_receive_detail.loc[
        (fba_receive_detail['接收天数'] <= 7) & (fba_receive_detail['差异量'] > 0),
        '是否已计入可售'
    ] = '是'

    
    # 保留不纳入库存计算的查验数据
    inspection_excluded_df = fba_inbound[fba_inbound["ShipmentId"].isin(inspection_excluded)]
    inspection_excluded_df = inspection_excluded_df[['仓库', 'SKU', 'MSKU', '发货量', '预计出运日期', '状态', '物流方式', 'ShipmentId']]

    inspection_excluded_total = (
            inspection_excluded_df.pivot_table(
                index=["仓库", "SKU", "MSKU"],
                values=["发货量"],
                aggfunc="sum",
            )
            .reset_index()
        )

    # 查验表：不纳入库存的货件直接剔除；纳入库存的货件保留并覆盖到仓日期。
    fba_inbound = fba_inbound[~fba_inbound["ShipmentId"].isin(inspection_excluded)].copy()
    print("fba_inbound 数量：", len(fba_inbound))

    fba_inbound["查验预计到仓日期"] = fba_inbound["ShipmentId"].map(inspection_eta)
    # 查验表指定日期优先级最高。
    fba_inbound["预计到仓日期"] = fba_inbound["查验预计到仓日期"].combine_first(fba_inbound["预计到仓日期"]) if "查验预计到仓日期" in fba_inbound.columns else fba_inbound["预计到仓日期"]

    # 如果预计到仓日期早于运行日期，说明实际还未到仓，则预计到仓日期加7天，若+7天还是早于运行日期，则先不计入在途。
    future_mask = fba_inbound["预计到仓日期"] < run_date
    # 将这部分数据输出到异常表，方便后续检查数据质量。
    if future_mask.any():
        future_records_detail = fba_inbound[future_mask].copy()
        future_records_detail["预计到仓日期+7天"] = future_records_detail["预计到仓日期"] + pd.Timedelta(days=7)
        # 新增列，说明是否计入在途，预计到仓日期+7天正常后则计入在途，否则不计入在途。
        future_records_detail["是否计入在途"] = (
            future_records_detail["预计到仓日期+7天"]    # 如果 预计到仓日期+7天 >= run_date，则该行 是否计入在途 为 True，否则为 False。
            .ge(run_date)
            .map({True: "是", False: "否"})
        )
        future_records_detail = future_records_detail[["发货单号", "ShipmentId", "物流方式", "预计出运日期", "预计到仓日期", "状态", "仓库", "SKU","MSKU", "发货量", "差异量", "预计到仓日期+7天", "是否计入在途", "离港时间", "到港时间", "签收时间"]]
        # future_records_detail.to_excel(INBOUND_DIR / f"{date_parm} 在途中货件预计到仓日期早于运行日期的记录.xlsx", index=False)
        future_records_total = (
            future_records_detail.pivot_table(
                index=["仓库", "SKU", "MSKU","是否计入在途"],
                values=["发货量", "差异量"],
                aggfunc="sum",
            )
            .reset_index()
        )
    else:
        print(f"{date_parm} 在途中货件预计到仓日期早于运行日期的记录为空，无需输出异常表。\n")

    fba_inbound.loc[future_mask, "预计到仓日期"] = fba_inbound.loc[future_mask, "预计到仓日期"] + pd.Timedelta(days=7)
    fba_inbound = fba_inbound.drop(columns=["离港时间", "到港时间", "签收时间"])

    if fba_receive_detail.empty:
        fba_receive_total = pd.DataFrame(columns=["店铺-站点", "SKU", "MSKU", "发货量", "差异量", "已收量"])
    else:
        fba_receive_total = (
            fba_receive_detail.pivot_table(
                index=["店铺-站点", "SKU", "MSKU", "是否已计入可售"],
                values=["发货量", "差异量", "已收量"],
                aggfunc="sum",
            )
            .reset_index()
        )

    return fba_inbound, fba_receive_total, fba_receive_detail, missing, future_records_detail,future_records_total, inspection_excluded_df, inspection_excluded_total


def build_yuzhan_data(date_parm, run_date, standard):
    """构建未出库调拨单来源的预占在途数据。"""
    yuzhan_path = f"E:/sontu/shipping_recommendations/src_data/预占单/未出库调拨单导出{date_parm}.xlsx"

    # 明细表提供 SKU/MSKU 和数量；单据表提供仓库、状态、物流方式。
    yuzhan_detail = pd.read_excel(
        yuzhan_path,
        sheet_name="明细数据",
        usecols=["SKU", "MSKU", "调出数量", "调拨单号"],
    )
    yuzhan_total = pd.read_excel(
        yuzhan_path,
        sheet_name="单据数据",
        usecols=[
            "调拨单号",
            "调出仓",
            "调入仓",
            "物流方式",
            "单据状态",
            "拣货状态",
            "出库状态",
            "预计出运日期",
            "头程单/ShipmentId",
        ],
    )

    # 数量为 0 的明细不进入在途库存。
    yuzhan_detail = yuzhan_detail[yuzhan_detail["调出数量"] > 0].copy()
    yuzhan_total["预计出运日期"] = parse_date_series(yuzhan_total["预计出运日期"])

    # 保留已审核、未出库、调入 FBA 的调拨单。
    mask = (
        yuzhan_total["调入仓"].astype(str).str.contains("FBA", na=False)
        & yuzhan_total["单据状态"].eq("审核通过")
        & yuzhan_total["出库状态"].eq("未出库")
    )
    yuzhan_total = yuzhan_total[mask].copy()
    yuzhan_total["预占天数"] = (run_date - yuzhan_total["预计出运日期"]).dt.days.fillna(0)

    # 超过预占天数阈值的单据不进正常在途，但会输出到异常 sheet。
    yuzhan_true = yuzhan_total[yuzhan_total["预占天数"] <= EXCEPTION_DAYS].copy()
    yuzhan_false = yuzhan_total[yuzhan_total["预占天数"] > EXCEPTION_DAYS].copy()

    if yuzhan_true.empty:
        return pd.DataFrame(), yuzhan_false, pd.DataFrame()

    # 未出库调拨单没有物流商和地区，按需求只用国家 + 物流方式匹配标准时效。
    yuzhan_true["参数国家"] = yuzhan_true["调入仓"].astype(str).str.extract(r".*:(.{2})_*")[0].map(normalize_country)
    yuzhan_true["参数物流商"] = ""
    yuzhan_true["参数物流渠道"] = clean_series(yuzhan_true["物流方式"])
    yuzhan_true["参数地区"] = ""

    for segment in TIME_SEGMENTS:
        # 未出库调拨单没有真实物流节点，三段时效全部依赖标准参数。
        yuzhan_true[f"{segment}_实际天数"] = pd.NA

    yuzhan_true, missing = add_eta_by_standard(yuzhan_true, standard, "未出库调拨单")
    yuzhan_true = yuzhan_true.merge(yuzhan_detail, on="调拨单号", how="inner")

    # 调拨单可能没有 ShipmentId，保留一个固定占位值方便后续识别来源。
    shipment_id = clean_series(yuzhan_true["头程单/ShipmentId"])
    yuzhan_true["ShipmentId"] = shipment_id.mask(shipment_id.eq(""), "预占单无ShipmentId")
    yuzhan_true["发货单号"] = shipment_id.mask(shipment_id.eq(""), "预占单无发货单号")

    # 对齐最终「在途货件-正常」sheet 的标准列。
    output = yuzhan_true[
        [
            "调入仓",
            "SKU",
            "MSKU",
            "调出数量",
            "预计到仓日期",
            "预计出运日期",
            "出库状态",
            "物流方式",
            "ShipmentId",
            "发货单号",
        ]
    ].copy()
    output.columns = ["仓库", "SKU", "MSKU", "发货量", "预计到仓日期", "预计出运日期", "状态", "物流方式", "ShipmentId", "发货单号"]
    return output, yuzhan_false, missing


def write_missing_and_raise(missing_frames):
    """如果存在缺失标准时效，先输出明细文件再中止运行。"""
    missing_frames = [frame for frame in missing_frames if not frame.empty]
    if not missing_frames:
        return

    # 缺参时不生成最终在途表，避免下游使用不完整结果。
    missing = pd.concat(missing_frames, axis=0, ignore_index=True).drop_duplicates()
    missing_path = INBOUND_DIR / "物流时效缺失详情.xlsx"
    missing.to_excel(missing_path, index=False)
    raise ValueError(f"{missing_path} 文件已生成，请补齐标准时效参数表后重跑")


def write_output(date_parm, fba_inbound, fba_receive_total, fba_receive_detail, yuzhan_output, yuzhan_false, future_records_detail, future_records_total, inspection_excluded_df, inspection_excluded_total):
    """写出最终在途货件 Excel，保持旧文件结构不变。"""
    standard_titles = ["仓库", "SKU", "MSKU", "发货量", "预计到仓日期", "预计出运日期", "状态", "物流方式", "ShipmentId", "发货单号"]

    df1 = fba_inbound[standard_titles].copy()
    ret_df = pd.concat([df1, yuzhan_output], axis=0, ignore_index=True)

    # 下游只需要日期粒度，统一输出 YYYY-MM-DD。
    ret_df["预计到仓日期"] = pd.to_datetime(ret_df["预计到仓日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    ret_df["仓库"] = ret_df["仓库"].apply(lambda x: "DE" if x == "EU" else x)

    output_path = INBOUND_DIR / f"在途货件{date_parm}.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        ret_df.to_excel(writer, sheet_name="在途货件-正常", index=False)
        fba_receive_total.to_excel(writer, sheet_name="接收中-汇总", index=False)
        fba_receive_detail.to_excel(writer, sheet_name="接收中-货件详情", index=False)
        future_records_detail.to_excel(writer, sheet_name=f"预计到仓日期早于当天日期的记录", index=False)
        future_records_total.to_excel(writer, sheet_name=f"预计到仓日期早于当天日期的记录-汇总", index=False)
        inspection_excluded_df.to_excel(writer, sheet_name='不计入在途的查验货件详情', index = False)
        inspection_excluded_total.to_excel(writer, sheet_name='不计入在途的查验货件MSKU汇总', index = False)
        yuzhan_false.to_excel(writer, sheet_name=f"预占超过{EXCEPTION_DAYS}天", index=False)
    print(f"已生成: {output_path}")


def main(date_parm=DEFAULT_DATE):
    """脚本主入口。"""
    run_date = pd.to_datetime(date_parm, format="%Y%m%d")
    standard = read_standard_params()

    # 分别构建发货单在途数据和未出库调拨单预占数据。
    fba_inbound, fba_receive_total, fba_receive_detail, inbound_missing, future_records_detail, future_records_total, inspection_excluded_df, inspection_excluded_total = build_inbound_data(date_parm, run_date, standard)
    yuzhan_output, yuzhan_false, yuzhan_missing = build_yuzhan_data(date_parm, run_date, standard)

    # 先处理缺参；只有参数完整时才生成最终结果。
    write_missing_and_raise([inbound_missing, yuzhan_missing])
    write_output(date_parm, fba_inbound, fba_receive_total, fba_receive_detail, yuzhan_output, yuzhan_false, future_records_detail, future_records_total, inspection_excluded_df, inspection_excluded_total)

if __name__ == "__main__":
    # 支持命令行传入处理日期，例如：python code/inbound_process_dynamic.py 20260518
    arg_date = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DATE
    main(arg_date)