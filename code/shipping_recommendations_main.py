"""
    主要实现内容：使用三种参数生成发货列表
"""
from tools import *


def main(parm_date, listing_template_path, sales_program_path, inventory_path, shipment_path, days, current_date):
    """
        该项目的主入口用于构建 建议发货列表并输出为Excel文件。

        Parameters:
        - parm_date: 参数日期，用于构建输出文件名
        - listing_template_path: 列表模板路径
        - sales_program_path: 销售计划路径
        - inventory_path: 库存数据路径
        - shipment_path: 船期数据路径
        - days：输入预估参数

        Returns:
        无返回值，将构建好的发货列表输出为Excel文件

     """

    # 1、加载数据

    # 读取Listing预估表，并指定‘月份’列为日期时间类型，
    predict_df = pd.read_excel(listing_template_path, parse_dates=['月份'])
    # 读取销占比程序表
    sales_df = pd.read_excel(sales_program_path,
                             usecols=['站点', '店铺-站点', '品牌', 'Listing', 'MSKU', '积加SKU', '款式销占比', 'SKU销占比',
                                      '规模定位','总发货天数','快递','空运','海运'])
    inventory_df = pd.read_excel(inventory_path,
                                 usecols=['仓库', 'SKU', 'MSKU', '预占总数', '快递_预占', '空运_预占', '海运_预占', 'FBA自提物流_预占', 'FBA可用库存', 'FBA在途', '7天日均', '15天日均', '30天日均', '已下单数量', '本地锁仓数','本地-在途', '已生产未发货'],
                                 sheet_name='库存汇总')
    # inventory_df = pd.read_excel(inventory_path,
    #                              usecols=['仓库', 'SKU', 'MSKU', '预占总数', 'FBA自提物流_预占', '空运_预占', '海运_预占', 'FBA可用库存', 'FBA在途', '7天日均', '15天日均', '30天日均', '已下单数量', '本地锁仓数','本地-在途', '已生产未发货'],
    #                              sheet_name='库存汇总')
    local_df = pd.read_excel(inventory_path, usecols=['SKU', 'CA', 'DE', 'JP', 'UK', 'US', 'TK本地仓', '共享'], sheet_name='本地仓库存透视')
    sea_df = pd.read_excel(inventory_path, sheet_name='海外仓库存')
    shipment_df = pd.read_excel(shipment_path, sheet_name='在途货件-正常', parse_dates=['预计到仓日期'])   # 这里的在途货件有三种状态：已出运、未出库（预占数据：未从本地出库）、提货中

    # 2、构建建议发货列表初始参数
    send_goods = pd.read_excel(sales_program_path,
                               usecols=['站点', '负责人', 'Listing', '款式', 'MSKU', '积加SKU', '规模定位', '需求定位',
                                        '发货定位', '备货定位', '店铺-站点', 'FNSKU','总发货天数', '快递在途天数', '空运在途天数', '海运在途天数', '快递打包时间', '空运打包时间', '海运打包时间',
                                        '仓储类型'])
    # 将 send_goods 数据框中 MSKU、积加SKU 和 店铺-站点 这三列的值按行相加，并将结果存储到新列 ID 中; sum(axis=1) 默认是对数值型数据进行求和操作
    send_goods['ID'] = send_goods[['MSKU', '积加SKU', '店铺-站点']].sum(axis=1)

    merge_df, shipment_ret, sku_inv = process_datas(predict_df, sales_df, inventory_df, shipment_df)
    # merge_df.to_excel(f'merge_df-{days}.xlsx', index = False)
    # print(merge_df.columns)


    stockout_dict = {'ID': [], 'FBA在库可售天数': [], 'FBA总可售天数': [], '总库存可售天数': [], '首次断货前可售天数': [], 'FBA库存': [], '断货风险总天数': [],
                     '断货总损失销量': [],
                     '断货详情': []}
    # 目标天数
    target_days_dict = {'短尾': 120, '中尾': 120, '长尾': 105, '迭代': 105}

    # 每次处理1个MSKU
    # 记录每个ID组的日均信息
    # 创建一个记录日均销量的空字典
    daily_sales_dict = {'预估日均': []}
    shipping_thresholds = {'dhl_start_days': [], 'dhl_deadline': [], 'airport_deadline': []}
    for item_id in merge_df['ID'].unique():
        # item_id = 'SP001-406 Black 42SP001-406 Black 42SEEKWAY:US'
        # print(item_id)

        # temp_df 是 merge_df 中当前 ID 组的数据。
        temp_df = merge_df[(merge_df['ID'] == item_id)]

        # 结果示例：(temp_shipment 是 shipment_ret 中当前 ID 组的数据。)
        # 店铺-站点              SKU          MSKU     预计到仓日期  发货量  \
        # 2372  rivbos:US  RB831-BlackGrey  A5-LU62-M9WH 2025-04-03   42
        temp_shipment = shipment_ret[shipment_ret['ID'] == item_id]


        # -----------------------------------------------------{初始化参数}---------------------------------------------------------------------
        # MSKU

        # 提取当前 ID 组的第一个 MSKU 值。 例如：A5-LU62-M9WH
        code_parm = temp_df['MSKU'].iloc[0]

        # 从sales参数表提取各物流阈值（空值使用默认值）
        dhl_start_days_val = int(temp_df['快递'].iloc[0]) if pd.notnull(temp_df['快递'].iloc[0]) else 5
        dhl_deadline_val = int(temp_df['空运'].iloc[0]) if pd.notnull(temp_df['空运'].iloc[0]) else 20
        airport_deadline_val = int(temp_df['海运'].iloc[0]) if pd.notnull(temp_df['海运'].iloc[0]) else 60
        
        shipping_thresholds['dhl_start_days'].append(dhl_start_days_val)
        shipping_thresholds['dhl_deadline'].append(dhl_deadline_val)
        shipping_thresholds['airport_deadline'].append(airport_deadline_val)

        # 规模定位。提取当前 ID 组的第一个 规模定位 值。 例如：短尾
        scale = temp_df['规模定位'].iloc[0]

        # 开始日期
        start_date_parm = current_date.strftime('%Y-%m-%d')

        # 初始化库存；例如：1798
        initial_inventory_parm = temp_df['FBA可用库存'].iloc[0]

        # 设定目标天数。根据 规模定位 从 target_days_dict 中获取对应的目标天数。
        Total_shipping_days = temp_df['总发货天数'].iloc[0]
        
        # 判断 Total_shipping_days 是否不为空
        if pd.notnull(Total_shipping_days): 
            target_days = Total_shipping_days
        else:
            target_days = target_days_dict.get(scale)
        # print('target_days:\n',target_days)
        
  
        # 获取每日销量
        # print(f'item_id：{item_id}', f'target_days：{target_days}', f'temp_df：{temp_df}')

        goal_sales = get_daily_sales_parm(target_days, temp_df)
        

        # 用于计算可售天数。TODO:目标可售天数都是165天，为什么还要写这么多同样的函数赋值给不同的变量？ --- 不管有什么库存目标可售天数都发165天
        daily_sales_parm = get_daily_sales_parm(165, temp_df)
        daily_sales_parm_135 = get_daily_sales_parm(135, temp_df)
        daily_sales_parm_165 = get_daily_sales_parm(165, temp_df)

        # temp_df.to_excel('xxx.xlsx', index=False)

        # 获取每一票 预计到仓日期和发货量
        restock_dates_parm = temp_shipment['预计到仓日期'].dt.strftime('%Y-%m-%d').to_list()
        # 获取每一票 发货量
        restock_quantities_parm = temp_shipment['发货量'].to_list()

        # 获取当前MSKU的FBA总库存（FBA总库存=FBA可用库存+发货量）
        total_inventory_parm = initial_inventory_parm + sum(restock_quantities_parm)

        # 获取本地锁仓数
        local_lock_parm = temp_df['本地锁仓数'].iloc[0]

        # 所有库存（所有库存=FBA可用库存+发货量+本地锁仓数）
        all_inventory_parm = total_inventory_parm + local_lock_parm

        msku_data = {
            'code': code_parm, # MSKU
            'start_date': start_date_parm,  # 开始日期，即当前程序输入的日期
            'initial_inventory': initial_inventory_parm,  # FBA可用库存
            'daily_sales': daily_sales_parm,    # 可售天数165，包含每日销售量的列表
            'restock_dates': restock_dates_parm,  # 预计到仓日期 列表
            'restock_quantities': restock_quantities_parm # 发货量 列表
        }
        # print(msku_data)
        # 格式化参数
        start_date = datetime.datetime.strptime(msku_data['start_date'], '%Y-%m-%d')
        initial_inventory = msku_data['initial_inventory']
        daily_sales = msku_data['daily_sales']
        restock_dates = [datetime.datetime.strptime(date, '%Y-%m-%d') for date in msku_data['restock_dates']]
        restock_quantities = msku_data['restock_quantities']
        # code = code_parm

        # 存储每个ID组的日均信息，后续用于计算补货数量
        daily_sales_dict['预估日均'].append(goal_sales)
        # daily_sales_dict.to_excel('./file/daily_sales_dict-预估日均.xlsx', index = False)

        # 格式化参数
        # start_date = datetime.datetime.strptime(start_date_parm, '%Y-%m-%d') # 开始日期，即当前程序输入的日期
        # initial_inventory = initial_inventory_parm # FBA可用库存
        # daily_sales = daily_sales_parm # 可售天数165，包含每日销售量的列表
        # restock_dates = [datetime.datetime.strptime(date, '%Y-%m-%d') for date in restock_dates_parm]  # 预计到仓日期 列表
        # restock_quantities =restock_quantities_parm # 发货量 列表

        # fba在库可售天数 165
        available_sales_days = calculate_available_sales(initial_inventory, daily_sales)

        # FBA总可售天数  （total_inventory_parm FBA总库存=FBA可用库存+发货量） 165
        available_sales_total = calculate_available_sales(total_inventory_parm, daily_sales_parm_135)

        # 总库存可售天数    （all_inventory_parm 所有库存=FBA可用库存+发货量+本地锁仓数）   165
        available_sales_all = calculate_available_sales(all_inventory_parm, daily_sales_parm_165)

        # print(f'initial_inventory：{initial_inventory}; available_sales_days：{available_sales_days}', f'days：{len(daily_sales)}')
        # print(f'total_inventory_parm：{total_inventory_parm}; available_sales_total：{available_sales_total}', f'days：{len(daily_sales_parm_135)}')
        # print(f'all_inventory_parm：{all_inventory_parm}; available_sales_all：{available_sales_all}', f'days：{len(daily_sales_parm_165)}')
        # -----------------------------------------------------{调用计算库存情况的函数}-------------------------------------------------------------------------------------------
        days_without_stock, lost_sales, stockout_periods, available_sales = calculate_inventory(start_date,
                                                                                                initial_inventory,
                                                                                                goal_sales,
                                                                                                restock_dates,
                                                                                                restock_quantities)

        # 构建断货详情信息：断货时间段、销售量损失和缺货天数
        details = display_stockout_details(stockout_periods=stockout_periods)

        stockout_dict['ID'].append(item_id)
        stockout_dict['FBA在库可售天数'].append(available_sales_days)
        stockout_dict['FBA总可售天数'].append(available_sales_total)
        stockout_dict['总库存可售天数'].append(available_sales_all)
        stockout_dict['首次断货前可售天数'].append(available_sales)
        stockout_dict['FBA库存'].append(initial_inventory)
        stockout_dict['断货风险总天数'].append(days_without_stock)
        stockout_dict['断货总损失销量'].append(process_value(lost_sales))
        stockout_dict['断货详情'].append(details)
        # print(stockout_dict)
        # break

    # 构件Excel格式的断货详情
    stockout_df = process_stockout_dataframe(stockout_dict=stockout_dict)

    # 生成 预估日均 发货建议
    shipment_suggestion = {'dhl_pre': [], 'air_pre': [], 'sea_pre': []}
    # print(stockout_df)

    # daily_sales = pd.DataFrame(daily_sales_dict)
    # daily_sales.to_csv('daily_sales_dict预估日均.csv', sep='\t')

    # 对DataFrame-stockout_df中的每一行进行操作，TODO: 一行为一个MSKU
    for index in stockout_df.index:
        target_dhl_pre, target_airport_pre, target_ship_pre = generate_shipment_suggestion(data=stockout_df,
                                                                                           daily_sales_list=daily_sales_dict['预估日均'],
                                                                                           index=index,
                                                                                           dhl_start_days=shipping_thresholds['dhl_start_days'][index],
                                                                                           dhl_deadline=shipping_thresholds['dhl_deadline'][index],
                                                                                           airport_deadline=shipping_thresholds['airport_deadline'][index])
        # 存储预估日均及时的发货建议
        shipment_suggestion['dhl_pre'].append(target_dhl_pre)
        shipment_suggestion['air_pre'].append(target_airport_pre)
        shipment_suggestion['sea_pre'].append(target_ship_pre)
    stockout_df = pd.concat(objs=[stockout_df, pd.DataFrame(shipment_suggestion)], axis=1)

    # stockout_df.to_excel('计算各物流方式损失的销量结果.xlsx', index = False)

    print('-------------------------------------------------------------------------------------')
    # print(stockout_df.head())
    # --------------------------------------------------------------------------------------------------------------------
    ret_df = pd.merge(left=send_goods, right=stockout_df, on=['ID'], how='left')
    ret_df.drop(columns=['ID', '断货详情'], inplace=True)

    # 匹配sku_inv
    ret_df = pd.merge(left=ret_df, right=sku_inv, left_on=['店铺-站点', '积加SKU', 'MSKU'],
                      right_on=['店铺-站点', 'SKU', 'MSKU'], how='left').drop(columns='SKU')
    # 匹配local_df 本地仓库存
    ret_df = pd.merge(left=ret_df, right=local_df, left_on=['积加SKU'], right_on=['SKU'], how='left').drop(columns='SKU')

    # 匹配sea_df 海外仓库存
    ret_df = pd.merge(left=ret_df, right=sea_df, left_on=['站点', '积加SKU'], right_on=['站点', 'SKU'], how='left').drop(columns='SKU')

    # --------------------------------------------------------------------------------------------------------------------
    ret_df.rename(columns={'未出库': 'FBA预占'}, inplace=True)    # 未出库：预占数据（未从本地出库），在途货件里，状态为“未出库”

    # 匹配近期销量
    ret_df = pd.merge(left=ret_df, right=inventory_df[['仓库', 'MSKU', '7天日均', '15天日均', '30天日均', '已下单数量', '本地锁仓数','本地-在途', '已生产未发货', '预占总数', '快递_预占', '空运_预占', '海运_预占', 'FBA自提物流_预占']], how='left',
                      left_on=['店铺-站点', 'MSKU'], right_on=['仓库', 'MSKU']).drop(columns=['仓库'])
    # ret_df = pd.merge(left=ret_df, right=inventory_df[['仓库', 'MSKU', '7天日均', '15天日均', '30天日均', '已下单数量', '本地锁仓数','本地-在途', '已生产未发货', '预占总数', 'FBA自提物流_预占', '空运_预占', '海运_预占']], how='left',
                    #   left_on=['店铺-站点', 'MSKU'], right_on=['仓库', 'MSKU']).drop(columns=['仓库'])

    # 匹配销占比，解决因近期销占比为0导致不发货、不下单等问题
    ret_df = pd.merge(left=ret_df, right=sales_df[['店铺-站点', 'MSKU', 'SKU销占比']], on=['店铺-站点', 'MSKU'], how='left')

    print(f"{ret_df.duplicated(['MSKU', '店铺-站点']).sum()}")

    # 修复因销占比为0导致部分MSKU在库天数异常问题：
    try:
        # 显式指定 inplace=True 以直接修改原始 DataFrame
        ret_df[['FBA库存', '已出运', '本地锁仓数']] = ret_df[['FBA库存', '已出运', '本地锁仓数']].fillna(0)
    except KeyError as e:
        print(f"错误：DataFrame中缺少列 {e}")
    except Exception as e:
        print(f"发生未知错误: {e}")

    # print(ret_df.columns)
    ret_df.loc[ret_df['FBA库存'] == 0, 'FBA在库可售天数'] = 0
    ret_df.loc[(ret_df['FBA库存'] == 0) & (ret_df['已出运'] == 0), 'FBA总可售天数'] = 0
    ret_df.loc[(ret_df['FBA库存'] == 0) & (ret_df['已出运'] == 0) & (ret_df['本地锁仓数'] == 0), '总库存可售天数'] = 0

    # TODO
    ret_df['预估在库日均'] = (ret_df['FBA库存'] / ret_df['FBA在库可售天数']).fillna(0).astype(dtype=float).round(2)
    ret_df['本地库存可售天数'] = ret_df['总库存可售天数'] - ret_df['FBA总可售天数']

    if 'FBA预占' not in ret_df.columns:
        ret_df['FBA预占'] = 0
    ret_df['FBA在途'] = ret_df['已出运'] + ret_df['FBA预占']
    ret_df['补货'] = ret_df['断货总损失销量'].apply(lambda x: '是' if x > 0 else '否')
    ret_df.loc[ret_df['发货定位'] == '不发货', '补货'] = '否'
    ret_df.loc[(ret_df['发货定位'] == '发货') & (ret_df['SKU销占比']) == 0, '补货'] = '该SKU销占比为0'

    # ----------------------------------------------分配货物-------------------------------------------------
    # 定义发货渠道优先级顺序
    shipping_priority = ['dhl', 'air', 'sea']
    sent_googds = {'dhl': [], 'air': [], 'sea': []}
    for index in range(0, ret_df.shape[0]):
        # print(f'--------------------------------------{index}-----------------------------------------------')
        msku_series = ret_df.iloc[index, :]

        # 设置目标仓库
        target_warehouse = msku_series['站点']

        # 初始化仓库库存
        warehouses = {
            f'{target_warehouse}': msku_series[f'{target_warehouse}'],  # US仓库现有库存
            # '抽质检数': msku_series['本地-在途'],  # CA仓库现有库存
        }

        # 初始化发货渠道需求
        shipping_channels = {
            'dhl': msku_series['dhl_pre'],  # dhl渠道需要发货数量
            'air': msku_series['air_pre'],  # air渠道需要发货数量
            'sea': msku_series['sea_pre']  # sea渠道需要发货数量
        }

        # 分配货物
        allocation, remaining_warehouses, shortfall = allocate_goods(warehouses, shipping_channels, shipping_priority,
                                                                     target_warehouse)

        # 打印结果
        # print("货物分配结果:")
        for channel, alloc in allocation.items():
            # print(f"{channel} 渠道分配:")
            detail_list = ''
            for warehouse, amount in alloc.items():
                if warehouse != target_warehouse:
                    # print(f"  从 {warehouse} 仓库发货: {amount} 个")
                    detail_list += f"{warehouse}:{amount}  "

            if shortfall[channel] > 0:
                # print(f"  {channel} 渠道还缺: {shortfall[channel]} 个货物")
                detail_list += f"缺货数:{shortfall[channel]}"

            sent_googds[channel].append(detail_list)

        # break

    ret_df = pd.concat(objs=[ret_df, pd.DataFrame(sent_googds)], axis=1)
    print(ret_df.shape)

    # 正则表达式模式
    pattern_qc = r'抽质检数:(\d+\.?\d*)'
    pattern_stock = r'缺货数:(\d+\.?\d*)'

    for channel in ['dhl', 'air', 'sea']:
        # 提取数据并创建新的列
        ret_df[f'{channel}_pre抽质检数'] = ret_df[channel].apply(lambda x: extract_numbers(x, pattern_qc))
        ret_df[f'{channel}_pre缺货数'] = ret_df[channel].apply(lambda x: extract_numbers(x, pattern_stock))
        ret_df[f'{channel}_pre实际可发货数'] = ret_df[f'{channel}_pre'] - ret_df[f'{channel}_pre缺货数']

    set_warehouse = {'US', 'CA', 'DE', 'JP', 'UK','TK本地仓'}
    ret_df['借调'] = '否'
    ret_df['总发货数'] = ret_df[['dhl_pre', 'air_pre', 'sea_pre']].sum(axis=1)
    for warehouse in set_warehouse:
        ret_df['temp'] = ret_df[list(set_warehouse - {warehouse})].sum(axis=1)
        ret_df.loc[(ret_df['temp'] > 0) & (ret_df[warehouse] < ret_df['总发货数']) & (0 < ret_df['总发货数']) & (ret_df['站点'] == warehouse), '借调'] = '是'
        print(f'--------------------------------{warehouse}-----------------------')
        print(ret_df.query('MSKU == "004-Wayfarer Black"')[['借调', 'temp', '总发货数', '站点', 'US']])

    ret_df.loc[ret_df['发货定位'] == '不发货', '借调'] = '否'

    queku_columns = list(ret_df.columns[ret_df.columns.get_indexer(['断货时间1'])[0]:ret_df.columns.get_indexer(['dhl_pre'])[0]])
    # print(queku_columns)

    ret_columns = ['店铺-站点', '站点', '负责人', 'Listing', '款式', 'MSKU', '积加SKU', 'FNSKU', '仓储类型',
     '规模定位', '需求定位', '发货定位', '备货定位',
     '总发货天数', '快递在途天数', '空运在途天数', '海运在途天数', '快递打包时间', '空运打包时间', '海运打包时间',
     '补货', '借调',
     'FBA在库可售天数', 'FBA总可售天数', '本地库存可售天数', '总库存可售天数', '预估在库日均',
     '断货风险总天数', '断货总损失销量','首次断货前可售天数', 
     'FBA库存', '已出运', 'FBA预占',  'FBA在途',
     'dhl_pre', 'air_pre', 'sea_pre',
     'dhl_pre实际可发货数', 'air_pre实际可发货数', 'sea_pre实际可发货数',
     'dhl_pre缺货数', 'air_pre缺货数', 'sea_pre缺货数',
     # 'dhl_pre抽质检数', 'air_pre抽质检数', 'sea_pre抽质检数',
     '预占总数', '快递_预占', '空运_预占', '海运_预占', 'FBA自提物流_预占',
    #  '预占总数', 'FBA自提物流_预占', '空运_预占', '海运_预占', 
     'CA', 'DE', 'JP', 'UK', 'US', 'TK本地仓', '共享',   
     '本地-在途', '已下单数量', '已生产未发货',
     'CA仓搜海外仓', 'DE商易海外仓', 'DE延讯海外仓', 'JP永翔海外仓', 'UK商易海外仓', 'UK延讯海外仓', 'US商易海外仓', 'US易速达海外仓', '九方欧洲海外仓', '元坤海外仓', 'IT永翔海外仓','CN易速达:易速达美东GA仓',
     '顺丰SF:美国特拉华S2仓', '顺丰SF:美国洛杉矶S5仓', '顺丰SF:美国达拉斯C1仓', '顺丰SF:美国芝加哥C1仓',
     '7天日均', '15天日均', '30天日均']

    ret_columns.extend(queku_columns)
    
    print(set(ret_columns) - set(ret_df.columns), f'---------------{days}------------------')
    print(ret_df.columns)
    ret_df.to_excel(f'E:/sontu/shipping_recommendations/程序建议发货列表/发货列表{parm_date}-new-7.0-120-{days}.xlsx', index=False, columns=ret_columns)

if __name__ == "__main__":
    # 文件路径

    # listing_template_path = 'src_data/Listing预估表-模板.xlsx'
    # sales_program_path = 'src_data/Listing销占比程序表.xlsx'

    # listing_template_path = 'src_data/Listing预估表-30.xlsx'
    # sales_program_path = 'src_data/Listing销占比程序表-30.xlsx'
    #
    # listing_template_path = 'src_data/Listing预估表-15.xlsx'
    # sales_program_path = 'src_data/Listing销占比程序表-15.xlsx'

    # 获取当前日期
    # current_date = datetime.datetime.now()
    year = input('请输入年份：')
    month = input('请输入月份：')
    day = input('请输入日期：')
    parm_date = f'{year}{month}{day}'
    current_date = datetime.datetime(year=int(year), month=int(month), day=int(day))

    inventory_path = f'E:/sontu/shipping_recommendations/src_data/处理后的库存/库存{parm_date}.xlsx'
    shipment_path = f'E:/sontu/shipping_recommendations/src_data/在途货件/在途货件{parm_date}.xlsx'
    for days in ['模板-CA', '模板', f'{parm_date}-15', f'{parm_date}-30', f'{parm_date}-60']:
    # for days in ['模板', f'{parm_date}-15', f'{parm_date}-30', f'{parm_date}-60']:
        listing_template_path = f'E:/sontu/shipping_recommendations/src_data/Listing预估表-{days}.xlsx'
        sales_program_path = f'E:/sontu/shipping_recommendations/src_data/Listing销占比程序表-{days}.xlsx'
        main(parm_date, listing_template_path, sales_program_path, inventory_path, shipment_path, days=days, current_date=current_date)