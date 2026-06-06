# Shipping Recommendations

基于 Excel 数据的 FBA 发货建议项目。项目会综合销售预估、销占比、库存、在途货件、未出库预占和查验货件信息，计算各 MSKU 的断货风险，并输出建议发货列表。

## 项目流程

整体流程分为四步：

1. 参数处理：整理销售预估参数和销占比参数。
2. 库存处理：合并 FBA 库存、产品库存、预占库存等数据。
3. 在途处理：计算在途货件、接收中货件、未出库预占货件的预计到仓日期。
4. 发货建议：根据库存、在途、销量预估和断货风险生成建议发货列表。

## 目录结构

```text
.
├── code/
│   ├── parms_process.ipynb              # 生成 Listing 预估表、销占比程序表
│   ├── inv_process.ipynb                # 生成处理后的库存表
│   ├── inbound_process.ipynb            # 生成在途货件表
│   ├── shipping_recommendations_main.py # 发货建议主程序
│   └── tools.py                         # 库存模拟、断货计算、发货分配等工具函数
├── src_data/                            # 输入数据和中间结果
│   ├── 销售预估参数/
│   ├── 销占比参数/
│   ├── 产品库存/
│   ├── FBA库存/
│   ├── 预占单/
│   ├── 在途货件/
│   └── 查验货件登记表.xlsx
└── 程序建议发货列表/                    # 输出的建议发货列表
```

## 运行环境

推荐使用 Python 3.11。

主要依赖：

```bash
pip install pandas numpy openpyxl jupyter
```

如果只运行 `shipping_recommendations_main.py`，核心依赖是 `pandas`、`numpy`、`openpyxl`。

## 数据准备

运行前需要准备以下 Excel 数据：

- `src_data/销售预估参数/`：销售预估参数表。
- `src_data/销占比参数/`：Listing、款式、SKU 销占比参数表。
- `src_data/产品库存/`：产品库存导出表。
- `src_data/FBA库存/`：FBA 库存导出表。
- `src_data/预占单/`：未出库调拨单导出表。
- `src_data/在途货件/`：发货单清单、物流追踪更新、标准时效参数表。
- `src_data/查验货件登记表.xlsx`：未放行查验货件登记表。

文件名通常带日期，例如：

```text
库存20260518.xlsx
发货单清单20260518.xlsx
未出库调拨单导出20260518.xlsx
物流追踪更新20260518.xlsx
```

## 运行方式

建议按下面顺序运行。

### 1. 生成参数表

打开并运行：

```text
code/parms_process.ipynb
```

输出：

```text
src_data/Listing预估表-模板.xlsx
src_data/Listing销占比程序表-模板.xlsx
```

### 2. 生成库存表

打开并运行：

```text
code/inv_process.ipynb
```

输出：

```text
src_data/处理后的库存/库存{日期}.xlsx
```

### 3. 生成在途货件表

打开并运行：

```text
code/inbound_process.ipynb
```

输出：

```text
src_data/在途货件/在途货件{日期}.xlsx
```

在途货件处理逻辑：

- 发货单清单是主表。
- 有物流追踪时，优先使用实际物流节点计算时效。
- 缺少实际节点时，使用 `标准时效参数表.xlsx` 补齐三段时效。
- 标准时效缺失时，会输出 `物流时效缺失详情.xlsx`，补齐参数后重跑。
- 查验货件中 `是否放行=否` 且 `是否纳入库存计算=否` 的货件不计入在途库存。
- 查验货件中 `是否放行=否` 且 `是否纳入库存计算=是` 的货件，预计到仓日期取 `纳入库存计算后签收日期`。

### 4. 生成建议发货列表

运行：

```bash
python code/shipping_recommendations_main.py
```

程序会提示输入年份、月份、日期，并读取对应日期的数据文件。

输出：

```text
程序建议发货列表/发货列表{日期}-new-7.0-120-{方案}.xlsx
```

注意：当前主程序里有部分本地绝对路径，例如 `E:/sontu/shipping_recommendations/...`。如果换机器运行，需要先把这些路径改成当前项目路径。

## 输出说明

建议发货列表包含：

- 当前库存和在途库存。
- FBA 在库可售天数、FBA 总可售天数、总库存可售天数。
- 断货风险天数和断货损失销量。
- 快递、空运、海运建议发货量。
- 缺货数、借调判断、预占数量等辅助决策字段。

## GitHub 上传注意

本项目包含大量业务 Excel 数据，可能涉及库存、销量、物流渠道和货件信息。上传 GitHub 前建议：

- 不上传真实业务数据。
- 将 `src_data/`、`程序建议发货列表/`、`code/__pycache__/`、日志文件加入 `.gitignore`。
- 如需保留示例，只保留脱敏后的模板文件。
- 检查代码中的本地绝对路径，避免暴露个人目录或公司目录。

## 当前状态

项目当前以 notebook 处理数据为主，主程序 `shipping_recommendations_main.py` 负责生成最终建议发货列表。适合先在本地按日期跑通完整流程，再逐步整理成更标准的命令行脚本。
