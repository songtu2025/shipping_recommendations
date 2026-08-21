# 处理JP在途数据
'''
程序处理逻辑说明：
1. 将已出库状态的头程调拨单：作为JP站在途数据；
2. 本地出库-到达FBA仓 的这段时效：空运需15天，海运需25天。
    即空运的到仓时间=本地出库时间+15，海运的到仓时间=本地出库时间+25；并且根据 实际在途时间矫正到仓时效：【2026发货跟踪表JP】 （这段时效包括：本地出库-到达海外仓-海外仓处理完成-海外仓发出-FBA接收）；
3、JP在途数据可查看发货列表的AF列【已出运】。

'''

import pandas as pd
DEFAULT_DATE = "20260601"
jp_in_trans_total = pd.read_excel(f'D:/shipping_recommendations2.0/src_data/JP在途-本地已出库/调拨单导出{DEFAULT_DATE}-JP在途.xlsx',sheet_name='单据数据',usecols=['调拨单号','实际出库日期'])
jp_in_trans_detail = pd.read_excel(f'D:/shipping_recommendations2.0/src_data/JP在途-本地已出库/调拨单导出{DEFAULT_DATE}-JP在途.xlsx',sheet_name='明细数据')

merge_df = pd.merge(jp_in_trans_detail, jp_in_trans_total, how='left', left_on='调拨单号', right_on='调拨单号')

merge_df = merge_df[['平台站点', 'SKU', 'MSKU','调出数量', '调拨单号','物流方式', '实际出库日期']]
merge_df['状态'] = '已出运'
merge_df['ShipmentId'] = 'JP头程出库无ShipmentId'

# 新增‘预计到仓日期’列，如果物流方式为‘空运’，则预计到仓日期为实际出库日期+15天，如果物流方式为‘海运’，则预计到仓日期为实际出库日期+25天
# 先将 ‘实际出库日期’列转换为日期类型
merge_df['实际出库日期'] = pd.to_datetime(merge_df['实际出库日期'])
merge_df['预计到仓日期'] = merge_df.apply(lambda x: x['实际出库日期'] + pd.Timedelta(days=15) if x['物流方式'] == '空运' else x['实际出库日期'] + pd.Timedelta(days=25), axis=1)

# 新增仓库列，去掉‘AMAZON:’字符串，并在字符串尾部添加‘_FBA’字符串
merge_df['仓库'] = merge_df['平台站点'].str.replace('AMAZON:', '') + '_FBA'
# 重命名‘调出数量’，‘实际出库日期’，'调拨单号'列名
merge_df.rename(columns={'调出数量': '发货量', '实际出库日期': '预计出运日期', '调拨单号': '发货单号'}, inplace=True)

merge_df['发货单号'] = '调拨单号：' + merge_df['发货单号'].astype(str)

# 删除不需要的列
merge_df.drop(columns=['平台站点'], inplace=True)


result_df = merge_df[['仓库', 'SKU', 'MSKU', '发货量', '预计到仓日期', '预计出运日期', '状态', '物流方式', 'ShipmentId', '发货单号']]

# 读取 在途货件表
in_transit_shipments = pd.read_excel(f'./src_data/在途货件/在途货件{DEFAULT_DATE}.xlsx', sheet_name='在途货件-正常')

# 将merge_df内容放入在途货件表中，这两个表列名完全一致
df = pd.concat([in_transit_shipments, result_df], ignore_index=True)

# 将 在途货件{DEFAULT_DATE}.xlsx的子表'在途货件-正常'内容替换为 df
with pd.ExcelWriter(f'./src_data/在途货件/在途货件{DEFAULT_DATE}.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='在途货件-正常', index=False)



