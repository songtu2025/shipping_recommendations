import re
import pandas as pd
import numpy as np
import calendar
import datetime
from datetime import timedelta

import warnings

warnings.filterwarnings('ignore')
pd.set_option('display.max_columns', None)
      

# 计算库存情况的函数
def calculate_inventory(start_date, initial_inventory, daily_sales, restock_dates, restock_quantities):
    """
    计算库存情况的函数
    Parameters:
    - start_date: datetime.datetime，起始日期
    - initial_inventory: float，初始库存
    - daily_sales: list，每日销售量的列表
    - restock_dates: list，补货日期的列表
    - restock_quantities: list，每次补货数量的列表

    Returns:
    Tuple，包含缺货天数、缺货期间销售量损失、缺货时间段列表、在库可售天数
    """

    current_date = start_date
    inventory = initial_inventory  # 初始库存
    days_without_stock = 0  # 缺货天数
    lost_sales = 0  # 缺货期间销售量损失
    stockout_periods = []  # 缺货时间段列表
    available_sales = len(daily_sales)  # 初始化总目标可售天数
    is_first = True  # 记录是否是第一个缺货时间段

    # 遍历每日销售量的列表，模拟每一天的销售情况。
    for sales_quantity in daily_sales:
        # 检查当前日期是否有补货
        if current_date in restock_dates:  # 如果当前日期 current_date 在 restock_dates 列表中，表示有补货。
            inventory += restock_quantities[restock_dates.index(current_date)]   # 通过 restock_dates.index(current_date) 找到当前日期对应的索引，然后从 restock_quantities 中获取补货数量并加到当前库存 inventory 中。

        # 根据每日销售更新库存
        inventory -= sales_quantity

        # 检查库存是否为负数（缺货）
        if inventory <= 0:
            # 如果 stockout_periods 为空或者上一个缺货时间段的结束日期不是前一天，表示这是一个新的缺货时间段
            if not stockout_periods or stockout_periods[-1]['end_date'] != current_date - timedelta(days=1):
                # 断货开始时间距离当前的天数，计算从 start_date 到当前日期的天数 lost_sales_day
                lost_sales_day = (current_date - start_date).days
                stockout_periods.append(
                    {'start_date': current_date, 'end_date': current_date, 'lost_sales': abs(inventory),
                     'lost_sales_day': lost_sales_day, 'duration': 1})
            # 否则，更新最后一个缺货时间段的结束日期、销售量损失和持续时间。
            else: 
                stockout_periods[-1]['end_date'] = current_date
                stockout_periods[-1]['lost_sales'] += abs(inventory)
                stockout_periods[-1]['duration'] += 1

            # 在库可售天数
            if is_first: # 如果是第一个缺货时间段，记录从 start_date 到当前日期的天数作为在库可售天数。
                # print(f'current_date: {current_date},type：{type(current_date)}')
                available_sales = (current_date - start_date).days
                is_first = False
            days_without_stock += 1 # 增加缺货天数
            lost_sales += abs(inventory) # 累加缺货期间的销售量损失
            inventory = 0  # 将库存设置为 0，表示缺货

        # 移动到下一天；将 current_date 增加一天，模拟下一天的销售情况。
        current_date += timedelta(days=1)

    return days_without_stock, lost_sales, stockout_periods, available_sales


def calculate_available_sales(initial_inventory, daily_sales):
    """
    计算在库可售天数

    Parameters:
        initial_inventory (float): 初始库存
        daily_sales (list): 每日销售量列表

    Returns:
        available_sales(int): 在库可售天数
    """
    inventory = initial_inventory  # 初始库存
    available_sales = 0  # 初始化在库可售天数
    for sales_quantity in daily_sales:
        inventory -= sales_quantity
        if inventory >= 0:
            available_sales += 1
        else:
            break
    return available_sales


# def process_datas(predict_df, sales_df, inventory_df, shipment_df):
    """
      处理数据，生成合并后的数据框、发货数据透视表和SKU库存状况表

    Parameters:
    predict_df (DataFrame): 预测数据
    sales_df (DataFrame): 销售数据
    inventory_df (DataFrame): 库存数据
    shipment_df (DataFrame): 发货数据

    Returns:
    merge_df (DataFrame): 合并后的数据框
    shipment_ret (DataFrame): 发货数据透视表
    sku_inv (DataFrame): SKU库存状况表
    """
    # 从仓库中提取 店铺-站点 信息
    shipment_df['店铺-站点'] = shipment_df['仓库'].str.split('_', expand=True)[0]

    # SKU维度的库存状况
    sku_inv = shipment_df.pivot_table(index=['店铺-站点', 'SKU', 'MSKU'], columns=['状态'], values='发货量',
                                      aggfunc=sum, fill_value=0).reset_index()

    # 透视并汇总发货数据
    shipment_ret = shipment_df.pivot_table(index=['店铺-站点', 'SKU', 'MSKU', '预计到仓日期'], values='发货量',
                                           aggfunc=sum, fill_value=0).reset_index()
    shipment_ret['预计到仓日期'] = pd.to_datetime(shipment_ret['预计到仓日期'])
    shipment_ret['ID'] = shipment_ret[['MSKU', 'SKU', '店铺-站点']].sum(axis=1)


    # 计算日期差异
    predict_df['当月总天数'] = predict_df['月份'].apply(lambda x: calendar.monthrange(x.year, x.month)[1])

    #  是一个新列，存储了每一行的 月份 对应的月末日期与当前日期之间的天数差（加 1 后的结果）
    predict_df['当前可用天数'] = (predict_df['月份'] + pd.offsets.MonthEnd(0) - datetime.datetime.now()).dt.days + 1
    predict_df = predict_df.query('当前可用天数 > 0')
    
    # 对 '当前可用天数' 列进行 clip 操作。np.clip 是 NumPy 提供的一个函数，用于将数组中的值限制在指定范围内。
    # 目的是确保 当前可用天数 的值在合理的范围内：
        # 最小值为 0：因为天数不能为负数。
        # 最大值为 当月总天数：因为 当前可用天数 不可能超过当月的总天数。
    predict_df['当前可用天数'] = np.clip(predict_df['当前可用天数'], a_min=0, a_max=predict_df['当月总天数'])

    # 合并数据框
    merge_df = pd.merge(left=sales_df,
                        right=predict_df[['站点', 'Listing', '当月总天数', '当前可用天数', 'Listing_月度预估销量']],
                        on=['站点', 'Listing'], how='left')

    # 计算MSKU每日均值
    merge_df['MSKU日均'] = (merge_df['Listing_月度预估销量'] / merge_df['当月总天数']) * merge_df['款式销占比'] * \
                           merge_df['SKU销占比']

    # 与库存数据合并
    merge_df = pd.merge(left=merge_df, right=inventory_df, left_on=['店铺-站点', '积加SKU', 'MSKU'],
                        right_on=['仓库', 'SKU', 'MSKU'], how='left')

    # 删除不必要的列
    merge_df.drop(columns=['仓库', 'SKU', '当月总天数'], inplace=True)

    # 计算ID列
    merge_df['ID'] = merge_df[['MSKU', '积加SKU', '店铺-站点']].sum(axis=1)

    return merge_df, shipment_ret, sku_inv
def process_datas(predict_df, sales_df, inventory_df, shipment_df):
    """
      处理数据，生成合并后的数据框、发货数据透视表和SKU库存状况表

    Parameters:
    predict_df (DataFrame): 预测数据
    sales_df (DataFrame): 销售数据
    inventory_df (DataFrame): 库存数据
    shipment_df (DataFrame): 发货数据

    Returns:
    merge_df (DataFrame): 合并后的数据框
    shipment_ret (DataFrame): 发货数据透视表
    sku_inv (DataFrame): SKU库存状况表
    """
    # 从仓库中提取 店铺-站点 信息
    shipment_df['店铺-站点'] = shipment_df['仓库'].str.split('_', expand=True)[0]

    # SKU维度的库存状况。这里的状态有：已出运、未出库、提货中
    sku_inv = shipment_df.pivot_table(index=['店铺-站点', 'SKU', 'MSKU'], columns=['状态'], values='发货量',
                                      aggfunc=sum, fill_value=0).reset_index()

    # 透视并汇总发货数据
    shipment_ret = shipment_df.pivot_table(index=['店铺-站点', 'SKU', 'MSKU', '预计到仓日期'], values='发货量',
                                           aggfunc=sum, fill_value=0).reset_index()
    shipment_ret['预计到仓日期'] = pd.to_datetime(shipment_ret['预计到仓日期'])
    shipment_ret['ID'] = shipment_ret[['MSKU', 'SKU', '店铺-站点']].sum(axis=1)

    # 计算日期差异
    predict_df['当月总天数'] = predict_df['月份'].apply(lambda x: calendar.monthrange(x.year, x.month)[1])  # 获取当月的总天数
    # predict_df['当前可用天数'] = (predict_df['月份'] + pd.offsets.MonthEnd(0) - datetime.datetime.now()).dt.days + 1
    reference_date = datetime.datetime.now() - datetime.timedelta(days=1) # 四天前的日期
    predict_df['当前可用天数'] = (predict_df['月份'] + pd.offsets.MonthEnd(0) - reference_date).dt.days + 1

    predict_df = predict_df.query('当前可用天数 > 0')
    # 对 '当前可用天数' 列进行 clip 操作
    predict_df['当前可用天数'] = np.clip(predict_df['当前可用天数'], a_min=0, a_max=predict_df['当月总天数'])

    # 合并数据框
    merge_df = pd.merge(left=sales_df,
                        right=predict_df[['站点', 'Listing', '当月总天数', '当前可用天数', 'Listing_月度预估销量']],
                        on=['站点', 'Listing'], how='left')

    # 计算MSKU每日均值
    merge_df['MSKU日均'] = (merge_df['Listing_月度预估销量'] / merge_df['当月总天数']) * merge_df['款式销占比'] * \
                           merge_df['SKU销占比']

    # 与库存数据合并
    merge_df = pd.merge(left=merge_df, right=inventory_df, left_on=['店铺-站点', '积加SKU', 'MSKU'],
                        right_on=['仓库', 'SKU', 'MSKU'], how='left')

    # 删除不必要的列
    merge_df.drop(columns=['仓库', 'SKU', '当月总天数'], inplace=True)

    # 计算ID列
    merge_df['ID'] = merge_df[['MSKU', '积加SKU', '店铺-站点']].sum(axis=1)

    return merge_df, shipment_ret, sku_inv

# def get_daily_sales_parm(target_days, temp_df):
    """
    获取每日销售量参数的函数。

    Parameters:
    - target_days: int，目标天数
    - temp_df: DataFrame，包含销售数据的DataFrame

    Returns:
    list，包含每日销售量的列表
    """

    # # 确保没有 NaN 值 todo
    # temp_df['MSKU日均'] = temp_df['MSKU日均'].fillna(0)
    # temp_df['当前可用天数'] = temp_df['当前可用天数'].fillna(0)

    # 提取 MSKU日均 和 当前可用天数 两列，并将它们重置索引（reset_index(drop=True)），确保它们是独立的序列
    sales_avg = temp_df.MSKU日均.reset_index(drop=True)
    curr_days = temp_df.当前可用天数.reset_index(drop=True)
    '''
    # print("get_daily_sales_parm函数中：\tsales_avg=", sales_avg, "\tcurr_days=", curr_days)
    # 以上代码输出的值如下：
    get_daily_sales_parm函数中：    
    sales_avg= 
    0    15.347564
    1    17.378456
    2    15.235300
    3    10.993792
    4     9.045033
    5     8.236815
    6     6.512905
    7     8.888224
    Name: MSKU日均, dtype: float64  
    curr_days= 
    0    25
    1    30
    2    31
    3    31
    4    30
    5    31
    6    30
    7    31
    Name: 当前可用天数, dtype: int64
    '''

    # 使用 np.cumsum 计算 当前可用天数 的累计和。cumulative_sum 是一个数组，表示从第一天开始的累计天数。cumsum函数用于计算数组的累积和（cumulative sum）。它会沿着指定的轴对数组中的元素进行逐项累加，生成一个新的数组，其中每个元素是原始数组中从开始到当前位置的所有元素之和。
    cumulative_sum = np.cumsum(curr_days) 

    # 找到累计和超过目标天数的索引；
    # 使用 np.argmax 找到第一个累计和超过 target_days目标天数 的索引
    index = np.argmax(cumulative_sum > target_days)
    if index == 0:
        index = len(cumulative_sum) - 1
    # print(f'index:{index}')
    # print('------------------------------')
    # print(f'cumulative_sum：{cumulative_sum}')

    # 更新 curr_days 和 sales_avg
    #     如果目标天数小于累计和的最后一个值，调整 curr_days当前可用天数 中的最后一个值，使其恰好达到 target_days目标天数。
    #     截取 curr_days-当前可用天数 和 sales_avg-MSKU日均，只保留到 index + 1 的部分
    if index < len(curr_days) and target_days <= cumulative_sum.iloc[-1]:
        curr_days.iloc[index] = target_days - (cumulative_sum[index] - curr_days.iloc[index])
    curr_days = curr_days.iloc[:index + 1]
    sales_avg = sales_avg.iloc[:index + 1]
    return [item for avg, days in zip(sales_avg, curr_days) for item in [avg] * int(days)]

def get_daily_sales_parm(target_days, temp_df):
    """
    获取每日销售量参数的函数。

    Parameters:
    - target_days: int，目标天数
    - temp_df: DataFrame，包含销售数据的DataFrame

    Returns:
    list，包含每日销售量的列表
    """

    # # 确保没有 NaN 值 todo
    # temp_df['MSKU日均'] = temp_df['MSKU日均'].fillna(0)
    # temp_df['当前可用天数'] = temp_df['当前可用天数'].fillna(0)

    sales_avg = temp_df.MSKU日均.reset_index(drop=True)
    curr_days = temp_df.当前可用天数.reset_index(drop=True)
    # print('curr_days:',curr_days)

    # 找到累计和超过目标天数的索引
    cumulative_sum = np.cumsum(curr_days)
    index = np.argmax(cumulative_sum > target_days)
    # 如果 index 为 0，说明所有累计和都小于 target_days，此时将 index 设置为数组的最后一个索引。
    if index == 0:
        index = len(cumulative_sum) - 1
    # print(f'index:{index}')
    # print('------------------------------')
    # print(f'cumulative_sum：{cumulative_sum}')

    # 更新 curr_days 和 sales_avg
    # 如果目标天数小于累计和的最后一个值，调整 curr_days 中的最后一个值，使其恰好达到 target_days。截取 curr_days 和 sales_avg，只保留到 index + 1 的部分
    if index < len(curr_days) and target_days <= cumulative_sum.iloc[-1]:
        curr_days.iloc[index] = target_days - (cumulative_sum[index] - curr_days.iloc[index])
    curr_days = curr_days.iloc[:index + 1]
    sales_avg = sales_avg.iloc[:index + 1]
    # print('curr_days:\n', curr_days)
    # 使用列表推导式生成每日销售量的列表。对于每一对 avg（日均销售量）和 days（天数），重复 avg 值 days 次。
    return [item for avg, days in zip(sales_avg, curr_days) for item in [avg] * int(days)]

def get_total_sales(target_days, temp_df):
    """
    获取指定目标天数的总销量的函数。

    Parameters:
    - target_days: int，目标天数
    - temp_df: DataFrame，包含销售数据的DataFrame

    Returns:
    int，指定目标天数的总销量
    """

    sales_avg = temp_df['MSKU日均'].reset_index(drop=True)
    curr_days = temp_df['当前可用天数'].reset_index(drop=True)
    # 找到累计和超过目标天数的索引
    cumulative_sum = np.cumsum(curr_days)
    index = np.argmax(cumulative_sum > target_days)
    # print(f'-------------{curr_days}----------------')
    # 更新 curr_days 和 sales_avg
    if index < len(curr_days):
        curr_days.iloc[index] = target_days - (cumulative_sum[index] - curr_days.iloc[index])
    total_sales = sum(sales_avg.iloc[:index + 1] * curr_days.iloc[:index + 1])

    return int(total_sales)


def display_stockout_details(stockout_periods):
    """
    构建断货详情：断货时间段、销售量损失和缺货天数的函数。

    Parameters:
    - stockout_periods: list，包含断货时间段信息的列表
    """
    details = ''

    # 存在在途货件且断货
    if stockout_periods:  #如果 stockout_periods 列表不为空，表示存在断货时间段。
        length = len(stockout_periods)
        for i, period in enumerate(stockout_periods):
            # 断货开始日期、断货结束日期、断货天数、损失销量
            info = f"{period['start_date'].strftime('%Y-%m-%d')}   {period['end_date'].strftime('%Y-%m-%d')}|{period['duration']}|{process_value(period['lost_sales'])}|{int(period['lost_sales_day'])}"
            if (i + 1) != length:
                details = details + info + '|'
            else:
                details = details + info
    else:
        details = '未发生断货'

    return details


def process_stockout_dataframe(stockout_dict):
    """
    处理断货信息字典并生成包含断货详情的DataFrame的函数。

    Parameters:
    - stockout_dict: dict，包含断货信息的字典

    Returns:
    - pd.DataFrame，包含断货详情的DataFrame
    """
    # 创建DataFrame
    stockout_df = pd.DataFrame(data=stockout_dict)

    # 将'断货详情'列的字符串按照 '|' 进行拆分，扩展为多列
    stockout_detail = stockout_df['断货详情'].str.split('|', expand=True)

    # 获取拆分后的 DataFrame 的列名列表
    columns = stockout_detail.columns.to_list()

    # 循环处理每组 4 列的数据
    for i in range(0, len(columns) // 4):
        # 提取当前一组列名
        src_names = columns[i * 4: i * 4 + 4]

        # 重命名列名，添加更具描述性的名字
        new_column_names = [f'断货时间{i + 1}', f'断货总天数{i + 1}', f'损失销量{i + 1}', f'断货开始天数{i + 1}']
        stockout_detail.rename(columns=dict(zip(src_names, new_column_names)), inplace=True)

        # 将新列中的 NaN 值填充为 0，然后将数据类型转换为整数
        stockout_detail[new_column_names[1:]] = stockout_detail[new_column_names[1:]]  # .fillna(0).astype(dtype=int)
    # 合并DataFrame
    stockout_df = pd.concat(objs=[stockout_df, stockout_detail], axis=1)
    # stockout_df.to_excel("构建断货详情_processStockoutDataframe函数输出.xlsx", index = False)

    return stockout_df


# 生成发货建议
def calculate_sales(daily_sales_list, start_day, end_day, index):
    """
    计算给定时间段内的总销售量。
    """
    # 如果 start_day 大于 end_day，表示时间段无效，返回 0
    if start_day > end_day:
        return 0
    return sum(daily_sales_list[index][start_day:end_day + 1]) # 使用切片操作 daily_sales_list[index][start_day:end_day + 1] 提取指定时间段内的销售量。使用 sum 函数计算该时间段内的总销售量并返回。


def process_data(data, daily_sales_list, index, dhl_start_days, dhl_deadline, airport_deadline):
    """
    处理数据，计算每种运输方式的损失销售量。
    """
    # 获取每一个SKU断货信息。  
        # 使用 iloc 和 get_indexer 获取当前行从 '断货时间1' 列开始的所有数据。
        # data.columns.get_indexer(['断货时间1'])[0] 返回 '断货时间1' 列的索引位置。
        # input_data 是一个 Pandas Series，包含当前行的所有断货信息。
    input_data = data.iloc[index, data.columns.get_indexer(['断货时间1'])[0]:]

    # 初始化物流方式发货量
    target_dhl, target_air_freight, target_sea_freight = 0, 0, 0

    # 每组 4 列表示一个断货时间段的详细信息，包括断货总天数、损失销量和断货开始天数。
    for loop_index in range(0, int(len(input_data) / 4)):
        
        dhl, air_freight, sea_freight = 0, 0, 0
        # 获取原始值。从 input_data 中提取当前断货时间段的断货总天数、损失销量和断货开始天数的原始值
        quekou_total_days_raw, lost_sales_raw, quekou_start_day_raw = input_data[loop_index * 4 + 1], input_data[
            loop_index * 4 + 2], input_data[loop_index * 4 + 3]

        # 检查是否为NaN，并给出默认值或处理逻辑。如果原始值不是 NaN，将其转换为整数；否则，设置为 0
        quekou_total_days = int(quekou_total_days_raw) if not pd.isna(quekou_total_days_raw) else 0  # 断货总天数
        quekou_start_day = int(quekou_start_day_raw) if not pd.isna(quekou_start_day_raw) else 0    # 断货开始天数

        # 校验数据的合理性。如果断货总天数或断货开始天数为负数，跳过当前循环。
        if quekou_total_days < 0 or quekou_start_day < 0:
            continue
        
        # 计算断货结束天数。等于断货开始天数+断货总天数
        quekou_end_day = quekou_start_day + quekou_total_days

        # 确保结束天数不超出界限。如果断货结束天数超出日均销量列表的长度，将其设置为日均销量列表的长度。
        if quekou_end_day > len(daily_sales_list[index]):
            quekou_end_day = len(daily_sales_list[index])

        # 计算不同情况下的销售损失。根据断货开始天数和运输方式的截止日期，计算每种运输方式的损失销售量。
        # 开始断货天数小于DHL期限
        if quekou_start_day <= dhl_deadline:
            # 只发DHl。如果断货结束天数小于或等于 DHL 的截止日期，
            if quekou_end_day <= dhl_deadline:
                # 考虑在库为0时DHL的在途时间。如果断货开始天数小于或等于 DHL 的起始天数，DHL 损失销售量为 0。
                if quekou_start_day <= dhl_start_days:
                    dhl = 0
                else: # 否则，计算 DHL 的损失销售量。
                    dhl = calculate_sales(daily_sales_list, quekou_start_day, quekou_end_day + 1, index)
            # 发DHL和空运。 如果断货结束天数小于或等于空运的截止日期，计算 DHL 和空运的损失销售量。
            elif quekou_end_day <= airport_deadline:
                # 考虑在库为0时DHL的在途时间
                if quekou_start_day <= dhl_start_days:
                    dhl = calculate_sales(daily_sales_list, dhl_start_days + 1, dhl_deadline, index)
                else:
                    dhl = calculate_sales(daily_sales_list, quekou_start_day, dhl_deadline, index)
                air_freight = calculate_sales(daily_sales_list, dhl_deadline + 1, quekou_end_day, index)
            # 发DHL、空运和海运 如果断货结束天数大于空运的截止日期，计算 DHL、空运和海运的损失销售量。
            else:
                # 考虑在库为0时DHL的在途时间
                if quekou_start_day <= dhl_start_days:
                    dhl = calculate_sales(daily_sales_list, dhl_start_days + 1, dhl_deadline, index)
                else:
                    dhl = calculate_sales(daily_sales_list, quekou_start_day, dhl_deadline, index)
                air_freight = calculate_sales(daily_sales_list, dhl_deadline + 1, airport_deadline, index)
                sea_freight = calculate_sales(daily_sales_list, airport_deadline + 1, quekou_end_day, index)
        # 开始断货天数小于空运期限  如果断货开始天数小于或等于空运的截止日期，进一步判断：
        elif quekou_start_day <= airport_deadline:
            # 只发空运  如果断货结束天数小于或等于空运的截止日期，计算空运的损失销售量。
            if quekou_end_day <= airport_deadline:
                air_freight = calculate_sales(daily_sales_list, quekou_start_day, quekou_end_day, index)
            # 发空运和海运  如果断货结束天数大于空运的截止日期，计算空运和海运的损失销售量。
            else:
                air_freight = calculate_sales(daily_sales_list, quekou_start_day, airport_deadline, index)
                sea_freight = calculate_sales(daily_sales_list, airport_deadline + 1, quekou_end_day, index)
        # 如果开始天数不是NaN，但不在DHL或空运期限内，则为海运。    计算海运的损失销售量
        elif not pd.isna(quekou_start_day):
            sea_freight = calculate_sales(daily_sales_list, quekou_start_day, quekou_end_day, index)

        # 累加对应运输方式的损失销售量
        target_dhl += dhl
        target_air_freight += air_freight
        target_sea_freight += sea_freight

    return process_value(target_dhl), process_value(target_air_freight), process_value(target_sea_freight)


def generate_shipment_suggestion(data, daily_sales_list, index, dhl_start_days, dhl_deadline, airport_deadline):
    """
    data: 断货详情，dataframe类型
    daily_sales_list: 日均销量
    index: 当前索引 断货详情的行

    生成运输建议，计算DHL、空运和海运的损失销售量。
    """
    # 确保输入数据有效性
    if not isinstance(data, pd.DataFrame) or not isinstance(daily_sales_list, list) or not isinstance(index, int):
        raise ValueError("Invalid input data types.")

    if index < 0 or index >= data.shape[0]:
        raise ValueError("Index out of bounds.")

    return process_data(data, daily_sales_list, index, dhl_start_days, dhl_deadline, airport_deadline)


# 定义一个函数来生成仓库优先级顺序
def generate_warehouse_priority(target_warehouse, warehouses):
    priority = [target_warehouse]  # 优先从目标仓库发货
    for warehouse in warehouses:
        if warehouse != target_warehouse:
            priority.append(warehouse)  # 将其他仓库按顺序添加到优先级列表中
    return priority


# 分配货物的函数
def allocate_goods(warehouses, shipping_channels, shipping_priority, target_warehouse):
    # 生成仓库优先级顺序
    warehouse_priority = generate_warehouse_priority(target_warehouse, warehouses)

    # 初始化分配结果
    allocation = {channel: {} for channel in shipping_channels}
    shortfall = {channel: 0 for channel in shipping_channels}  # 初始化缺口记录

    # 按发货渠道优先级分配货物
    for channel in shipping_priority:
        remaining = shipping_channels[channel]  # 当前渠道需要发货的剩余数量
        for warehouse in warehouse_priority:
            if remaining <= 0:
                break  # 如果当前渠道的货物已经分配完毕，跳出循环
            available = warehouses[warehouse]  # 当前仓库可用库存
            if available > 0:
                if available >= remaining:
                    allocation[channel][warehouse] = remaining  # 从当前仓库分配所需的全部货物
                    warehouses[warehouse] -= remaining  # 更新当前仓库的剩余库存
                    remaining = 0  # 当前渠道的货物分配完毕
                else:
                    allocation[channel][warehouse] = available  # 从当前仓库分配所有可用货物
                    warehouses[warehouse] = 0  # 当前仓库的货物分配完毕
                    remaining -= available  # 更新当前渠道需要分配的剩余数量

        if remaining > 0:
            shortfall[channel] = remaining  # 如果货物分配不足，记录缺口

    return allocation, warehouses, shortfall  # 返回分配结果、更新后的仓库库存和缺口记录


# 定义提取数字的函数
def extract_numbers(text, pattern):
    if text == '':
        return 0
    match = re.search(pattern, text)
    if match:
        return float(match.group(1))
    return 0


def process_value(value):
    """
    根据传入值判断返回结果：
    - 如果传入的值在 0 和 1 之间，返回 1。
    - 否则，将传入的值转换为整数并返回。
    """
    if 0 < value < 1:
        return 1
    else:
        return int(value)