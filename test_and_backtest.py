import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
import pandas as pd
import numpy as np
from datetime import datetime
import math
import warnings
warnings.filterwarnings('ignore')

def backtest(weight, industry_type, start_date, end_date, frequency, fee, file_name, sheet):

    start_date = datetime.strptime(start_date, '%Y%m%d')
    end_date = datetime.strptime(end_date, '%Y%m%d')

    #weight.set_index(weight.columns[0], True)
    weight.index = pd.to_datetime(weight.index)
    weight = weight.loc[start_date:end_date,:]

    if industry_type == '1':
        price = pd.read_excel('.\\input\\industry_day_price.xlsx', '中信一级', index_col=0)
    if industry_type == '2':
        price = pd.read_excel('.\\input\\industry_day_price.xlsx', '中信二级', index_col=0)
    if industry_type == '3':
        price = pd.read_excel('.\\input\\industry_day_price.xlsx', '申万一级', index_col=0)
    if industry_type == '4':
        price = pd.read_excel('.\\input\\industry_day_price.xlsx', '申万二级', index_col=0)

    # price.set_index(price.columns[0])
    price = price.loc[price.index.isin(weight.index)]
    price = price[weight.columns]

    if weight.shape != price.shape:
        print('Wrong,Weight and Prices should be same length!')


    # T总期数，KW候选标的个数
    T = price.shape[0]
    KW = price.shape[1]

    # 最新期的被移动到第一行，设置首期无信号
    weight = weight.shift(1)
    weight.iloc[0, :] = 0
    ret = price / price.shift(1) - 1
    ret.iloc[0, :] = 0
    strategy_ret = [0]
    bench_ret = [0]


    for i in range(1, len(ret)):
        bench_ret_temp = ret.iloc[i].mean()
        bench_ret.append(bench_ret_temp)
        strategy_ret_temp = ret.iloc[i] * weight.iloc[i]
        strategy_ret_temp.fillna(0)
        strategy_ret_temp = strategy_ret_temp.sum()
        strategy_ret.append(strategy_ret_temp)
    bench_ret = pd.Series(bench_ret, ret.index)
    strategy_ret = pd.Series(strategy_ret, ret.index)
    if fee == 0:
        strategy_net_value = np.cumprod(strategy_ret + 1)
    else:
        change = np.abs(weight - weight.shift(-1).fillna(0)).sum(1)
        change.iloc[0] = 1
        strategy_net_value = np.cumprod((1 - fee * change) * (strategy_ret + 1))
        strategy_ret = strategy_net_value / strategy_net_value.shift(1) - 1
        strategy_ret.iloc[0] = 0
    bench_net_value = np.cumprod(bench_ret + 1)
    excess_ret = strategy_ret - bench_ret
    relative_net_value = strategy_net_value / bench_net_value
    net_values = pd.DataFrame({
        '策略净值': strategy_net_value,
        '基准净值': bench_net_value,
        '相对净值': relative_net_value }, strategy_net_value.index)
    writer = pd.ExcelWriter('output\\' + file_name + sheet + '策略回测表现.xlsx')
    pd.DataFrame().to_excel(writer, 'Sheet1')
    net_values.to_excel(writer, '净值统计')
    
    def statis(ret, frequency):
        if frequency == 'M':
            N = 12
        elif frequency == 'W':
            N = 52
        elif frequency == 'D':
            N = 252


        ET = len(ret)
        net_value = np.cumprod(ret + 1)
        interval_ret = 100 * (net_value.iloc[-1] - 1)
        interval_std = 100 * np.std(ret)
        annual_ret = 100 * (pow(net_value[-1], N / ET) - 1)
        annual_std = 100 * np.std(ret) * np.sqrt(N)


        if annual_std == 0:
            sharpe = np.nan
        else:
            sharpe = annual_ret / annual_std


        drawdown = pd.Series(range(ET))


        for j in range(1, ET):
            drawdown.iloc[j] = 1 - net_value.iloc[j] / max(net_value[:j])
        if len(drawdown) == 1:
            max_drawdown = 0
        else:
            max_drawdown = 100 * max(drawdown[1:])
        return [
            interval_ret,
            interval_std,
            annual_ret,
            annual_std,
            sharpe,
            max_drawdown]

    (annual_strategy_ret, annual_strategy_std, sharpe, max_drawdown) = statis(strategy_ret, frequency)[2:]
    annual_bench_ret = statis(bench_ret, frequency)[2]
    (annual_excess_ret, annual_excess_std, excess_sharpe, excess_max_drawdown) = statis(excess_ret, frequency)[2:]




    print('年超额 annual_excess_ret | ', annual_excess_ret)
    return annual_excess_ret






    strategy_winper = (sum(strategy_ret.iloc[1:] > bench_ret.iloc[1:]) / (T - 1)) * 100
    turnover_rate = np.zeros([T, 1])
    turnover_rate[0, :] = 1
    for t in range(1, T):
        turnover_rate[(t, 0)] = sum(abs(weight.iloc[t] - weight.iloc[t - 1]))
    turnover_rate = pd.DataFrame(turnover_rate, weight.index, ['TurnOverRate'])
    turnover_rate_monthly = turnover_rate.resample('M').sum()
    turnover_rate_yearly = turnover_rate.resample('A').sum()
    perfmon = pd.DataFrame({
        '年化收益(%)': annual_strategy_ret,
        '基准收益(%)': annual_bench_ret,
        '超额收益(%)': annual_excess_ret,
        '年化波动(%)': annual_strategy_std,
        '年化超额波动(%)': annual_excess_std,
        '最大回撤(%)': max_drawdown,
        '超额最大回撤(%)': excess_max_drawdown,
        '夏普比率': sharpe,
        '信息比率': excess_sharpe,
        '胜率(%)': strategy_winper,
        '换手率(年均)': sum(turnover_rate_monthly['TurnOverRate']) / len(turnover_rate_monthly) / 12 }, [0])
    perfmon.to_excel(writer, '策略收益总体统计', False)
    perfmon_yearly_l = []
    date_list_yearly = turnover_rate.to_period('A').index.unique()
    for i in range(len(date_list_yearly)):
        bench_ret_year = bench_ret.to_period('A')
        bench_ret_year = bench_ret_year[bench_ret_year.index == date_list_yearly[i]]
        strategy_ret_year = strategy_ret.to_period('A')
        strategy_ret_year = strategy_ret_year[strategy_ret_year.index == date_list_yearly[i]]
        excess_ret_year = strategy_ret_year - bench_ret_year
        (interval_strategy_ret, interval_strategy_std, temp1, temp2, sharpe, max_drawdown) = statis(strategy_ret_year, frequency)
        interval_bench_ret = statis(bench_ret_year, frequency)[0]
        winper = (sum((strategy_ret_year > bench_ret_year).values) / len(strategy_ret_year.values)) * 100
        (interval_excess_ret, interval_excess_std, temp1, temp2, excess_sharpe, excess_maxdrawdown) = statis(excess_ret_year, frequency)
        perform_yearly_temp = pd.DataFrame({
            '策略年度收益(%)': interval_strategy_ret,
            '基准年度收益(%)': interval_bench_ret,
            '年度超额收益(%)': interval_excess_ret,
            '年度收益波动(%)': interval_strategy_std,
            '年度超额波动(%)': interval_excess_std,
            '最大回撤(%)': max_drawdown,
            '超额最大回撤(%)': excess_maxdrawdown,
            '夏普比率': sharpe,
            '信息比率': excess_sharpe,
            '胜率': winper,
            '换手率': turnover_rate_yearly['TurnOverRate'].iloc[i] }, [
            date_list_yearly[i]])
        perfmon_yearly_l.append(perform_yearly_temp)
    perfmon_yearly = pd.concat(perfmon_yearly_l)
    perfmon_yearly.to_excel(writer, '策略收益年度统计')
    perfmon_monthly_l = []
    date_list_monthly = turnover_rate.to_period('M').index.unique()
    for i in range(len(date_list_monthly)):
        bench_ret_month = bench_ret.to_period('M')
        bench_ret_month = bench_ret_month[bench_ret_month.index == date_list_monthly[i]]
        strategy_ret_month = strategy_ret.to_period('M')
        strategy_ret_month = strategy_ret_month[strategy_ret_month.index == date_list_monthly[i]]
        excess_ret_month = strategy_ret_month - bench_ret_month
        (interval_strategy_ret, interval_strategy_std, temp1, temp2, sharpe, max_drawdown) = statis(strategy_ret_month, frequency)
        interval_bench_ret = statis(bench_ret_month, frequency)[0]
        winper = (sum((strategy_ret_month > bench_ret_month).values) / len(strategy_ret_month.values)) * 100
        (interval_excess_ret, interval_excess_std, temp1, temp2, excess_sharpe, excess_maxdrawdown) = statis(excess_ret_month, frequency)
        perform_monthly_temp = pd.DataFrame({
            '策略月度收益(%)': interval_strategy_ret,
            '基准月度收益(%)': interval_bench_ret,
            '月度超额收益(%)': interval_excess_ret,
            '月度收益波动(%)': interval_strategy_std,
            '月度超额波动(%)': interval_excess_std,
            '最大回撤(%)': max_drawdown,
            '超额最大回撤(%)': excess_maxdrawdown,
            '夏普比率': sharpe,
            '信息比率': excess_sharpe,
            '胜率': winper,
            '换手率': turnover_rate_monthly['TurnOverRate'].iloc[i] }, [date_list_monthly[i]])
        perfmon_monthly_l.append(perform_monthly_temp)
    perfmon_monthly = pd.concat(perfmon_monthly_l)
    perfmon_monthly.to_excel(writer, '策略收益月度统计')
    position = weight.copy()
    position = position.apply(lambda x: x.apply(lambda y: 1 if y != 0 else 0))
    position_count = position.sum(0)
    industry_win_count = pd.DataFrame(position.index, position.columns)
    industry_excess_win_count = pd.DataFrame(position.index, position.columns)
    for ind in industry_win_count.index:
        for col in industry_win_count.columns:
            if position.loc[(ind, col)] > 0 and ret.loc[(ind, col)] > 0:
                industry_win_count.loc[(ind, col)] = 1
            if position.loc[(ind, col)] > 0 and ret.loc[(ind, col)] > bench_ret.loc[ind]:
                industry_excess_win_count.loc[(ind, col)] = 1
                continue
                industry_win_count = industry_win_count.sum(0)
                industry_excess_win_count = industry_excess_win_count.sum(0)
                industry_winper = industry_win_count.copy()
                industry_excess_winper = industry_excess_win_count.copy()
                for i in range(len(industry_winper)):
                    if position_count.iloc[i] == 0:
                        industry_winper.iloc[i] = np.nan
                        industry_excess_winper.iloc[i] = np.nan
                    else:
                        industry_winper.iloc[i] = industry_win_count.iloc[i] / position_count.iloc[i]
                        industry_excess_winper.iloc[i] = industry_excess_win_count.iloc[i] / position_count.iloc[i]
                industry_perform = pd.DataFrame({
                    '行业': position.columns,
                    '配置次数': position_count,
                    '配置频率': position_count / (T - 1),
                    '正绝对收益次数': industry_win_count,
                    '正相对收益次数': industry_excess_win_count,
                    '正绝对收益胜率': industry_winper,
                    '正相对收益胜率': industry_excess_winper })
                industry_perform.to_excel(writer, '分行业胜率', False)
                industry_count_yearly_l = []
                industry_freq_yearly_l = []
                for i in range(len(date_list_yearly)):
                    position_year = position.to_period('A')
                    position_year = position_year[position_year.index == date_list_yearly[i]]
                    position_year_count = position_year.sum(0).to_frame().T
                    position_year_count.index = [
                        date_list_yearly[i]]
                    position_year_freq = position_year_count / position_year.shape[0]
                    position_year_freq.index = [
                        date_list_yearly[i]]
                    industry_count_yearly_l.append(position_year_count)
                    industry_freq_yearly_l.append(position_year_freq)
                industry_count_yearly = pd.concat(industry_count_yearly_l)
                industry_freq_yearly = pd.concat(industry_freq_yearly_l)
                industry_count_yearly.to_excel(writer, '行业年度配置次数')
                industry_freq_yearly.to_excel(writer, '行业年度配置频率')
                writer.save()
