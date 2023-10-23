import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

from test_and_backtest import backtest


if __name__ == '__main__':

    '''   
    设置手续费，策略频率，行业类型,回测区间
    frequency频率（'D'：日频，'W'：周频，'M'：月频）
    industry_type行业类型（1：中信一级，2：中信二级，3：申万一级，4：申万二级）
    '''

    # 测试样例
    file_name = '中信一级行业 CON_NP'
    sheet = '周频'
    frequency = 'W'
    industry_type = '1'
    fee = 0.001
    start_date = '20120104'
    end_date = '20221125'

    weight = pd.read_excel(r'.\input\\'+file_name+'.xlsx', sheet_name=sheet, index_col=0)
    backtest(weight,industry_type,start_date,end_date,frequency,fee,file_name,sheet)
