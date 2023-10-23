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
    frequency（'D': Daily，'W': Weekly，'M': Monthly）
    industry_type行业类型（1：CITIC LVL1，2：CITIC LVL2，3：SWHY LVL1，4：SWHY LVL2）
    '''

    # Example
    file_name = '中信一级行业 CON_NP'
    sheet = '周频'
    frequency = 'W'
    industry_type = '1'
    fee = 0.001
    start_date = '20120104'
    end_date = '20221125'

    weight = pd.read_excel(r'.\input\\'+file_name+'.xlsx', sheet_name=sheet, index_col=0)
    backtest(weight,industry_type,start_date,end_date,frequency,fee,file_name,sheet)
