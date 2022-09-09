import pandas as pd
import os
import re

def getdata(dirpath, filename) -> pd.DataFrame:
    df = pd.read_excel(dirpath + r'/' + filename, sheet_name = 'Z01 收入支出决算总表(财决01表)', header = None)
    ptn = re.compile('编制单位：')
    m = {}
    dw_name = re.sub(ptn, '', df[0][2])
    m['单位名称'] = dw_name
    m['年初支出预算数总计'] = df[12][32]
    m['年初支出预算数总计（去除上年结转和结余）'] = df[12][32] - df[2][34]
    m['决算数总计'] = df[14][32]
    m['决算数总计（去除上年结转和结余）'] = df[14][32] - df[2][34]
    m['决算与预算之差'] = df[14][32] - df[12][32]
    data = []
    data.append(m)
    dfout = pd.DataFrame(data)
    return dfout


if __name__ == '__main__':

    filePath = '/Users/wengboyu/文档/部门决算'
    df_out = pd.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-3:].lower() == 'xlsx':
                df_out = pd.concat([df_out, getdata(dirpath, filename)])

    df_out.to_excel('/Users/wengboyu/Downloads/贵溪决算与预算之差.xlsx', index=None)
    print('------ e n d ------')
