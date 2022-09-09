import pandas as pd
import os
import re

def getdata(dirpath, filename) -> pd.DataFrame:
    names = ['科目编码', '科目名称', '合计', '工资福利支出', '商品和服务支出', '印刷费', '差旅费', '维修（护）费', '会议费', '培训费', '公务接待费', '公务用车运行维护费',
             '其他交通费用', '对个人和家庭的补助', '债务利息及费用支出', '资本性支出（基本建设）', '资本性支出', '对企业补助（基本建设）', '对企业补助', '对社会保障基金补助', '其他支出']
    usecols = [0, 3, 4, 5, 19, 21, 29, 30, 33, 34, 35, 43, 44, 47, 60, 65, 78, 95, 98, 104, 108]
    df = pd.read_excel(dirpath + r'/' + filename, sheet_name='Z05 支出决算明细表(财决05表)', header=[8], names=names, usecols=usecols)
    ptn = re.compile('.xls[x]*', re.IGNORECASE)
    dw_name = re.sub(ptn, '', filename)
    df.insert(loc = 0, column = '单位名称', value = dw_name)
    df = df[df['科目名称'].notnull()]
    return df


if __name__ == '__main__':

    filePath = '/Users/wengboyu/文档/部门决算'
    df_out = pd.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-3:].lower() == 'xlsx':
                df_out = pd.concat([df_out, getdata(dirpath, filename)])

    df_out.to_excel('/Users/wengboyu/Downloads/贵溪决算数据汇总.xlsx', index=None)
    print('------ e n d ------')

