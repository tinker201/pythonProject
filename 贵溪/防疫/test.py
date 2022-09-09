import os
import sys
import time

import pandas as pd

if __name__ == '__main__':
    # print(sys.path[0])
    # print(os.getcwd())
    # file_path = sys.path[0]
    file_path = os.getcwd()
    # file_path = '/Users/wengboyu/Desktop/导入文件夹'
    usecols = ['被采样人姓名', '证件号码', '证件类型', '手机号', '与本人关系', '采集时间', '采样点名称', '采样员姓名', '采样人手机号', '试管码']

    all_file_path = file_path + os.sep + 'all'
    add_file_path = file_path + os.sep + 'add'

    all_df = pd.DataFrame()
    add_df = pd.DataFrame()

    for dirpath, dirnames, filenames in os.walk(all_file_path):
        for filename in filenames:
            if filename[0] != r'.' and (filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx'):
                for sheet_name in pd.read_excel(dirpath + os.sep + filename, sheet_name=None).keys():
                    if '总表' in sheet_name:
                        df = pd.read_excel(dirpath + os.sep + filename, sheet_name='总表', header=0, dtype=str, usecols=usecols)
                        break
                else:
                    df = pd.read_excel(dirpath + os.sep + filename, sheet_name=0, header=0, dtype=str, usecols=usecols)
            all_df = pd.concat([all_df, df])

    for dirpath, dirnames, filenames in os.walk(add_file_path):
        for filename in filenames:
            if filename[0] != r'.' and (filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx'):
                for sheet_name in pd.read_excel(dirpath + os.sep + filename, sheet_name=None).keys():
                    if '总表' in sheet_name:
                        df = pd.read_excel(dirpath + os.sep + filename, sheet_name='总表', header=0, dtype=str, usecols=usecols)
                        break
                else:
                    df = pd.read_excel(dirpath + os.sep + filename, sheet_name=0, header=0, dtype=str, usecols=usecols)
            add_df = pd.concat([add_df, df])

    print(all_df)
    print(add_df)

    for i, j in all_df[['证件号码', '与本人关系']].values.tolist():
        print('i:{}'.format(i))
        print('j:{}'.format(j))
        print(add_df[(add_df['证件号码'] == i) & (add_df['与本人关系'])].index)
        add_df = add_df.drop(add_df[(add_df['证件号码'] == i) & (add_df['与本人关系'])].index)

    out_file_name = file_path + os.sep + str(int(time.time())) + r'.xlsx'
    add_df.to_excel(out_file_name, index=False)
    # add_df.query()
    # pd.read_excel()