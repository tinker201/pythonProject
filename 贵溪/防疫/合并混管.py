import cx_Oracle
import pandas as pd
import os

def getdata(dirpath, filename) -> pd.DataFrame:
    # 老方法
    # df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name=None)
    # names = ['被采样人姓名', '证件号码', '证件类型', '手机号', '与本人关系', '采集时间', '采样点名称', '采样员姓名', '采样人手机号', '试管码']
    # # 判断是否需要使用第二张表
    # if len(df_temp.keys()) == 3:
    #     usecols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    #     df = pd.read_excel(dirpath + r'/' + filename, sheet_name='总表', names=names, usecols=usecols)
    #     return df
    # usecols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    # df = pd.read_excel(dirpath + r'/' + filename, names=names, usecols=usecols)
    # return df
    # names = ['被采样人姓名', '证件号码', '证件类型', '手机号', '与本人关系', '采集时间', '采样点名称', '采样员姓名', '采样人手机号', '试管码']
    df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name = None)
    for i in df_temp.keys():
        if i == '总表':
            df = pd.read_excel(dirpath + r'/' + filename, sheet_name = '总表', header = 0, usecols = ['试管码'])
            # print(df.head())
            return df
    df = pd.read_excel(dirpath + r'/' + filename, sheet_name=0, header=0, usecols = ['试管码'])
    print(df['试管码'].dropna().unique().tolist())
    # print(df.head())
    return df

def importOracle(l):

    cx_Oracle.init_oracle_client(lib_dir=r"/Applications/instantclient_19_8")
    connection = cx_Oracle.connect("sjj/123@10.211.55.3:1521/orcl")
    cursor = connection.cursor()
    cursor.executemany("insert into 阳管(采集管号, 收到时间, 备注) values (:1 , sysdate, :2)", l)
    connection.commit()
    connection.close()

    # todo 增加删除重复的管号
    # cursor.execute("delete from 阳管temp where 备注 is")

    # data = cursor.fetchone()
    # print('------------ {} ------------'.format('已存数据'))

    # for i in l:




if __name__ == '__main__':

    # filePath = '/Users/wengboyu/Desktop/混管阳性名单/14：10 新增混阳'
    filePath = '/Users/wengboyu/Desktop/导入文件夹/8月30日原始数据'
    # 重要 一定要填
    cause = '发送人：蔡蕾 文件名：8月30日原始数据'
    df_out = pd.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[0:1].lower() != '.' and (filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx'):
                print(filename)
                df_out = pd.concat([df_out, getdata(dirpath, filename)])
                # if filename

    l = df_out['试管码'].dropna().unique().tolist()
    # print(l)
    l = [[str(int(a)), cause] for a in l]

    print(l)
    # print(len(l))

    # df_out.to_sql()

    importOracle(l)

    print('------------ 已存{}条数据 ------------'.format(len(l)))

