from typing import List

import cx_Oracle
import pandas as pd
import os
import warnings

from cx_Oracle import DatabaseError


def getImportedFile(connection) -> List:
    cursor = connection.cursor()
    cursor.execute("select 文件名 from 已导入的EXCEL表")
    filenames = [row[0] for row in cursor.fetchall()]
    cursor.close()
    return filenames


def importFilename(connection, filename):
    cursor = connection.cursor()
    cursor.execute("insert into 已导入的EXCEL表(文件名, 导入时间) values (:1 , sysdate)", [filename])
    cursor.close()
    # connection.commit()


def importTestData(connection, dirpath, filename):
    print('--------------------正在读取：{}--------------------'.format(filename))
    df = pd.read_excel(dirpath + r'/' + filename, header=0, parse_dates=['采集时间', '检测时间'], dtype=str)
    print(df.shape[0])
    print('--------------读取完成：{}  共有{}条数据--------------'.format(filename, df.shape[0]))
    df = df.where(df.notna(), None)

    data_list = df.iloc[:, 1:].values.tolist()

    cursor = connection.cursor()
    cursor.execute('truncate table 核酸TEMP')
    cursor.executemany('insert into 核酸TEMP(市, 区, 街道, 社区, 采集地点, 采集管号, 身份证号, 姓名, 性别, 电话, 住址, 年龄, 类别, 备注, 采集时间, 采集人姓名, 采集人电话, 标本类型, 接收实验室, 检测时间, 人员关系, 箱号) values (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17,:18,:19,:20,:21,:22)', data_list)
    connection.commit()
    cursor.execute('merge into 核酸 a using 核酸TEMP b on (a.身份证号 = b.身份证号 and a.人员关系 = b.人员关系 and a.采集时间 = b.采集时间) when matched then update set a.市 = b.市, a.区 = b.区, a.街道 = b.街道, a.社区 = b.社区, a.采集地点 = b.采集地点, a.采集管号 = b.采集管号, a.姓名 = b.姓名, a.性别 = b.性别, a.电话 = b.电话, a.住址 = b.住址, a.年龄 = b.年龄, a.类别 = b.类别, a.备注 = b.备注, a.采集人姓名 = b.采集人姓名, a.采集人电话 = b.采集人电话, a.标本类型 = b.标本类型, a.接收实验室 = b.接收实验室, a.检测时间 = b.检测时间, a.箱号 = b.箱号 when not matched then insert values (b.市, b.区, b.街道, b.社区, b.采集地点, b.采集管号, b.身份证号, b.姓名, b.性别, b.电话, b.住址, b.年龄, b.类别, b.备注, b.采集时间, b.采集人姓名, b.采集人电话, b.标本类型, b.接收实验室, b.检测时间, b.人员关系, b.箱号)')
    connection.commit()
    cursor.close()

    importFilename(connection, filename)
    return df.shape[0]

def importPositiveData(connection, l):

    cursor = connection.cursor()
    cursor.execute('truncate table 阳管temp')
    cursor.executemany("insert into 阳管temp(采集管号, 收到时间, 备注) values (:1 , sysdate, :2)", l)
    connection.commit()
    cursor.execute('merge into 阳管 a using 阳管temp b on (a.采集管号 = b.采集管号) when not matched then insert values (b.采集管号, b.收到时间, b.备注)')
    connection.commit()
    cursor.close()

def readExcel(dirpath, filename):
    print(filename)
    l = []
    df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name=None)
    for i in df_temp.keys():
        if i == '总表':
            df = pd.read_excel(dirpath + r'/' + filename, sheet_name='总表', header=0, dtype=str)
            break
    else:
        df = pd.read_excel(dirpath + r'/' + filename, sheet_name=0, header=0, dtype=str)

    for i in df.columns.values.tolist():
        if '试管' in i:
            return df[i].dropna().unique().tolist()
    else:
        raise Exception('未找到试管列')

def readExcel(dirpath, filename, columns) -> list:
    print(filename)
    df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name=None)
    sheet_name = 0
    column_names = []
    flag = -1

    for i in df_temp.keys():
        if i == '总表':
            sheet_name = '总表'
            break

    for a in range(10):
        df = pd.read_excel(dirpath + r'/' + filename, sheet_name=sheet_name, header=a, dtype=str)
        for j in df.columns.values.tolist():
            if columns[0] in j:
                flag = a
        if flag != -1:
            break

    for i in range(len(columns)):
        for j in df.columns.values.tolist():
            if columns[i] in j:
                column_names.append(j)
                break
        else:
            raise Exception('未找到{}列'.format(columns[i]))

    df = df.where(df.notnull(), None)

    return df[column_names].dropna(how='all').values.tolist()

def importIsolationPerson(connection, l):

    # l = l[0:3]
    # print(l)
    # for i in l:
    #     for j in i:
    #         if type(j) != type(str(1)):
    #             print(j)
    #             print(type(j))


    cursor = connection.cursor()
    cursor.executemany("insert into 隔离点人员(身份证号码, 隔离人姓名, 备注) values (:1, :2，:3)", l)
    cursor.close()


def readIsolationPerson(dirpath, filename):
    print(filename)
    df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name=None)
    for i in df_temp.keys():
        if i == '总表':
            df = pd.read_excel(dirpath + r'/' + filename, sheet_name='总表', header=0, dtype=str)
            break
    else:
        df = pd.read_excel(dirpath + r'/' + filename, sheet_name=0, header=0, dtype=str)

    # print(df[['证件号码', '姓名']].head())

    for i in df.columns.values.tolist():
        if '身份证号' in i:
            df = df[['身份证号', '姓名']]
            break
        if '证件号码' in i:
            df = df[['证件号码', '姓名']]
            break
    else:
        raise Exception('未找到身份证号或者姓名列')

    df = df.where(df.notnull(), None)
    return df.dropna(how='all').values.tolist()


def readTransportPerson(dirpath, filename):
    print(filename)
    df_temp = pd.read_excel(dirpath + r'/' + filename, sheet_name=None)
    for i in df_temp.keys():
        if i == '总表':
            df = pd.read_excel(dirpath + r'/' + filename, sheet_name='总表', header=0, dtype=str)
            break
    else:
        for j in range(4):
            df = pd.read_excel(dirpath + r'/' + filename, sheet_name=0, header=j, dtype=str)
            for k in df.columns.values.tolist():
                # print('k: ', k)
                if '姓名' in k or '人员' == k or '密接' == k:
                    # print('kk: ', k)
                    for i in df.columns.values.tolist():
                        if '身份证' in i:
                            df = df[[i, k]]
                            df = df.where(df.notnull(), None)
                            return df.dropna(how='all').values.tolist()
                        if '证件号码' in i:
                            df = df[[i, k]]
                            df = df.where(df.notnull(), None)
                            return df.dropna(how='all').values.tolist()
                    else:
                        return []
        else:
            raise Exception('未找到表头')


def importTransportPerson(connection, l3):
    cursor = connection.cursor()
    cursor.execute('truncate table 转运temp')
    cursor.executemany('insert into 转运temp(身份证号, 姓名) values (:1, :2)', l3)
    connection.commit()
    cursor.execute('merge into 转运 a using 转运temp b on (a.身份证号 = b.身份证号) when not matched then insert values (b.身份证号, b.姓名)')
    connection.commit()
    cursor.close()


def writeSql(connection, l, columns, table_name):
    cursor = connection.cursor()
    table_name_tmp = table_name + 'temp'

    try:
        cursor.execute('select 1 from {}'.format(table_name))
    except DatabaseError:
        sql = 'create table {} ({} varchar2(255)'.format(table_name, columns[0])
        for i in range(1, len(columns)):
            sql += ', {} varchar2(255)'.format(columns[i])
        sql += ')'
        cursor.execute(sql)

    try:
        cursor.execute('select 1 from {}'.format(table_name_tmp))
    except DatabaseError:
        sql = 'create table {} ({} varchar2(255)'.format(table_name_tmp, columns[0])
        for i in range(1, len(columns)):
            sql += ', {} varchar2(255)'.format(columns[i])
        sql += ')'
        cursor.execute(sql)

    cursor.execute('truncate table {}'.format(table_name_tmp))
    insert_sql = 'insert into {} ({}'.format(table_name_tmp, columns[0])
    for i in range(1, len(columns)):
        insert_sql += ', {}'.format(columns[i])
    insert_sql += ') values (:1'
    for i in range(1, len(columns)):
        insert_sql += ', :{}'.format(i+1)
    insert_sql += ')'

    # print(insert_sql)
    # print(l)
    cursor.executemany(insert_sql, l)
    connection.commit()

    insert_sql2 = 'merge into {} a using {} b on (a.姓名 = b.姓名 and a.电话 = b.电话) when not matched then insert values (b.{}'.format(table_name, table_name_tmp, columns[0])
    for i in range(1, len(columns)):
        insert_sql2 += ', b.{}'.format(columns[i])
    insert_sql2 += ')'

    # print(insert_sql2)
    cursor.execute(insert_sql2)
    connection.commit()
    cursor.close()


if __name__ == '__main__':
    warnings.filterwarnings('ignore')
    filePath = '/Users/wengboyu/Desktop/导入文件夹/'
    cx_Oracle.init_oracle_client(lib_dir=r"/Applications/instantclient_19_8")
    connection = cx_Oracle.connect("sjj/123@10.211.55.3:1521/orcl")
    existedFiles = getImportedFile(connection)
    l = []
    l2 = []
    l3 = []
    l4 = []
    test_number = 0

    cause = '涂主任更新'
    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[0:1].lower() != '.' and (filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx'):
                for existedFile in existedFiles:
                    if filename == existedFile:
                        break
                else:
                    if '采集明细' in filename:
                        test_number += importTestData(connection, dirpath, filename)
                    # if '单管' in filename or '混管' in filename :
                    #     l += readExcel(dirpath, filename)
                    #     importFilename(connection, filename)
                    # if '隔离点' in filename:
                    #     l2 += readIsolationPerson(dirpath, filename)
                    #     importFilename(connection, filename)
                    # if '转运人员情况统计表' in filename:
                    #     l3 += readTransportPerson(dirpath, filename)
                    #     importFilename(connection, filename)
                    # if 1 == 1:
                    #     l4 += readExcel(dirpath, filename, ['姓名', '电话', '单位'])

    print(l4)
    writeSql(connection, l4, ['姓名', '电话', '单位'], '支援')


    print('本次导入的采集信息数量为: {}'.format(test_number))

    l = [[i, cause] for i in l]
    print('本次导入的阳性试管号为: {}'.format(l))
    print('--------------- 共有{}条数据 ---------------'.format(len(l)))

    for i in l2:
        i += [cause]
    print('本次导入的隔离点人员名单为: {}'.format(l2))
    print('--------------- 共有{}条数据 ---------------'.format(len(l2)))

    print('本次导入的转运人员名单为: {}'.format(l2))
    print('--------------- 共有{}条数据 ---------------'.format(len(l2)))

    importPositiveData(connection, l)
    importIsolationPerson(connection, l2)
    importTransportPerson(connection, l3)
    connection.commit()
    connection.close()