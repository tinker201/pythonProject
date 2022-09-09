import cx_Oracle
import pandas
import os

def getdata(filename) -> pandas.DataFrame:
    df = pandas.read_excel(filename, header=None, sheet_name=1)
    dic = {
        '科目编码': 0,
        '科目名称': 3,
        '项目类别': 5,
        '合计': 10
    }
    data = []
    dw_name = df.values[2][0][5:]
    m = {}
    m['单位名称'] = dw_name
    m['归属于母公司所有者的净利润_本期金额'] = df.values[9][6]
    m['归属于母公司所有者的净利润_上期金额'] = df.values[9][7]
    data.append(m)
    df_out = pandas.DataFrame(data)
    return df_out

if __name__ == '__main__':
    filePath = '/Users/wengboyu/Downloads/社保学习第一周'
    df_out = pandas.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx':
                filename = dirpath + r'/' + filename
                df_out = pandas.concat([df_out, getdata(filename)])

    # df_out.to_excel('/Users/wengboyu/Downloads/task.xlsx')


