import pandas
import os


def getdata(filename) -> pandas.DataFrame:
    df = pandas.read_excel(filename, header=0, sheet_name='固定资产')
    data = []
    header = df.columns.to_list()
    values = df.values.tolist()
    # print(values)
    # print('------------------')
    for i in range(len(values)):
        m = {}
        for j in range(len(header)):
            m[header[j]] = values[i][j]
        if m['分类名称'] == '原值合计':
            data.append(m)
    df_out = pandas.DataFrame(data)
    # df_out = pandas.DataFrame()
    return df_out


if __name__ == '__main__':
    filePath = r'/Users/wengboyu/Downloads/财报资产云'
    fileName = '财报数据.xls'
    df_out = getdata(filePath + r'/' + fileName)
    df_out.to_excel(r'/Users/wengboyu/Downloads/财报资产云/财报数据整理.xlsx')

    def joyfunc(x):
        return x * 2

    l = [joyfunc(i) for i in range(5)]
    print(l)