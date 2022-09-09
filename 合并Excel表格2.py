import pandas
import os

def getdata(filename) -> pandas.DataFrame:
    df = pandas.read_excel(filename, header=None, sheet_name=7)
    dic = {
        '科目编码': 0,
        '科目名称': 3,
        '项目类别': 5,
        '合计': 10
    }
    data = []
    dw_name = df.values[2][0][5:]
    for i in range(9, len(df.values)):
        if pandas.isnull(df.values[i][5]) or df.values[i][5] == '':
            continue
        m = {'单位': dw_name}
        for j in dic.keys():
            m[j] = df.values[i][dic[j]]
        data.append(m)
    df_out = pandas.DataFrame(data)
    return df_out


if __name__ == '__main__':
    filePath = '/Users/wengboyu/Downloads/决算报表（不含汇总）'
    df_out = pandas.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx':
                filename = dirpath + r'/' + filename
                df_out = pandas.concat([df_out, getdata(filename)])

    df_out.to_excel('/Users/wengboyu/Downloads/task3.xlsx')


