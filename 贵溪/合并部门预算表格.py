import pandas
import os

def getdata(dirpath, filename) -> pandas.DataFrame:
    df = pandas.read_excel(dirpath + r'/' + filename, header=None, sheet_name='基本')
    dic = {
        '单位编码': 0,
        '科目编码（类）': 1,
        '科目编码（款）': 2,
        '科目编码（项）': 3,
        '科目名称': 4,
        '经济分类名称': 5,
        '年初预算': 6,
        '财政拨款': 7,
        '事业收入': 16,
        '事业单位经营收入': 17,
        '附属单位上缴收入': 18,
        '上级补助收入': 19,
        '其他收入': 20,
        '用事业基金弥补收支差额': 21,
        '上年结转（结余）': 22
    }
    data = []
    m = {}
    m['单位名称'] = filename
    for i in range(9, len(df.values)):
        if pandas.isnull(df.values[i][5]) or df.values[i][5] == '':
            continue
        for j in dic.keys():
            m[j] = df.values[i][dic[j]]
        data.append(m)
    df_out = pandas.DataFrame(data)
    return df_out


if __name__ == '__main__':
    filePath = '/Users/wengboyu/Downloads/社保学习第一周'
    df_out = pandas.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx':
                # filename = dirpath + r'/' + filename
                df_out = pandas.concat([df_out, getdata(dirpath, filename)])

    # df_out.to_excel('/Users/wengboyu/Downloads/task.xlsx')


