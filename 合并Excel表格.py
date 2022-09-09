import pandas
import os

def getdata(filename) -> pandas.DataFrame:
    df = pandas.read_excel(filename, header=None)
    dic = {
        '科目编码': 0,
        '科目名称': 3,
        '项目类别': 5,
        '基建项目属性': 8,
        '合计': 10,
        '工资福利支出': 11,
        '商品和服务支出': 25,
        '对个人和家庭的补助': 53,
        '债务利息及费用支出': 66,
        '资本性支出（基本建设）': 71,
        '资本性支出': 84,
        '对企业补助（基本建设）': 101,
        '对企业补助': 104,
        '对社会保障基金补助': 110,
        '其他支出': 114
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
    filePath = '/Users/wengboyu/Downloads/项目库比对/新建文件夹'
    df_out = pandas.DataFrame()

    for dirpath, dirnames, filenames in os.walk(filePath):
        for filename in filenames:
            if filename[-3:].lower() == 'xls' or filename[-3:].lower() == 'xlsx':
                filename = dirpath + r'/' + filename
                df_out = pandas.concat([df_out, getdata(filename)])

    df_out.to_excel('/Users/wengboyu/Downloads/项目库比对/test.xlsx')


