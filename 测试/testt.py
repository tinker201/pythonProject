import pandas

if __name__ == '__main__':
    df = pandas.read_excel(r'/Users/wengboyu/Downloads/项目库比对/国家税务总局台州市路桥区税务局.XLS', header=None)
    print(df.values[10][5])
    print(pandas.isnull(df.values[10][5]))
