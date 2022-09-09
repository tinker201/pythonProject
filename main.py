# 这是一个示例 Python 脚本。

# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。

import numpy as np
import pandas as pd

def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 ⌘F8 切换断点。


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    print_hi('world')

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助

    print(np.random.rand(3))
    print(np.random.randn(50))
    print(np.arange(8,19))

    target = np.arange(9).reshape(3, 3)
    print(target.any())
    print(pd.__version__)

    df = pd.read_excel('/Users/wengboyu/Downloads/财报资产云/财报数据整理.xlsx')
    print(df.index)
    print(df.columns)
    print(df['增加额'])
    print(df.dtypes)
    # print(df.T[233])
    print(list('abc'))

    print(list('abcdefg')[:6])