import pandas as pd
from pathlib import Path, PureWindowsPath
import os


def get_file(patterns, path):
    all_files = []
    p = Path(path)
    for item in patterns:
        file_name = p.rglob(f'**/*.{item}')
        all_files.extend(file_name)
    return all_files

def search(dirname, filename):
    """
    :param dirname: 需要查找的目录
    :param filename: 文件类型
    :return:
    """
    result =[]
    for item in os.listdir(dirname):
        item_path = os.path.join(dirname, item)
        if os.path.isdir(item_path):
            search(item_path, filename)
        elif os.path.isfile(item_path):
            if filename in item:
                result.append(item_path)
    return result





if __name__ == '__main__':
    # 存储位置及文件名
    # save_path = r'E:/有资产数据的无形资产.xlsx'
    save_path = r'/Users/wengboyu/Downloads/项目库比对/项目库.xlsx'
    # dirname =r'E:/有资产数据的无形资产/'
    dirname = r'/Users/wengboyu/Downloads/项目库比对/新建文件夹'
    filename = ".xls"
    merge_file_name = pd.DataFrame()
    for i in search(dirname, filename):
        print(i)
        df3 = pd.read_excel(str(i), index=None, header=2)
        merge_file_name = merge_file_name.append(df3, ignore_index=False)
    # merge_file_name = merge_file_name.drop_duplicates()
    merge_file_name.to_excel(save_path, index=False)

