import os
import re

filePath = '/Users/wengboyu/Desktop/导入文件夹/8.31晚12点/单管'
a = 0

for dirpath, dirnames, filenames in os.walk(filePath):
    for filename in filenames:
        if filename[0:1].lower() != '.' and (filename[-3:].lower() == 'xls' or filename[-4:].lower() == 'xlsx'):
            a += int(re.search('\d+人', filename).group(0).replace('人', ''))
            print(a)