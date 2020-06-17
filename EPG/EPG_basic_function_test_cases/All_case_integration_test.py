#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import re
import time
import platform

test_cast_list = []

for filename in os.listdir():
    if os.path.splitext(filename)[1] == '.py':
        if re.match(r"^case", filename):
            print(filename)
            test_cast_list.append(filename)


def get_case_numb(name):        # 提取每个文件名称中的数字
    return int(re.findall(r'(\d+)', name)[0])


test_cast_list.sort(key=get_case_numb)     # 按照文件名称中的数字进行先后排序
print(test_cast_list)

for i in range(len(test_cast_list)):
    if platform.system() == "Windows":
        os.system("python %s" % test_cast_list[i])
    elif platform.system() == "Linux":
        os.system("python3 %s" % test_cast_list[i])
    time.sleep(10)
