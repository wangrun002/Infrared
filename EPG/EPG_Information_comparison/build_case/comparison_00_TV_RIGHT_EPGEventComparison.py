#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import shutil
import platform

choice_case_numb = 0

case_info = ['00', 'GX', 'TV', 'RIGHT', 'EPGEventComparison']

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path, "MainProgram", "EPG_Information_comparison.py")
test_file_path = os.path.join(os.getcwd(), "EPG_Information_comparison.py")

shutil.copy(main_file_path, os.getcwd())
if platform.system() == "Windows":
    os.system("python %s %d" % (test_file_path, choice_case_numb))
elif platform.system() == "Linux":
    os.system("python3 %s %d" % (test_file_path, choice_case_numb))
os.unlink(test_file_path)
