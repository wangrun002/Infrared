#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import shutil
import platform

choice_case_numb = 43

case_info = []

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path, "MainProgram", "Edit_same_start_time_res_event.py")
test_file_path = os.path.join(os.getcwd(), "Edit_same_start_time_res_event.py")

shutil.copy(main_file_path, os.getcwd())
if platform.system() == "Windows":
    os.system("python %s %d" % (test_file_path, choice_case_numb))
elif platform.system() == "Linux":
    os.system("python3 %s %d" % (test_file_path, choice_case_numb))
os.unlink(test_file_path)


