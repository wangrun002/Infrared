#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import shutil
import sys

def build_all_test_case():
    global ALL_TEST_CASE
    ALL_TEST_CASE = [
                        [ ["screen", "numb_key", "All", "TV", INTERVAL_TIME[1]],
                          NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[0], EXIT_TO_SCREEN],

                        [ ["screen", "numb_key", "All", "TV", INTERVAL_TIME[1]],
                          NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[1], EXIT_TO_SCREEN],

                        [ ["screen", "numb_key", "All", "TV", INTERVAL_TIME[1]],
                          NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[2], EXIT_TO_SCREEN],
    ]

choice_switch_mode = 2
ser_cable_numb = 4

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path,"main_program","main.py")
test_file_path = os.path.join(os.getcwd(),"main.py")
# print(parent_path)
# print(test_file_path)

shutil.copy(main_file_path,os.getcwd())
os.system("python %s %d %d" % (test_file_path,choice_switch_mode,ser_cable_numb))
os.unlink(test_file_path)