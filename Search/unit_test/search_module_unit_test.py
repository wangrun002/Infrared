#!/usr/bin/python
# -*- coding: utf-8 -*-

from unittest import defaultTestLoader
import unittest
import os
import shutil


# print(os.path.abspath(os.path.dirname(os.path.dirname(__file__))))
# print(os.path.abspath(os.path.dirname(os.getcwd())))
# print(os.path.abspath(os.path.join(os.getcwd(), "..")))
# print(os.path.dirname(os.getcwd()))


def get_all_case():
    discover = unittest.defaultTestLoader.discover(case_path,pattern="test*.py")
    suite = unittest.TestSuite()
    suite.addTest(discover)
    return suite

if __name__ == "__main__":
    parent_path = os.path.dirname(os.getcwd())
    main_file_path = os.path.join(parent_path,"MainProgram","NewAddSatBlind_IncludeArgvParam.py")
    shutil.copy(main_file_path,os.getcwd())
    case_path = os.path.join(parent_path,"test_cases")
    print(case_path)
    runner = unittest.TextTestRunner()
    runner.run(get_all_case())