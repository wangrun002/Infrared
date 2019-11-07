#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import time

COMMD_FILE_NAME = [
                "Z6Sat_6F_Blind_SearchCommand.txt",                 #0
                "Z6Sat_6F_SuperBlind_SearchCommand.txt",            #1

                "Y3Sat_6F_Blind_SearchCommand.txt",                 #2
                "Y3Sat_6F_SuperBlind_SearchCommand.txt",            #3

                "88Sat_6F_Blind_SearchCommand.txt",                 #4直连
                "88Sat_6F_SuperBlind_SearchCommand.txt",            #5直连

                "138Sat_6F_Blind_SearchCommand.txt",                #6
                "138Sat_6F_SuperBlind_SearchCommand.txt",           #7
                "138Sat_6F_BlindAdd_SearchCommand.txt",             #8

                "PLPDSat_6F_Blind_SearchCommand.txt",               #9
                "PLPDSat_6F_SuperBlind_SearchCommand.txt",          #10

                "Z6Sat_6F_UpperLimitTP_SearchCommand.txt",          #11
                "Y3Sat_6F_UpperLimitChannel_SearchCommand.txt",     #12

                'Factory_6F_Reset_SearchCommand.txt',               #13
                'Add_6F_20NewSat_SearchCommand.txt',                #14
                'USBUpgradeUser20SatCommand.txt'                    #15
                ]

serial_ser_value =  {
                        "1": "FTDVKA2HA",
                        "2": "FTGDWJ64A",
                        "3": "FT9SP964A",
                        "4": "FTHB6SSTA"
                    }

choice_commd_file_numb = 12                 # 选择想要搜索的卫星的指令文件
search_numb = 72                            # 搜索次数
ser_cable_numb = 4                          # USB转串口线编号

parent_of_current_path = os.path.abspath(os.path.join(os.getcwd(),".."))        # 当前程序路径的上级路径
main_prog_path = os.path.join(parent_of_current_path,"MainProgram","SatSearchIncludeArgvPara_Class.py")

shutil.copy(main_prog_path,os.getcwd())
os.system("python ./SatSearchIncludeArgvPara_Class.py %d %d %d" % (13,1,ser_cable_numb))
time.sleep(15)
os.system("python ./SatSearchIncludeArgvPara_Class.py %d %d %d" % (choice_commd_file_numb,search_numb,ser_cable_numb))
os.unlink(os.path.join(os.getcwd(),"SatSearchIncludeArgvPara_Class.py"))