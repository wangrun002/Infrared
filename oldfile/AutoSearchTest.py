#!/usr/bin/python
# -*- coding: utf-8 -*-

import os,sys,shutil

ReadFileName = [
                'UpperLimitTPSearchCommand(Z6).txt',        #0
                'UpperLimitChannelSearchCommand(Y3).txt',   #1

                '88Sat6FBlindSearchCommand.txt',            #2直连
                '88Sat6FSuperBlindSearchCommand.txt',       #3直连

                'Z6Sat6FBlindSearchCommand.txt',            #4
                'Z6Sat6FSuperBlindSearchCommand.txt',       #5

                'Y3Sat6FBlindSearchCommand.txt',            #6
                'Y3Sat6FSuperBlindSearchCommand.txt',       #7

                '138Sat6FBlindSearchCommand.txt',           #8
                '138Sat6FSuperBlindSearchCommand.txt',      #9
                '138Sat6FBlindSearchAddCommand.txt',        #10

                'PLPDSat6FBlindSearchCommand.txt',          #11
                'PLPDSat6FSuperBlindSearchCommand.txt',     #12

                'FactoryResetSearchCommand.txt',            #13
                'AddNewSat20SearchCommand.txt',             #14
                'USBUpgradeUser20SatCommand.txt'            #15
                ]

Serial_SER_Dict = {
                        "1":"FTDVKA2HA",
                        "2":"FTGDWJ64A",
                        "3":"FT9SP964A",
                        "4":"FTHB6SSTA"
                }

SearchTimes = 2  #搜索次数
CommandFileNumber = 8  #指定搜索命令文件编号
SER_Value = 4   #指定串口线的序号

ParentOfCurProPath = os.path.abspath(os.path.join(os.getcwd(), "..")) #当前程序路径的上级路径
#print(ParentOfCurProPath)
MainProgramPath = os.path.join(ParentOfCurProPath,"MainProgram","NewSatSearchIncludeArgvPara.py")
#print(MainProgramPath)

#os.system("python ./FactoryReset.py")
#os.system("python ./NewSatSearch.py %d %d" % (3,8))

#os.system("cd {} && python NewSatSearch.py {} {}".format(MainProgramPath,2,8))

shutil.copy(MainProgramPath, os.getcwd())
os.system("python ./NewSatSearchIncludeArgvPara.py %d %d %d" % (SearchTimes,CommandFileNumber,SER_Value))
os.unlink(os.path.join(os.getcwd(),"NewSatSearchIncludeArgvPara.py"))