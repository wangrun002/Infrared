#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
voltage = { "0":"13V",
            "1":"18V",
            "2":"Off"
            }

22k = { "0":"On",
        "1":"Off"
        }

diseqc 1.0 = {  "0":"Off",
                "1":"Port1",
                "2":"Port2",
                "3":"Port3",
                "4":"Port4"
                }

all_sat_commd = [
                    choice_enter_antenna_mode,
                    search_preparatory_work,
                    sat_param_list,
                    choice_search_mode,
                    choice_save_type,
                    choice_exit_mode,
                    other_operate,
                    normal_cycle_times,
                    control_upper_limit_cycle_times,
                ]
'''

from datetime import datetime,timedelta
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font,colors,Alignment
from openpyxl.utils import get_column_letter,column_index_from_string
from threading import Timer
import serial
import serial.tools.list_ports
import re
import time
import os
import sys
import random
import logging

LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s %(message)s"   # 配置输出日志的格式
DATE_FORMAT = "%Y-%m-%d %H:%M:%S %a"    # 配置输出时间的格式
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT)


KEY = { "0": "A1 F1 22 DD 00",
        "UP":"A1 F1 22 DD 11",
        "DOWN":"A1 F1 22 DD 14",
        "LEFT":"A1 F1 22 DD 12",
        "RIGHT": "A1 F1 22 DD 13",
        "OK": "A1 F1 22 DD 15",
        "MENU": "A1 F1 22 DD 0C",
        "EXIT": "A1 F1 22 DD 0D",
        "RED": "A1 F1 22 DD 19",
        "GREEN":"A1 F1 22 DD 1A",
        "BLUE": "A1 F1 22 DD 1C",
        "INFO": "A1 F1 22 DD 1F",}

PRESET_SAT_NAME = ['Nilesat', 'Hotbird', 'Badr 4/5/6/7 K', 'Thor 5/6/7', 'Turksat 2A/3A', 'BulgariaSat-1',
                  'Eutelsat 3B C', 'Eutelsat 4A', 'Eutelsat 9B', 'Amos 5 K', 'Astra 1E/3B', 'Arabsat 5A C',
                  'Arabsat 5A K', 'Eutelsat 33E K', 'C_Paksat 1R', 'Intelsat 12', 'Azerspace K', 'Intelsat 10',
                  'Yamal 202', 'Turksat 4B K', 'Belintersat K', 'TurkmenAlem', 'Yahsat 1A', 'Intelsat 707 C',
                  'Yamal 402 K', 'NSS 12 C', 'Intelsat 33e C', 'Intelsat 33e K', 'Intelsat 902 C', 'Intelsat 20 K',
                  'ABS 2/2A K', 'APSTAR 7 C', 'Thaicom 5/6 C', 'Thaicom 5/8 K', 'Express MD1 C', 'Insat 4A K',
                  'ST 2 K', 'Yamal 201 K', 'Measat 3/3A K', 'Measat 3/3A C', 'NSS 6', 'Express AM33 K',
                  'Koreasat 5 K', 'JCSat 3A K', 'JCSat 3A C', 'Vinasat 1 K', 'Telstar 18 K', 'Express AM5 K',
                  'Express AM5 C', 'Optus D1', 'Superbird B2', 'Intelsat 2/8', 'Amos 2/3/7', 'Eutelsat 5 C',
                  'Eutelsat 5 K', 'Eutelsat 8 C', 'Express AM44 K', 'Eutelsat 12', 'Telstar 12V', 'ABS-3 K',
                  'SES 4 K', 'Intelsat 905 C', 'AlComSat 1', 'Intelsat 907 C', 'Intelsat 907 K', 'Hispasat 4/5/6',
                  'Intelsat 35e', 'Intelsat 707 K', 'Intelsat 21 K', 'Amazonas 2/3 K', 'Asiasat 7 C', 'Chinas6b_C']

NORMAL_SEARCH_TIMES = 3                  # 10 普通盲扫次数
SUPER_SEARCH_TIMES = 10                  # 10 超级盲扫次数
INCREMENTAL_SEARCH_TIMES = 3             # 15 累加搜索次数
UPPER_LIMIT_SEARCH_TIMES = 72            # 72 上限搜索初始次数
UPPER_LIMIT_CYCLE_TIMES = 10              # 5  上限搜索循环次数
UPPER_LIMIT_LATER_SEARCH_TIMES = 1       # 20 上限搜索后其他情况执行测试
ONLY_EXECUTE_ONE_TIME = 1                # 单独场景只执行一次
NOT_UPPER_LIMIT_LATER_SEARCH_TIME = 0


ENTER_ANTENNA_SETTING = [KEY["MENU"],KEY["OK"]]
DELETE_ALL_SAT = [KEY["RED"],KEY["0"],KEY["RED"],KEY["OK"]]
ADD_ONE_SAT = [KEY["GREEN"],KEY["UP"],KEY["OK"],KEY["INFO"]]
SEARCH_PREPARATORY_WORK = [[],[DELETE_ALL_SAT,ADD_ONE_SAT]]
CHOICE_BLIND_MODE = [KEY["RIGHT"],KEY["OK"],KEY["OK"]]
CHOICE_SUPERBLIND_MODE = [KEY["BLUE"], KEY["RIGHT"], KEY["OK"], KEY["OK"]]
CHOICE_NOT_SEARCH = []
CHOICE_SAVE_TYPE = [[KEY["OK"]],[KEY["LEFT"],KEY["OK"]]]
EXIT_ANTENNA_SETTING = [KEY["EXIT"],KEY["EXIT"]]
NOT_OTHER_OPERATE = []
RESET_FACTORY = [KEY["MENU"],KEY["RIGHT"],KEY["DOWN"],KEY["OK"],
                 KEY["0"],KEY["0"],KEY["0"],KEY["0"],
                 KEY["OK"]]
DELETE_SPECIFY_SAT_ALL_TP = [KEY["GREEN"],KEY["0"],KEY["RED"],KEY["OK"]]
DELETE_ALL_CH = [KEY["MENU"],KEY["LEFT"],KEY["LEFT"],KEY["UP"],KEY["OK"],KEY["OK"]]
UPPER_LIMIT_LATER_NOT_DEL_SAT_TP_SEARCH_CONT = [KEY["EXIT"]]
EXIT_TO_SCREEN = [KEY["EXIT"],KEY["EXIT"],KEY["EXIT"]]


xlsx_title = [
				"搜索模式",
				"搜索次数",
				"搜索TP数",
				"搜索节目数",
				"保存TP数",
				"保存节目数",
				"搜索时间",
				{"数据类别":["TP","All","TV","Radio","CH_Name"]},
				"TP"
			]

class MyGlobal():
    def __init__(self):
        self.ser_cable_num = 5                          # USB转串口线编号
        self.switch_commd_stage = 0                     # 切换发送命令的阶段
        self.setting_option_numb = 0                    # 设置项位置number

        self.switch_lnb_power_state = True              # 用来控制切换本振power选项参数切换的状态变量

        self.blind_judge_polar = [[],[],set(),'']       # 用于判断极化
        self.all_tp_list = []                           # 用于存放搜索到的TP
        self.channel_info = {}                          # 用于存放各个TP下搜索到的电视和广播节目名称
        self.search_datas = [0,0,0,0,0,0,0,0,0]         # 用于存放xlsx_title中的数据
        self.tv_radio_tp_count = [0,0,0,0,0]            # [GL.tv_radio_tp_count[0],GL.tv_radio_tp_count[1],GL.tv_radio_tp_count[2],GL.tv_radio_tp_count[3],GL.tv_radio_tp_count[4]]
        self.tv_radio_tp_accumulated = [[],[],[],[]]    # 用于统计每轮搜索累加的TV、Radio、TP数以及保存TP数的值
        self.xlsx_data_interval = 0                     # 用于计算每轮搜索写xlsx时的间隔
        self.delay_reset_factory_time = 30              # 用于恢复出厂设置延时停止程序

        self.sub_thread_delay_count_state = True        # 用于简单的Reset Factory程序,延时程序退出的执行状态

        self.MAIN_LOOP_STATE = True                     # 主程序状态
        self.commd_global_length = 0                    # 用于某阶段单个list时,为命令长度,多个list时,为list个数
        self.commd_global_pos = 0                       # 用于某阶段单个list时,为当前命令位置,多个list时,为当前命令所在的list位置
        self.commd_single_length = 0                    # 用于某阶段有多个list时,各个list的长度
        self.commd_single_pos = 0                       # 用于某阶段多个list时,当前命令所在的list的位置
        self.delay_save_channel_time = 2                # 用于搜索结束后的命令发送延时间隔
        self.delay_save_channel_state = True            # 用于控制搜索结束后的命令发送延时状态
        # self.search_end_state = False                   # 用于识别搜索结束的状态变量
        self.delay_delete_all_sat_time = 5              # 用于卫星参数设置前准备工作中删除所有卫星延时
        self.delay_delete_all_sat_state = True          # 用于控制卫星参数设置前准备工作中删除所有卫星后命令发送延时状态
        self.delay_cyclic_search_time = 5               # 用于当前轮次搜索结束后,切换到下一轮次时,延时5秒
        self.delay_cyclic_search_state = False          # 用于当前轮次搜索结束后,切换到下一轮次搜索时的状态变量

        self.searched_sat_name = []                     # 用于保存搜索过程中搜索过的卫星的名称,便于搜索达到上限后,删除指定的卫星,不能被清空
        self.upper_limit_state = False                  # 用于控制搜索达到上限的时其他操作的状态变量，false不执行，true时执行
        self.delay_reset_factory_state = True           # 用于控制恢复出厂设置延时的状态变量
        self.delay_delete_all_channels_time = 20        # 用于控制删除所有节目后的延时时间
        self.delay_delete_all_channels_state = True     # 用于控制删除所有节目的状态变量
        self.other_operate_sub_stage = 0                # 用于other operate步骤中的子阶段
        self.delay_delete_specify_sat_tp_time = 10      # 用于控制删除指定的卫星TP的延时时间
        self.delay_delete_specify_sat_tp_state = True   # 用于控制删除指定的卫星TP延时的状态变量

        self.upper_limit_to_save_channel_stage_time = 2 # 用于控制搜索达到上限后到进入保存节目阶段的延时时间
        self.upper_limit_to_save_channel_stage_state = False    # 用于控制搜索达到上限后到进入保存节目阶段的延时的状态变量
        self.random_choice_sat = []                     # 用于存放搜索达到上限后每次随机选择的卫星,然后进行删除其TP
        self.delete_ch_finish_kws = "[PTD]All programs deleted successfully"
        self.delete_ch_finish_state = False
        self.save_ch_finish_state = False


        self.sat_param_save = ["", "", "", "", "", ""]      # [sat_name,LNB_Power,LNB_Fre,22k,diseqc1.0,diseqc1.1]
        self.sat_param_kws =    [
                                    "[PTD]sat_name=",
                                    "[PTD]LNB1=",
                                    "[PTD]--[0:ON,1:OFF]---22K",
                                    "[PTD]--set diseqc 1.0",
                                    "--------set diseqc 1.1",
                                ]

        self.search_monitor_kws = [
                                    "[PTD]SearchStart",		#0
                                    "[PTD]TV------",		#1
                                    "[PTD]Radio-----",		#2
                                    "[PTD]SearchFinish",	#3
                                    "[PTD]get :  fre",		#4
                                    "[PTD]TP_save=",		#5
                                    "[PTD]TV_save=",		#6
                                    "[PTD]maximum_tp",		#7
                                    "[PTD]maximum_channel",	#8
                                    "get blind - fre",      #9
                                    ]

        self.all_sat_commd =   [
                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Chinas6b_C", "Polar=0", "5150/5750", "22K=1", "2", "0", "Blind"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        NORMAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Chinas6b_C", "Polar=0", "5150/5750", "22K=1", "2", "0", "SuperBlind"],
                                        CHOICE_SUPERBLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        SUPER_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Asiasat 7 C", "Polar=0", "5150/5750", "22K=1", "1", "0", "Blind"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        NORMAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Asiasat 7 C", "Polar=0", "5150/5750", "22K=1", "1", "0", "SuperBlind"],
                                        CHOICE_SUPERBLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        SUPER_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "Blind"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        NORMAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "SuperBlind"],
                                        CHOICE_SUPERBLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        SUPER_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["ST 2 K", "Polar=0", "10600/0", "22K=1", "0", "0", "Blind"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        NORMAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["ST 2 K", "Polar=0", "10600/0", "22K=1", "0", "0", "SuperBlind"],
                                        CHOICE_SUPERBLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        SUPER_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["PLPD", "Polar=0", "5150/5750", "22K=0", "1", "0", "Blind"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        NORMAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[1],
                                        ["PLPD", "Polar=0", "5150/5750", "22K=0", "1", "0", "SuperBlind"],
                                        CHOICE_SUPERBLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        SUPER_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "Incremental"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, NOT_OTHER_OPERATE,
                                        INCREMENTAL_SEARCH_TIMES, NOT_UPPER_LIMIT_LATER_SEARCH_TIME],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "ChUL"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, DELETE_ALL_CH,
                                        UPPER_LIMIT_SEARCH_TIMES,UPPER_LIMIT_CYCLE_TIMES],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "ChUL_Cont."],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, UPPER_LIMIT_LATER_NOT_DEL_SAT_TP_SEARCH_CONT,
                                        UPPER_LIMIT_SEARCH_TIMES,UPPER_LIMIT_LATER_SEARCH_TIMES],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Telstar 18 K", "Polar=0", "10600/0", "22K=0", "4", "0", "ChUL_DelTp_Cont."],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[0],
                                        EXIT_ANTENNA_SETTING, DELETE_SPECIFY_SAT_ALL_TP,
                                        UPPER_LIMIT_SEARCH_TIMES, UPPER_LIMIT_LATER_SEARCH_TIMES],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Chinas6b_C", "Polar=0", "5150/5750", "22K=1", "2", "0", "TpUL"],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[1],
                                        EXIT_ANTENNA_SETTING, RESET_FACTORY,
                                        UPPER_LIMIT_SEARCH_TIMES,UPPER_LIMIT_CYCLE_TIMES],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Chinas6b_C", "Polar=0", "5150/5750", "22K=1", "2", "0", "TpUL_Cont."],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[1],
                                        EXIT_ANTENNA_SETTING, UPPER_LIMIT_LATER_NOT_DEL_SAT_TP_SEARCH_CONT,
                                        UPPER_LIMIT_SEARCH_TIMES,UPPER_LIMIT_LATER_SEARCH_TIMES],

                                    [   ENTER_ANTENNA_SETTING, SEARCH_PREPARATORY_WORK[0],
                                        ["Chinas6b_C", "Polar=0", "5150/5750", "22K=1", "2", "0", "TpUL_DelTp_Cont."],
                                        CHOICE_BLIND_MODE, CHOICE_SAVE_TYPE[1],
                                        EXIT_ANTENNA_SETTING, DELETE_SPECIFY_SAT_ALL_TP,
                                        UPPER_LIMIT_SEARCH_TIMES,UPPER_LIMIT_LATER_SEARCH_TIMES],

                                    [   RESET_FACTORY, SEARCH_PREPARATORY_WORK[0],
                                        ["Reset","Factory"],ONLY_EXECUTE_ONE_TIME],

                                    [   DELETE_ALL_CH, SEARCH_PREPARATORY_WORK[0],
                                        ["Delete","AllCH"],ONLY_EXECUTE_ONE_TIME],
                                ]

def check_ports(ser_cable_num):
    serial_ser =    {
                        "1":"FTDVKA2HA",
                        "2":"FTGDWJ64A",
                        "3":"FT9SP964A",
                        "4":"FTHB6SSTA",
                        "5":"FTDVKPRSA",
                    }
    send_port_desc = "USB-SERIAL CH340"
    receive_port_desc = "USB Serial Port"
    ports = list(serial.tools.list_ports.comports())
    for i in range(len(ports)):
        logging.info("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device,ports[i].description,ports[i].hwid))
    if len(ports) <= 0:
        logging.info("无可用端口")
    elif len(ports) == 1:
        logging.info("只有一个可用端口:{}".format(ports[0].device))
    elif len(ports) >=2:
        # for i in range(len(ports)):
        #     ports_com.append(str(ports[i]))
        #     if send_port_desc in str(ports[i]):
        #         send_com = ports[i][0]
        #     elif receive_port_desc in str(ports[i]) and serial_ser[str(ser_cable_num)] in str(ports[i][2]):
        #         receive_com = ports[i][0]
        #         logging.info(ports[i][2])
        if serial.tools.list_ports.grep(send_port_desc):
            send_com = next(serial.tools.list_ports.grep(send_port_desc)).device
        if serial.tools.list_ports.grep(receive_port_desc):
            receive_com = next(serial.tools.list_ports.grep(receive_port_desc)).device

    return send_com,receive_com

def serial_set(ser,ser_name,ser_baudrate):
    ser.port = ser_name
    ser.baudrate = ser_baudrate
    ser.bytesize = 8
    ser.parity = "N"
    ser.stopbits = 1
    ser.timeout = 1
    ser.open()

def hex_strs_to_bytes(strings):
    strs = strings.replace(" ", "")
    return bytes.fromhex(strs)

def delay_single_reset_factory():
    global t
    GL.delay_reset_factory_time -= 1
    logging.info("恢复出厂设置退出倒计时:{}".format(GL.delay_reset_factory_time))
    if GL.delay_reset_factory_time == 0:
        GL.MAIN_LOOP_STATE = False
        sys.exit(0)
    if GL.delay_reset_factory_time > 0:
        t = Timer(1.0,delay_single_reset_factory).start()

def delay_reset_factory():
    global t
    GL.delay_reset_factory_time -= 1
    logging.info("恢复出厂设置到大画面延时退出倒计时:{}".format(GL.delay_reset_factory_time))
    if GL.delay_reset_factory_time == 0:
        GL.delay_reset_factory_state =True
        GL.delay_reset_factory_time = 30
        sys.exit(0)
    if GL.delay_reset_factory_time > 0:
        t = Timer(1.0, delay_reset_factory).start()

def delay_delete_all_channels():
    global t
    GL.delay_delete_all_channels_time -= 1
    logging.info("删除所有节目后延时退出倒计时:{}".format(GL.delay_delete_all_channels_time))
    if GL.delay_delete_all_channels_time == 0:
        GL.delay_delete_all_channels_state = True
        GL.delay_delete_all_channels_time = 20
        sys.exit(0)
    if GL.delay_delete_all_channels_time > 0:
        t = Timer(1.0, delay_delete_all_channels).start()

def delay_delete_specify_sat_tp():
    global t
    GL.delay_delete_specify_sat_tp_time -= 1
    logging.info("删除指定卫星下TP后延时退出倒计时:{}".format(GL.delay_delete_specify_sat_tp_time))
    if GL.delay_delete_specify_sat_tp_time == 0:
        GL.delay_delete_specify_sat_tp_state = True
        GL.delay_delete_specify_sat_tp_time = 10
        sys.exit(0)
    if GL.delay_delete_specify_sat_tp_time > 0:
        t = Timer(1.0, delay_delete_specify_sat_tp).start()

def delay_upper_limit_to_save_channel_stage():
    global t
    GL.upper_limit_to_save_channel_stage_time -= 1
    logging.info("达到上限后进入保存节目阶段延时退出倒计时:{}".format(GL.upper_limit_to_save_channel_stage_time))
    if GL.upper_limit_to_save_channel_stage_time == 0:
        GL.upper_limit_to_save_channel_stage_state = True
        GL.upper_limit_to_save_channel_stage_time = 2  # 将该值恢复为初始值
        sys.exit(0)
    if GL.upper_limit_to_save_channel_stage_time > 0:
        t = Timer(1.0, delay_upper_limit_to_save_channel_stage).start()

def delay_save_channel():
    global t
    GL.delay_save_channel_time -= 1
    logging.info("保存节目延时退出倒计时:{}".format(GL.delay_save_channel_time))
    if GL.delay_save_channel_time == 0:
        GL.delay_save_channel_state = True
        GL.delay_save_channel_time = 2  # 将该值恢复为初始值
        sys.exit(0)
    if GL.delay_save_channel_time > 0:
        t = Timer(1.0,delay_save_channel).start()

def delay_cyclic_search():
    global t
    GL.delay_cyclic_search_time -= 1
    logging.info("进入下一轮搜索延时退出倒计时:{}".format(GL.delay_cyclic_search_time))
    if GL.delay_cyclic_search_time == 0:
        GL.delay_cyclic_search_state = False
        GL.delay_cyclic_search_time = 5  # 将该值恢复为初始值
        sys.exit(0)
    if GL.delay_cyclic_search_time > 0:
        t = Timer(1.0, delay_cyclic_search).start()

def delay_delete_all_sat():
    global t
    GL.delay_delete_all_sat_time -= 1
    logging.info("删除卫星保存延时倒计时:{}".format(GL.delay_delete_all_sat_time))
    if GL.delay_delete_all_sat_time == 0:
        GL.delay_delete_all_sat_state = True
        GL.delay_delete_all_sat_time = 5
        sys.exit(0)
    if GL.delay_delete_all_sat_time > 0:
        t = Timer(1.0,delay_delete_all_sat).start()


def add_write_data_to_txt(file_path,write_data):    # 追加写文本
    with open(file_path,"a+",encoding="utf-8") as fo:
        fo.write(write_data)

def judge_write_file_exist():
    if not os.path.exists(write_xlsx_relative_path):
        os.mkdir(write_xlsx_relative_path)
    if not os.path.exists(write_txt_relative_path):
        os.mkdir(write_txt_relative_path)

def judge_and_wirte_data_to_xlsx():
    alignment = Alignment(horizontal="center",vertical="center",wrapText=True)
    if not os.path.exists(write_xlsx_path):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.column_dimensions['A'].width = 11
        for i in range(len(xlsx_title)):
            if i < len(xlsx_title) - 2:
                ws.cell(i + 1, 1).value = xlsx_title[i]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(xlsx_title) - 2:
                ws.cell(i + 1, 1).value = list(xlsx_title[i].keys())[0]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(xlsx_title) - 1:
                ws.cell(i + 1, 1).value = xlsx_title[i]
                ws.cell(i + 1, 1).alignment = alignment

    elif os.path.exists(write_xlsx_path):
        wb = load_workbook(write_xlsx_path)
        sheets_name_list = wb.sheetnames
        logging.info(sheets_name_list)
        if sheet_name in sheets_name_list:
            ws = wb[sheet_name]
        elif sheet_name not in sheets_name_list:
            ws = wb.create_sheet(sheet_name)
        ws.column_dimensions['A'].width = 11
        for i in range(len(xlsx_title)):
            if i < len(xlsx_title) - 2:
                ws.cell(i + 1, 1).value = xlsx_title[i]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(xlsx_title) - 2:
                ws.cell(i + 1, 1).value = list(xlsx_title[i].keys())[0]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(xlsx_title) - 1:
                ws.cell(i + 1, 1).value = xlsx_title[i]
                ws.cell(i + 1, 1).alignment = alignment

    tp_column_numb = column_index_from_string("A") + GL.xlsx_data_interval
    all_column_numb = column_index_from_string("A") + GL.xlsx_data_interval + 1
    tv_column_numb = column_index_from_string("A") + GL.xlsx_data_interval + 2
    radio_column_numb = column_index_from_string("A") + GL.xlsx_data_interval + 3
    tp_column_char = get_column_letter(tp_column_numb)
    all_column_char = get_column_letter(all_column_numb)
    tv_column_char = get_column_letter(tv_column_numb)
    radio_column_char = get_column_letter(radio_column_numb)
    ws.column_dimensions[tp_column_char].width = 12
    ws.column_dimensions[all_column_char].width = 3
    ws.column_dimensions[tv_column_char].width = 3
    ws.column_dimensions[radio_column_char].width = 3

    for m in range(len(GL.search_datas)):
        if m < len(GL.search_datas) - 2:
            ws.cell((m + 1),(1 + GL.xlsx_data_interval)).value = GL.search_datas[m]
            ws.merge_cells(start_row=(m + 1),start_column=(1 + GL.xlsx_data_interval),\
                           end_row=(m + 1),end_column=(1 + GL.xlsx_data_interval + 4))
            ws.cell((m + 1),(1 + GL.xlsx_data_interval)).alignment = alignment
        elif m == len(GL.search_datas) - 2:
            for n in range(len(xlsx_title[7]["数据类别"])):
                ws.cell((m + 1),(1 + GL.xlsx_data_interval + n)).value = list(xlsx_title[m].values())[0][n]
                ws.cell((m + 1), (1 + GL.xlsx_data_interval + n)).alignment = alignment
                ws.row_dimensions[(m+1)].height = 13.5
        elif m == len(GL.search_datas) - 1:
            for j in range(len(GL.all_tp_list)):
                ws.cell((m+1+j),(1+GL.xlsx_data_interval)).value = GL.search_datas[m][j]
                ws.cell((m+1+j),(1+GL.xlsx_data_interval)+1).value = len(GL.channel_info[str(j+1)][0]) + \
                                                                     len(GL.channel_info[str(j+1)][1])
                ws.cell((m+1+j),(1+GL.xlsx_data_interval)+2).value = len(GL.channel_info[str(j+1)][0])
                ws.cell((m+1+j),(1+GL.xlsx_data_interval)+3).value = len(GL.channel_info[str(j+1)][1])
                ws.cell((m+1+j),(1+GL.xlsx_data_interval)+4).value = ",".join(GL.channel_info[str(j+1)][0] + \
                                                                              GL.channel_info[str(j+1)][1])
                for k in range(len(xlsx_title[7]["数据类别"])):
                    ws.cell((m+1+j),(1+GL.xlsx_data_interval)+k).alignment = alignment
                ws.row_dimensions[(m+1+j)].height = 13.5
    wb.save(write_xlsx_path)

GL = MyGlobal()

sat_search_mode_list = [
                        "6b_blind",                         #0
                        "6b_super_blind",                   #1

                        "y3_blind",                         #2
                        "y3_super_blind",                   #3

                        "138_blind",                        #4
                        "138_super_blind",                  #5

                        "88_blind",                         #6
                        "88_super_blind",                   #7

                        "plp_blind",                        #8
                        "plp_super_blind",                  #9

                        "138_incremental_blind",            #10 累加搜索

                        "y3_ch_upper_limit_blind",          #11 搜索节目达到上限,会删除所有节目,重新搜索
                        "y3_ch_ul_later_cont_blind",        #12 搜索节目达到上限后,不删除指定卫星下的tp,继续搜索
                        "y3_ch_ul_later_del_tp_blind",      #13 搜索节目达到上限后,删除指定卫星下的tp,继续搜索

                        "z6_tp_upper_limit_blind",          #14 搜索tp达到上限,会恢复出厂设置,重新搜索
                        "z6_tp_ul_later_cont_blind",        #15 搜索tp达到上限后,不删除指定卫星下的tp,继续搜索
                        "z6_tp_ul_later_del_tp_blind",      #16 搜索tp达到上限后,删除指定卫星下的tp,继续搜索

                        "reset_factory",                    #17 恢复出厂设置
                        "delete_all_channel",               #18 删除所有节目
                        ]

simplify_sat_name = {
    "Chinas6b_C":"Z6",
    "Asiasat 7 C":"Y3",
    "Telstar 18 K":"138",
    "ST 2 K":"88",
    "PLP D":"PLP D",
    "Reset":"Reset",
    "Delete":"Delete",
}

choice_search_sat = int(sys.argv[1])                        # 参考sat_list中的选项进行卫星选择

# 选择执行轮次
if len(GL.all_sat_commd[choice_search_sat]) < 9:
    search_time = GL.all_sat_commd[choice_search_sat][-1]
elif len(GL.all_sat_commd[choice_search_sat]) == 9:
    search_time = GL.all_sat_commd[choice_search_sat][7]

sat_name = GL.all_sat_commd[choice_search_sat][2][0]
search_mode = GL.all_sat_commd[choice_search_sat][2][-1]
timestamp = re.sub(r'[-: ]','_',str(datetime.now())[:19])
sheet_name = "{}_{}".format(sat_name,search_mode)

write_xlsx_file_name = "{}_{}_{}_{}_Result_{}.xlsx".format(choice_search_sat,simplify_sat_name[sat_name],sat_name,search_mode,timestamp)
write_xlsx_relative_path = r".\Result"
write_xlsx_path = os.path.join(write_xlsx_relative_path,write_xlsx_file_name)

write_txt_file_name = "{}_{}_{}_{}_{}.txt".format(choice_search_sat,simplify_sat_name[sat_name],sat_name,search_mode,timestamp)
write_txt_relative_path = r".\PrintLog"
write_txt_path = os.path.join(write_txt_relative_path,write_txt_file_name)

judge_write_file_exist()

send_ser_name,receive_ser_name = check_ports(GL.ser_cable_num)
send_ser = serial.Serial()
receive_ser = serial.Serial()
serial_set(send_ser, send_ser_name, 9600)
serial_set(receive_ser, receive_ser_name, 115200)

msg = "现在开始执行的是:{}_{}".format(sat_name,search_mode)
logging.critical(format(msg,'*^150'))

while GL.MAIN_LOOP_STATE:
    data = receive_ser.readline()

    if not data:
        # 执行单次运行的场景
        if len(GL.all_sat_commd[choice_search_sat]) < 9:
            if search_time >= 1:
                GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][0])
                if GL.commd_global_pos < GL.commd_global_length:
                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][0][GL.commd_global_pos]))
                    GL.commd_global_pos += 1
                elif GL.commd_global_pos == GL.commd_global_length:
                    search_time -= 1

            elif search_time < 1 and GL.sub_thread_delay_count_state:
                GL.sub_thread_delay_count_state = False
                delay_single_reset_factory()

        # 执行多次运行的场景
        elif len(GL.all_sat_commd[choice_search_sat]) == 9:
            # 新增卫星搜索进入天线设置界面
            if GL.switch_commd_stage == 0 and not GL.delay_cyclic_search_state:
                logging.debug("Enter Antenna Setting")
                GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][0])
                if GL.commd_global_pos < GL.commd_global_length:
                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][0][GL.commd_global_pos]))
                    # time.sleep(1)
                    GL.commd_global_pos += 1
                elif GL.commd_global_pos == GL.commd_global_length:
                    GL.switch_commd_stage += 1
                    GL.commd_global_pos = 0
                    GL.commd_global_length = 0

            elif GL.switch_commd_stage == 1:
                # 判断是否有搜索前的准备工作
                if len(GL.all_sat_commd[choice_search_sat][1]) == 0:
                    logging.debug("Not Search Preparatory Work")
                    GL.switch_commd_stage += 1

                elif len(GL.all_sat_commd[choice_search_sat][1]) > 0:
                    GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][1])
                    if GL.commd_global_pos < GL.commd_global_length:
                        GL.commd_single_length = len(GL.all_sat_commd[choice_search_sat][1][GL.commd_global_pos])
                        if GL.commd_single_pos < GL.commd_single_length and GL.delay_delete_all_sat_state:
                            preparatory_work_commd_data = GL.all_sat_commd[choice_search_sat][1][GL.commd_global_pos][GL.commd_single_pos]
                            send_ser.write(hex_strs_to_bytes(preparatory_work_commd_data))
                            GL.commd_single_pos += 1
                        elif GL.commd_single_pos == GL.commd_single_length:
                            if GL.commd_global_pos + 1 < GL.commd_global_length:    # 延迟只作用在删除卫星的list
                                GL.delay_delete_all_sat_state = False
                                delay_delete_all_sat()
                                GL.commd_single_pos = 0
                                GL.commd_global_pos += 1
                            else:
                                GL.commd_single_pos = 0
                                GL.commd_global_pos += 1
                    elif GL.commd_global_pos == GL.commd_global_length:
                        logging.debug("Search Preparatory Work")
                        GL.switch_commd_stage += 1
                        GL.commd_global_pos = 0
                        GL.commd_global_length = 0
                        GL.commd_single_pos = 0
                        GL.commd_single_length = 0
            # 进入参数配置阶段
            elif GL.switch_commd_stage == 2:
                if GL.setting_option_numb == 0:
                    logging.debug("Satellite")
                    if GL.all_sat_commd[choice_search_sat][1] == SEARCH_PREPARATORY_WORK[0]:    # upper limit or incremental search
                        if GL.sat_param_save[0] in GL.searched_sat_name:
                            logging.info("sat in list")
                            send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                            logging.info("{},{}".format(GL.sat_param_save[0],GL.searched_sat_name))

                        elif GL.sat_param_save[0] not in GL.searched_sat_name:
                            logging.info("sat not in list")
                            send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                            GL.setting_option_numb += 1
                            GL.searched_sat_name.append(GL.sat_param_save[0])
                            logging.info("{},{}".format(GL.sat_param_save[0],GL.searched_sat_name))

                    elif GL.all_sat_commd[choice_search_sat][1] == SEARCH_PREPARATORY_WORK[1]:  # normal sat search
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 1:
                    logging.debug("LNB POWER")
                    power_off = "Polar=2"
                    if GL.sat_param_save[1] != power_off and GL.switch_lnb_power_state:
                        send_ser.write(hex_strs_to_bytes(KEY["LEFT"]))
                    elif GL.sat_param_save[1] == power_off and GL.switch_lnb_power_state:
                        GL.switch_lnb_power_state = False
                        send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 2:
                    logging.debug("LBN FREQUENCY")
                    logging.info(GL.sat_param_save)
                    if GL.sat_param_save[2] != GL.all_sat_commd[choice_search_sat][2][2]:
                        send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                    elif GL.sat_param_save[2] == GL.all_sat_commd[choice_search_sat][2][2]:
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 3:
                    logging.debug("22k")
                    if GL.sat_param_save[3] != GL.all_sat_commd[choice_search_sat][2][3]:
                        send_ser.write(hex_strs_to_bytes(KEY["LEFT"]))
                    elif GL.sat_param_save[3] == GL.all_sat_commd[choice_search_sat][2][3]:
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 4:
                    logging.debug("Diseqc 1.0")
                    if GL.sat_param_save[4] != GL.all_sat_commd[choice_search_sat][2][4]:
                        send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                    elif GL.sat_param_save[4] == GL.all_sat_commd[choice_search_sat][2][4]:
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 5:
                    logging.debug("Diseqc 1.1")
                    if GL.sat_param_save[5] != GL.all_sat_commd[choice_search_sat][2][5]:
                        send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                    elif GL.sat_param_save[5] == GL.all_sat_commd[choice_search_sat][2][5]:
                        send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                        GL.setting_option_numb += 1

                elif GL.setting_option_numb == 6:
                    logging.debug("TP")
                    send_ser.write(hex_strs_to_bytes(KEY["DOWN"]))
                    GL.setting_option_numb += 1
                    GL.switch_commd_stage += 1

            elif GL.switch_commd_stage == 3 and GL.setting_option_numb == 7:
                logging.debug("Choice Search Mode And Start Search")
                for i in range(len(GL.all_sat_commd[choice_search_sat][3])):
                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][3][i]))
                    time.sleep(0.25)
                GL.setting_option_numb += 1
                # GL.switch_commd_stage += 1

            elif GL.switch_commd_stage == 4:
                logging.debug("Upper Limit To Save Channel Stage")
                if not GL.upper_limit_state:
                    logging.debug("Not Upper Limit")
                    GL.switch_commd_stage += 1
                    GL.upper_limit_to_save_channel_stage_state = True
                elif GL.upper_limit_state:
                    logging.debug("Upper Limit")
                    GL.switch_commd_stage += 1
                    delay_upper_limit_to_save_channel_stage()

            elif GL.switch_commd_stage == 5 and GL.upper_limit_to_save_channel_stage_state:
                logging.debug("Whether Or Not Save And End Search")
                GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][4])
                if GL.commd_global_pos < GL.commd_global_length:
                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][4][GL.commd_global_pos]))
                    # time.sleep(0.25)
                    GL.commd_global_pos += 1
                elif GL.commd_global_pos == GL.commd_global_length:
                    GL.commd_global_pos = 0
                    GL.commd_global_length = 0

                    if GL.all_sat_commd[choice_search_sat][4] == CHOICE_SAVE_TYPE[0]:
                        GL.delay_save_channel_state = False
                        delay_save_channel()
                        GL.switch_commd_stage += 1
                    elif GL.all_sat_commd[choice_search_sat][4] == CHOICE_SAVE_TYPE[1]:
                        GL.switch_commd_stage += 1

            elif GL.switch_commd_stage == 6 and GL.delay_save_channel_state:
                logging.debug("Write data to Excel and clear data")
                GL.search_datas[0] = sheet_name
                GL.search_datas[2] = len(GL.all_tp_list)
                GL.search_datas[3] = "{}/{}".format(GL.tv_radio_tp_count[0], GL.tv_radio_tp_count[1])
                GL.search_datas[6] = search_dur_time
                GL.search_datas[8] = GL.all_tp_list
                judge_and_wirte_data_to_xlsx()
                # 处理循环数据
                GL.all_tp_list.clear()
                GL.blind_judge_polar[0].clear()
                GL.blind_judge_polar[1].clear()
                GL.blind_judge_polar[2].clear()
                GL.channel_info.clear()
                GL.tv_radio_tp_count[0], GL.tv_radio_tp_count[1] = 0, 0
                GL.tv_radio_tp_count[4] = 0
                GL.search_datas[5] = '0/0'
                GL.tv_radio_tp_count[2], GL.tv_radio_tp_count[3] = 0, 0

                GL.switch_commd_stage += 1

            elif GL.switch_commd_stage == 7 and GL.save_ch_finish_state:
                logging.debug("Exit Antenna Setting")
                GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][5])
                if GL.commd_global_pos < GL.commd_global_length:
                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][5][GL.commd_global_pos]))
                    # time.sleep(0.25)
                    GL.commd_global_pos += 1
                elif GL.commd_global_pos == GL.commd_global_length:
                    GL.switch_commd_stage += 1
                    GL.commd_global_pos = 0
                    GL.commd_global_length = 0

            elif GL.switch_commd_stage == 8:
                if len(GL.all_sat_commd[choice_search_sat][6]) == 0:    # 没有额外操作
                    logging.debug("Not Other Operate")
                    GL.switch_commd_stage += 1
                elif len(GL.all_sat_commd[choice_search_sat][6]) > 0:   # 有额外操作
                    logging.debug("Exist Other Operate")
                    if not GL.upper_limit_state:
                        logging.debug("Exist Other Operate But Not Upper Limit")
                        GL.switch_commd_stage += 1

                    elif GL.upper_limit_state:
                        if GL.all_sat_commd[choice_search_sat][8] < 0:
                            GL.switch_commd_stage += 1
                        else:
                            if GL.all_sat_commd[choice_search_sat][6] == RESET_FACTORY:
                                logging.debug("Reset Factory")
                                GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][6])
                                if GL.commd_global_pos < GL.commd_global_length:
                                    send_ser.write(hex_strs_to_bytes(GL.all_sat_commd[choice_search_sat][6][GL.commd_global_pos]))
                                    GL.commd_global_pos += 1
                                elif GL.commd_global_pos == GL.commd_global_length:
                                    GL.delay_reset_factory_state = False
                                    delay_reset_factory()
                                    GL.switch_commd_stage += 1
                                    GL.commd_global_pos = 0
                                    GL.commd_global_length = 0
                                    GL.searched_sat_name.clear()

                            elif GL.all_sat_commd[choice_search_sat][6] == DELETE_ALL_CH:
                                logging.debug("Delete All Channels And Choice Searched First Sat")
                                if GL.other_operate_sub_stage == 0:
                                    GL.commd_global_length = len(GL.all_sat_commd[choice_search_sat][6])
                                    if GL.commd_global_pos < GL.commd_global_length:
                                        del_all_chs_commd_data = GL.all_sat_commd[choice_search_sat][6][GL.commd_global_pos]
                                        send_ser.write(hex_strs_to_bytes(del_all_chs_commd_data))
                                        GL.commd_global_pos += 1

                                    elif GL.commd_global_pos == GL.commd_global_length:
                                        GL.commd_global_pos = 0
                                        GL.commd_global_length = 0
                                        GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 1 and GL.delete_ch_finish_state:
                                    send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                                    GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 2:
                                    send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                                    GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 3:
                                    send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                                    GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 4:
                                    first_sat_name = GL.searched_sat_name[0]
                                    if GL.sat_param_save[0] == first_sat_name:
                                        GL.other_operate_sub_stage += 1
                                    elif GL.sat_param_save[0] != first_sat_name:
                                        send_ser.write(hex_strs_to_bytes(KEY["LEFT"]))

                                elif GL.other_operate_sub_stage == 5:
                                    GL.commd_global_length = len(EXIT_TO_SCREEN)
                                    if GL.commd_global_pos < GL.commd_global_length:
                                        send_ser.write(hex_strs_to_bytes(EXIT_TO_SCREEN[GL.commd_global_pos]))
                                        GL.commd_global_pos += 1
                                    elif GL.commd_global_pos == GL.commd_global_length:
                                        GL.switch_commd_stage += 1
                                        GL.commd_global_pos = 0
                                        GL.commd_global_length = 0
                                        GL.other_operate_sub_stage = 0
                                        GL.searched_sat_name.clear()

                            elif GL.all_sat_commd[choice_search_sat][6] == DELETE_SPECIFY_SAT_ALL_TP:
                                logging.debug("Delete Specify Sat TP And Choice Random Sat")
                                if GL.other_operate_sub_stage == 0:
                                    send_ser.write(hex_strs_to_bytes(KEY["MENU"]))
                                    GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 1:
                                    send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                                    GL.other_operate_sub_stage += 1

                                elif GL.other_operate_sub_stage == 2:
                                    GL.random_choice_sat.append(random.choice(GL.searched_sat_name))
                                    if GL.sat_param_save[0] == GL.random_choice_sat[0]:
                                        logging.info("{},{},{}".format(GL.sat_param_save[0], GL.random_choice_sat[0], GL.searched_sat_name))
                                        GL.other_operate_sub_stage += 1
                                        GL.searched_sat_name.remove(GL.random_choice_sat[0])
                                        GL.random_choice_sat.clear()
                                    elif GL.sat_param_save[0] != GL.random_choice_sat[0]:
                                        if PRESET_SAT_NAME.index(GL.sat_param_save[0]) > PRESET_SAT_NAME.index(GL.random_choice_sat[0]):
                                            send_ser.write(hex_strs_to_bytes(KEY["LEFT"]))
                                        elif PRESET_SAT_NAME.index(GL.sat_param_save[0]) < PRESET_SAT_NAME.index(GL.random_choice_sat[0]):
                                            send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))

                                elif GL.other_operate_sub_stage == 3:
                                    GL.commd_global_length = len(DELETE_SPECIFY_SAT_ALL_TP)
                                    if GL.commd_global_pos < GL.commd_global_length:
                                        send_ser.write(hex_strs_to_bytes(DELETE_SPECIFY_SAT_ALL_TP[GL.commd_global_pos]))
                                        GL.commd_global_pos += 1
                                    elif GL.commd_global_pos == GL.commd_global_length:
                                        GL.delay_delete_specify_sat_tp_state = False
                                        delay_delete_specify_sat_tp()
                                        GL.other_operate_sub_stage += 1
                                        GL.commd_global_pos = 0
                                        GL.commd_global_length = 0

                                elif GL.other_operate_sub_stage == 4 and GL.delay_delete_specify_sat_tp_state:
                                    GL.commd_global_length = len(EXIT_TO_SCREEN)
                                    if GL.commd_global_pos < GL.commd_global_length:
                                        send_ser.write(hex_strs_to_bytes(EXIT_TO_SCREEN[GL.commd_global_pos]))
                                        GL.commd_global_pos += 1
                                    elif GL.commd_global_pos == GL.commd_global_length:
                                        GL.switch_commd_stage += 1
                                        GL.commd_global_pos = 0
                                        GL.commd_global_length = 0
                                        GL.other_operate_sub_stage = 0

                            elif GL.all_sat_commd[choice_search_sat][6] == UPPER_LIMIT_LATER_NOT_DEL_SAT_TP_SEARCH_CONT:
                                logging.debug("Not Delete Specify Sat Tp And Search Continue")
                                send_ser.write(hex_strs_to_bytes(UPPER_LIMIT_LATER_NOT_DEL_SAT_TP_SEARCH_CONT[0]))
                                GL.switch_commd_stage += 1
                                GL.searched_sat_name.remove(random.choice(GL.searched_sat_name))  # 达到上限后切下一个卫星搜索
                                # GL.searched_sat_name.clear()        # 达到上限后重复搜索最后一个卫星

            elif GL.switch_commd_stage == 9 and GL.delay_reset_factory_state:
                logging.debug("Cyclic Search Setting")
                GL.commd_global_pos = 0
                GL.commd_global_length = 0
                GL.commd_single_pos = 0
                GL.commd_single_length = 0
                GL.delay_cyclic_search_state = True   # 控制循环搜索延时状态
                if GL.all_sat_commd[choice_search_sat][8] == NOT_UPPER_LIMIT_LATER_SEARCH_TIME:
                    if GL.delay_cyclic_search_state:
                        GL.switch_commd_stage = 0
                        GL.setting_option_numb = 0
                        delay_cyclic_search()
                        GL.switch_lnb_power_state = True
                        GL.upper_limit_state = False  # 恢复默认状态
                        GL.upper_limit_to_save_channel_stage_state = False  # 搜索上限到保存节目阶段状态恢复默认状态
                        GL.sat_param_save = ["", "", "", "", "", ""]  # 获取卫星的参数保存数据恢复默认状态
                        GL.delete_ch_finish_state = False   # 删除所有节目成功状态恢复默认
                        GL.save_ch_finish_state = False     # 保存节目成功状态恢复默认

                        search_time -= 1
                        if search_time < 1 :
                            logging.info("程序结束")
                            GL.MAIN_LOOP_STATE = False

                elif GL.all_sat_commd[choice_search_sat][8] != NOT_UPPER_LIMIT_LATER_SEARCH_TIME:
                    if GL.delay_cyclic_search_state:
                        GL.switch_commd_stage = 0
                        GL.setting_option_numb = 0
                        delay_cyclic_search()
                        GL.switch_lnb_power_state = True
                        GL.upper_limit_state = False  # 恢复默认状态
                        GL.upper_limit_to_save_channel_stage_state = False  # 搜索上限到保存节目阶段状态恢复默认状态
                        GL.sat_param_save = ["", "", "", "", "", ""]  # 获取卫星的参数保存数据恢复默认状态
                        GL.delete_ch_finish_state = False  # 删除所有节目成功状态恢复默认
                        GL.save_ch_finish_state = False  # 保存节目成功状态恢复默认

                        # GL.all_sat_commd[choice_search_sat][8] -= 1
                        if GL.all_sat_commd[choice_search_sat][8] < 0:
                            logging.info("程序结束")
                            GL.MAIN_LOOP_STATE = False

    if data:
        tt = datetime.now()
        data1 = data.decode("ISO-8859-1")
        data2 = re.compile('[\\x00-\\x08\\x0b-\\x0c\\x0e-\\x1f]').sub('', data1).strip()
        data3 = "[{}]     {}\n".format(str(tt), data2)
        # print(data2)
        add_write_data_to_txt(write_txt_path, data3)

        if GL.sat_param_kws[0] in data2:  # 判断卫星名称
            GL.sat_param_save[0] = re.split("=", data2)[-1]

        if GL.sat_param_kws[1] in data2:  # 判断LNB Fre
            lnb1 = re.split("[,\]=]", data2)[2]
            lnb2 = re.split("[,\]=]", data2)[-1]
            GL.sat_param_save[2] = "{}/{}".format(lnb1, lnb2)
        if GL.sat_param_kws[2] in data2:  # 判断22k
            GL.sat_param_save[3] = list(filter(None, re.split("-{2,}|,", data2)))[-1].strip()
        if GL.sat_param_kws[3] in data2:  # 判断diseqc 1.0和Polar(LNB Power)
            GL.sat_param_save[4] = re.split("[,\]\s]", data2)[3].split('=')[-1]  # 判断diseqc 1.0
            GL.sat_param_save[1] = re.split("[,\]\s]", data2)[-1]  # 判断LNB Power
        if GL.sat_param_kws[4] in data2:  # 判断diseqc 1.1
            GL.sat_param_save[5] = list(filter(None, re.split("-{2,}|,", data2)))[-1].strip()

        if GL.search_monitor_kws[0] in data2:  # 监控搜索起始
            start_time = datetime.now()
            GL.search_datas[1] += 1
            GL.xlsx_data_interval = 1 + 5 * (GL.search_datas[1] - 1)

        if GL.search_monitor_kws[9] in data2:  # 监控极化方向
            GL.blind_judge_polar[0].append(data2)

        if len(GL.blind_judge_polar[0]) != 0:
            if len(GL.blind_judge_polar[0]) not in GL.blind_judge_polar[1]:
                GL.blind_judge_polar[1].append(len(GL.blind_judge_polar[0]))
            elif len(GL.blind_judge_polar[0]) in GL.blind_judge_polar[1]:
                GL.blind_judge_polar[2].add(len(GL.blind_judge_polar[1]))
                if (len(GL.blind_judge_polar[2]) % 2) != 0:
                    GL.blind_judge_polar[3] = "H"
                elif (len(GL.blind_judge_polar[2]) % 2) == 0:
                    GL.blind_judge_polar[3] = "V"

        if GL.search_monitor_kws[4] in data2:  # 监控频点信息
            fre = data2.split(" ")[5]
            symb = data2.split(" ")[9]
            tp = "{}{}{}".format(fre, GL.blind_judge_polar[3], symb)
            GL.all_tp_list.append(tp)
            GL.channel_info[str(len(GL.all_tp_list))] = [[], []]

        if GL.search_monitor_kws[1] in data2:  # 监控搜索过程电视个数和名称信息
            GL.tv_radio_tp_count[0] = re.split("-{2,}|\s{2,}", data2)[1]  # 提取电视节目数
            tv_name = re.split("-{2,}|\s{2,}", data2)[2]  # 提取电视节目名称
            GL.channel_info[str(len(GL.all_tp_list))][0].append('[T]{}'.format(tv_name))

        if GL.search_monitor_kws[2] in data2:  # 监控搜索过程广播个数和名称信息
            GL.tv_radio_tp_count[1] = re.split("-{2,}|\s{2,}", data2)[1]  # 提取广播节目数
            radio_name = re.split("-{2,}|\s{2,}", data2)[2]  # 提取电视节目名称
            GL.channel_info[str(len(GL.all_tp_list))][1].append('[R]{}'.format(radio_name))

        if GL.search_monitor_kws[7] in data2 or GL.search_monitor_kws[8] in data2:  # 监控搜索达到上限

            limit_type = re.split(r"[\s_]",data2)[1]
            logging.debug(limit_type)
            logging.info("搜索{}达到上限:{}".format(limit_type,data2))

            if int(GL.tv_radio_tp_count[0]) != 0 or int(GL.tv_radio_tp_count[1]) != 0:
                send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                logging.info("搜索达到上限,发送OK成功==================================================================")
            elif int(GL.tv_radio_tp_count[0]) == 0 and int(GL.tv_radio_tp_count[1]) == 0:
                logging.info("搜索达到上限,不发送OK+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            search_time = 72
            GL.all_sat_commd[choice_search_sat][8] -= 1
            GL.upper_limit_state = True

        if GL.search_monitor_kws[3] in data2:  # 监控搜索结束
            # GL.search_end_state = True
            end_time = datetime.now()
            search_dur_time = str(end_time - start_time)[2:10]
            GL.switch_commd_stage += 1
            for i in range(len(GL.all_tp_list)):
                print(GL.all_tp_list[i])
            print("第{}次搜索节目总数为TV/Radio:{}/{},TP总数为:{},盲扫时长:{}".format(GL.search_datas[1], \
                                                                      GL.tv_radio_tp_count[0], GL.tv_radio_tp_count[1], len(GL.all_tp_list), \
                                                                      search_dur_time))

        if GL.search_monitor_kws[5] in data2:  # 监控保存TP的个数
            GL.tv_radio_tp_count[4] = int(re.split("=", data2)[1])
            GL.search_datas[4] = GL.tv_radio_tp_count[4]

        if GL.search_monitor_kws[6] in data2:  # 监控保存TV和Radio的个数
            split_result = re.split(r"[,\]]", data2)
            GL.tv_radio_tp_count[2] = re.split("=", split_result[1])[1]
            GL.tv_radio_tp_count[3] = re.split("=", split_result[2])[1]
            GL.search_datas[5] = "{}/{}".format(GL.tv_radio_tp_count[2], GL.tv_radio_tp_count[3])
            GL.save_ch_finish_state = True
            GL.tv_radio_tp_accumulated[0].append(int(GL.tv_radio_tp_count[0]))
            GL.tv_radio_tp_accumulated[1].append(int(GL.tv_radio_tp_count[1]))
            GL.tv_radio_tp_accumulated[2].append((int(len(GL.all_tp_list))))
            GL.tv_radio_tp_accumulated[3].append(GL.tv_radio_tp_count[4])

            print("本次搜索实际保存TV/Radio:{},保存TP数为:{}".format(GL.search_datas[5], GL.search_datas[4]))
            print("当前轮次:{},累计搜索节目个数:{}/{},累计搜索TP个数:{},累计保存TP个数：{}".format(GL.search_datas[1], \
                                                          sum(GL.tv_radio_tp_accumulated[0]), \
                                                          sum(GL.tv_radio_tp_accumulated[1]), \
                                                          sum(GL.tv_radio_tp_accumulated[2]), \
                                                          sum(GL.tv_radio_tp_accumulated[3])))

        if GL.delete_ch_finish_kws in data2:    # 监控删除所有节目成功的关键字
            GL.delete_ch_finish_state = True
