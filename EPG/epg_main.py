#!/usr/bin/python3
# -*- coding: utf-8 -*-

from datetime import datetime
from random import sample
import serial
import serial.tools.list_ports
import logging
import platform
import threading
import time
import re


class MyGlobal():
    def __init__(self):
        # [频道号,频道名称,tp,lock,scramble,频道类型,组别,epg_info]
        self.channel_info = ['', '', '', '', '', '', '', '']
        self.prog_group_name = ''                               # 组别名称
        self.prog_group_total = ''                              # 组别下的节目总数
        self.TV_channel_groups = {}                             # 存放电视节目的组别和节目数信息
        self.Radio_channel_groups = {}                          # 存放广播节目的组别和节目数信息
        self.main_loop_state = True
        self.current_stage = 0                                  # 控制执行用例的各个阶段
        self.sub_stage = 0                                      # 控制信息获取的各个阶段
        self.TV_ch_attribute = [[], [], []]                     # 用于存放TV节目属性的列表(免费\加密\加锁)
        self.Radio_ch_attribute = [[], [], []]                  # 用于存放Radio节目属性的列表(免费\加密\加锁)

def check_ports():
    global send_com,receive_com
    ports_info = []
    if platform.system() == "Windows":
        ser_cable_num = 4
        serial_ser = {
            "1": "FTDVKA2HA",
            "2": "FTGDWJ64A",
            "3": "FT9SP964A",
            "4": "FTHB6SSTA",
            "5": "FTDVKPRSA",
            "6": "FTHI8UIHA",
             }
        send_port_desc = "USB-SERIAL CH340"
        receive_port_desc = serial_ser[str(ser_cable_num)]
    elif platform.system() == "Linux":
        ser_cable_num = 5
        serial_ser = {
            "1": "FTDVKA2H",
            "2": "FTGDWJ64",
            "3": "FT9SP964",
            "4": "FTHB6SST",
            "5": "FTDVKPRS",
            "6": "FTHI8UIH",
            }
        send_port_desc = "USB2.0-Serial"
        receive_port_desc = serial_ser[str(ser_cable_num)]
    ports = list(serial.tools.list_ports.comports())
    for i in range(len(ports)):
        logging.info("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device, ports[i].description, ports[i].hwid))
        #print("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device, ports[i].description, ports[i].hwid))
        ports_info.append("{}~{}~{}".format(ports[i].device, ports[i].description, ports[i].hwid))
    if len(ports) <= 0:
        logging.info("无可用端口")
    elif len(ports) == 1:
        logging.info("只有一个可用端口:{}".format(ports[0].device))
    elif len(ports) >= 2:
        for i in range(len(ports_info)):
            if send_port_desc in ports_info[i]:
                send_com = ports_info[i].split("~")[0]
            elif receive_port_desc in ports_info[i]:
                receive_com = ports_info[i].split("~")[0]
    return send_com,receive_com

def hex_strs_to_bytes(strings):
    return bytes.fromhex(strings)

def send_commd(commd):
    send_ser.write(hex_strs_to_bytes(commd))
    time.sleep(1)

def send_numb_key_commd(commd):
    send_ser.write(hex_strs_to_bytes(commd))
    time.sleep(0.5)

def write_logs_to_txt(file_path,logs):
    with open(file_path, "a+", encoding="utf-8") as fo:
        fo.write(logs)

def change_numbs_to_commds_list(numbs_list):
    # 将数值列表转换为指令集列表
    channel_commds_list = []
    for i in range(len(numbs_list)):
        channel_commds_list.append([])
        if len(numbs_list[i]) == 1:
            channel_commds_list[i].append(KEY[numbs_list[i]])
        elif len(numbs_list[i]) > 1:
            for j in range(len(numbs_list[i])):
                channel_commds_list[i].append(KEY[numbs_list[i][j]])
    return channel_commds_list

def get_group_channel_total_info():
    # 切台前获取case节目类别,分组,分组节目数量,以及获取节目属性前的去除加锁的判断
    if GL.sub_stage == 0:  # 切台用于判断当前节目类别属性(TV/Radio)
        # 根据所选case切换到对应类型节目的界面
        logging.debug("GL.sub_stage == 0")
        while GL.channel_info[5] != TEST_CASE[1]:
            send_commd(KEY["TV/R"])
            if GL.channel_info[3] == "1":
                send_commd(KEY["EXIT"])
        GL.sub_stage += 1
    elif GL.sub_stage == 1:  # 调出频道列表,用于判断组别信息
        logging.debug("GL.sub_stage == 1")
        send_commd(KEY["OK"])
        GL.sub_stage += 1
    elif GL.sub_stage == 2: # 采集所有分组的名称和分组下节目总数信息
        logging.debug("GL.sub_stage == 2")
        if TEST_CASE[1] == "TV":
            while GL.prog_group_name not in GL.TV_channel_groups.keys():
                print(GL.prog_group_name)
                GL.TV_channel_groups[GL.prog_group_name] = GL.prog_group_total
                send_commd(KEY["RIGHT"])
                if GL.channel_info[3] == "1":
                    send_commd(KEY["EXIT"])
            GL.sub_stage += 1
        elif TEST_CASE[1] == "Radio":
            while GL.prog_group_name not in GL.Radio_channel_groups.keys():
                GL.Radio_channel_groups[GL.prog_group_name] = GL.prog_group_total
                send_commd(KEY["RIGHT"])
                if GL.channel_info[3] == "1":
                    send_commd(KEY["EXIT"])
            GL.sub_stage += 1
    elif GL.sub_stage == 3:  # 根据所选case切换到对应的分组
        logging.debug("GL.sub_stage == 3")
        if TEST_CASE[1] == "TV":
            while GL.prog_group_name != TEST_CASE[0]:
                send_commd(KEY["RIGHT"])
                if GL.channel_info[3] == "1":
                    send_commd(KEY["EXIT"])
            GL.sub_stage += 1

        elif TEST_CASE[1] == "Radio":
            while GL.prog_group_name != TEST_CASE[0]:
                send_commd(KEY["RIGHT"])
                if GL.channel_info[3] == "1":
                    send_commd(KEY["EXIT"])
            GL.sub_stage += 1
    elif GL.sub_stage == 4:  # 退出频道列表,回到大画面界面
        logging.debug("GL.sub_stage == 4")
        send_commd(KEY["EXIT"])
        logging.debug(GL.channel_info)
        logging.debug(GL.TV_channel_groups)
        logging.debug(GL.Radio_channel_groups)
        GL.sub_stage += 1
        GL.current_stage += 1

def check_ch_type():
    while GL.channel_info[5] != TEST_CASE[1]:
        send_commd(KEY["TV/R"])
        if GL.channel_info[3] == "1":
            for i in range(4):
                send_numb_key_commd(KEY["0"])
    GL.current_stage += 1

def check_preparatory_work():
    if isinstance(TEST_CASE_COMMD[1], str):
        GL.current_stage += 1
    elif isinstance(TEST_CASE_COMMD[1], list):
        send_data = TEST_CASE_COMMD[1]
        for i in range(len(send_data)):
            send_commd(send_data[i])
        if GL.channel_info[3] == "1":
            for i in range(4):
                send_numb_key_commd(KEY["0"])
        GL.current_stage += 1

def get_group_all_ch_type(choice_group_ch_total_numb):
    send_commd(KEY["EPG"])
    for i in range(int(choice_group_ch_total_numb)):
        GL.channel_info = ['', '', '', '', '', '', GL.prog_group_name, '']
        send_numb_key_commd(KEY["DOWN"])
        time.sleep(1)
        if GL.channel_info[3] == "1":
            for i in range(4):
                send_numb_key_commd(KEY["0"])
        if TEST_CASE[1] == "TV":
            if GL.channel_info[3] == "1":  # 加锁电视节目
                GL.TV_ch_attribute[2].append(GL.channel_info[0])
            elif GL.channel_info[4] == "0":  # 免费电视节目
                GL.TV_ch_attribute[0].append(GL.channel_info[0])
            elif GL.channel_info[4] == "1":  # 加密电视节目
                GL.TV_ch_attribute[1].append(GL.channel_info[0])
        elif TEST_CASE[1] == "Radio":
            if GL.channel_info[5] == "1":  # 加锁广播节目
                GL.Radio_ch_attribute[2].append(GL.channel_info[0])
            elif GL.channel_info[6] == "0":  # 免费广播节目
                GL.Radio_ch_attribute[0].append(GL.channel_info[0])
            elif GL.channel_info[6] == "1":  # 加密广播节目
                GL.Radio_ch_attribute[1].append(GL.channel_info[0])
        logging.info(GL.channel_info)
    logging.info(GL.TV_ch_attribute)
    logging.info(GL.Radio_ch_attribute)
    GL.current_stage += 1

def choice_test_channel():
    if TEST_CASE[1] == "TV":
        if "Free" == TEST_CASE[2]:
            if len(GL.TV_ch_attribute[0]) == 0:
                logging.info("没有免费电视节目")
                GL.main_loop_state = False
            elif len(GL.TV_ch_attribute[0]) > 0:
                free_tv_numb = sample(GL.TV_ch_attribute[0], 1)
                logging.debug("所选免费电视节目为:{}".format(free_tv_numb))
                free_tv_commd = change_numbs_to_commds_list(free_tv_numb)
                send_commd(KEY["EXIT"])
                for i in range(len(free_tv_commd)):
                    for j in range(len(free_tv_commd[i])):
                        send_numb_key_commd(free_tv_commd[i][j])
                send_commd(KEY["OK"])
        GL.current_stage += 1
    elif TEST_CASE[1] == "Radio":
        if "Free" == TEST_CASE[2]:
            if len(GL.Radio_ch_attribute[0]) == 0:
                logging.info("没有免费广播节目")
                GL.main_loop_state = False
            elif len(GL.Radio_ch_attribute[0]) > 0:
                free_radio_numb = sample(GL.Radio_ch_attribute[0], 1)
                logging.debug("所选免费广播节目为:{}".format(free_radio_numb))
                free_radio_commd = change_numbs_to_commds_list(free_radio_numb)
                send_commd(KEY["EXIT"])
                for i in range(len(free_radio_commd)):
                    for j in range(len(free_radio_commd[i])):
                        send_numb_key_commd(free_radio_commd[i][j])
                send_commd(KEY["OK"])
        GL.current_stage += 1

def data_send_thread():
    while GL.main_loop_state:
        if GL.current_stage == 0:
            get_group_channel_total_info()

        elif GL.current_stage == 1:  # 采集All分组下的节目属性和是否有EPG信息
            if TEST_CASE[1] == "TV":
                get_group_all_ch_type(GL.TV_channel_groups[TEST_CASE[0]])
            elif TEST_CASE[1] == "Radio":
                get_group_all_ch_type(GL.Radio_channel_groups[TEST_CASE[0]])

        elif GL.current_stage == 2:  # 判断节目类型:TV/Radio
            check_ch_type()

        elif GL.current_stage == 3:  # 判断是否有准备工作,比如进入某界面
            check_preparatory_work()

        elif GL.current_stage == 4:  # 选择并切到符合条件的节目
            choice_test_channel()

        elif GL.current_stage == 5:  # 发送所选用例的切台指令
            pass

def data_receiver_thread():
    while GL.main_loop_state:
        data = receive_ser.readline()
        if data:
            tt = datetime.now()
            data1 = data.decode("GB18030", "ignore")
            data2 = re.compile('[\\x00-\\x08\\x0b-\\x0c\\x0e-\\x1f]').sub('', data1).strip()
            data4 = "[{}]     {}\n".format(str(tt), data2)
            print(data2)

            if SWITCH_CHANNEL_KWS[0] in data2:
                ch_info_split = re.split(r"[\],]", data2)
                for i in range(len(ch_info_split)):
                    if CH_INFO_KWS[0] in ch_info_split[i]:  # 提取频道号
                        GL.channel_info[0] = re.split("=", ch_info_split[i])[-1]
                    if CH_INFO_KWS[1] in ch_info_split[i]:  # 提取频道名称
                        GL.channel_info[1] = re.split("=", ch_info_split[i])[-1]

            if SWITCH_CHANNEL_KWS[1] in data2:
                flag_info_split = re.split(r"[\],]", data2)
                for i in range(len(flag_info_split)):
                    if CH_INFO_KWS[2] in flag_info_split[i]:  # 提取频道所属TP
                        GL.channel_info[2] = re.split(r"=", flag_info_split[i])[-1].replace(" ", "")
                    if CH_INFO_KWS[3] in flag_info_split[i]:  # 提取频道Lock_flag
                        GL.channel_info[3] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[4] in flag_info_split[i]:  # 提取频道Scramble_flag
                        GL.channel_info[4] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[5] in flag_info_split[i]:  # 提取频道类别:TV/Radio
                        GL.channel_info[5] = re.split(r"=", flag_info_split[i])[-1]

            if SWITCH_CHANNEL_KWS[3] in data2:
                group_info_split = re.split(r"[\],]", data2)
                for i in range(len(group_info_split)):
                    if GROUP_INFO_KWS[0] in group_info_split[i]:  # 提取频道所属组别
                        GL.prog_group_name = re.split(r"=", group_info_split[i])[-1]
                        GL.channel_info[6] = GL.prog_group_name
                    if GROUP_INFO_KWS[1] in group_info_split[i]:  # 提取频道所属组别下的节目总数
                        GL.prog_group_total = re.split(r"=", group_info_split[i])[-1]

if __name__ == "__main__":
    KEY = {
        "POWER": "A1 F1 22 DD 0A", "TV/R": "A1 F1 22 DD 42", "MUTE": "A1 F1 22 DD 10",
        "1": "A1 F1 22 DD 01", "2": "A1 F1 22 DD 02", "3": "A1 F1 22 DD 03",
        "4": "A1 F1 22 DD 04", "5": "A1 F1 22 DD 05", "6": "A1 F1 22 DD 06",
        "7": "A1 F1 22 DD 07", "8": "A1 F1 22 DD 08", "9": "A1 F1 22 DD 09",
        "FAV": "A1 F1 22 DD 1E", "0": "A1 F1 22 DD 00", "ZOOM": "A1 F1 22 DD 16",
        "MENU": "A1 F1 22 DD 0C", "EPG": "A1 F1 22 DD 0E", "INFO": "A1 F1 22 DD 1F", "EXIT": "A1 F1 22 DD 0D",
        "UP": "A1 F1 22 DD 11", "DOWN": "A1 F1 22 DD 14",
        "LEFT": "A1 F1 22 DD 12", "RIGHT": "A1 F1 22 DD 13", "OK": "A1 F1 22 DD 15",
        "P/N": "A1 F1 22 DD 0F", "R/L": "A1 F1 22 DD 17", "PAGE_UP": "A1 F1 22 DD 41", "PAGE_DOWN": "A1 F1 22 DD 18",
        "RED": "A1 F1 22 DD 19", "GREEN": "A1 F1 22 DD 1A", "YELLOW": "A1 F1 22 DD 1B", "BLUE": "A1 F1 22 DD 1C",
        "FIND": "A1 F1 22 DD 46", "PAUSE": "A1 F1 22 DD 45", "SUB": "A1 F1 22 DD 44", "RECALL": "A1 F1 22 DD 43",
        "REWIND": "A1 F1 22 DD 1D", "FF": "A1 F1 22 DD 47", "PLAY": "A1 F1 22 DD 0B", "RECORD": "A1 F1 22 DD 40",
        "PREVIOUS": "A1 F1 22 DD 4A", "NEXT": "A1 F1 22 DD 49", "TIMESHIFT": "A1 F1 22 DD 48", "STOP": "A1 F1 22 DD 4D",

        "CH+": "A1 F1 22 DD 11", "CH-": "A1 F1 22 DD 14", "VOL-": "A1 F1 22 DD 12", "VOL+": "A1 F1 22 DD 13",
        "MULTIFEED": "A1 F1 22 DD 0F", "SAT": "A1 F1 22 DD 17", "AUDIO": "A1 F1 22 DD 19", "TTX": "A1 F1 22 DD 1C",
    }

    SWITCH_CHANNEL_KWS = [
        "[PTD]Prog_numb=",
        "[PTD]TP=",
        "[PTD]video_height=",
        "[PTD]Group_name=",
        "Swtich Video interval"]

    CH_INFO_KWS = [
        "Prog_numb",
        "Prog_name",
        "TP",
        "Lock_flag",
        "Scramble_flag",
        "Prog_type",
        "Group_name", ]

    GROUP_INFO_KWS = [
        "Group_name",
        "Prog_total"
    ]

    PREPARATORY_WORK = 'no_preparatory_work'
    ENTER_EXIT_EPG_COMMD = [KEY["EPG"], KEY["EXIT"]]
    EXIT_TO_SCREEN = [KEY["EXIT"], KEY["EXIT"], KEY["EXIT"]]

    TEST_CASE = ["All", "TV", "Free", "epg_info_flag"]
    TEST_CASE_COMMD = [TEST_CASE, PREPARATORY_WORK, ENTER_EXIT_EPG_COMMD, EXIT_TO_SCREEN]

    GL = MyGlobal()

    # 配置日志信息
    LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s %(message)s"  # 配置输出日志的格式
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S %a"  # 配置输出时间的格式
    logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT)

    # 获取串口并配置串口信息
    send_com = ''
    receive_com = ''
    send_ser_name, receive_ser_name = check_ports()
    send_ser = serial.Serial(send_ser_name, 9600)
    receive_ser = serial.Serial(receive_ser_name, 115200, timeout=1)

    thread_send = threading.Thread(target=data_send_thread)
    thread_receive = threading.Thread(target=data_receiver_thread)
    thread_receive.start()
    thread_send.start()