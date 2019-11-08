#!/usr/bin/python
# -*- coding: utf-8 -*-

from datetime import datetime
from random import choice
from threading import Timer
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font,colors,Alignment
from openpyxl.utils import get_column_letter,column_index_from_string
import serial
import serial.tools.list_ports
import logging
import re
import time
import copy
import sys
import os

class MyGlobal():
    def __init__(self):
        self.ser_cable_numb = sys.argv[2]                   # 用于指定USB转串口线的编号
        self.main_loop_state = True                         # 主程序循环状态变量
        self.send_commd_state = True                        #
        self.control_stage_delay_state = True               # 控制延时的状态变量
        self.get_group_channel_total_info_state = True      # 控制获取组别等信息的进程与用例执行代码切换的状态变量
        self.current_stage = 0                              # 控制执行用例的各个阶段
        self.sub_stage = 0                                  # 控制信息获取的各个阶段
        self.commd_global_pos = 0                           # 发送指令在list中的位置
        self.commd_global_length = 0                        # 发送指令的总长度
        self.all_group_prog_total = 0                       # All分组的节目总数
        self.TV_channel_groups = {}                         # 存放电视节目的组别和节目数信息
        self.Radio_channel_groups = {}                      # 存放广播节目的组别和节目数信息

        # [频道号,频道名称,tp,ttx,sub,lock,scramble,频道类型,视频编码,音频编码,视频高度,视频宽度]
        self.channel_info = ['','','','','','','','','','','','']
        self.prog_group_name = ''                           # 组别名称
        self.prog_group_total = ''                          # 组别下的节目总数
        self.numb_key_switch_commd = []                     # 数字键切台指令集
        self.all_test_case = []                             # 存放所有播放控制测试用例的数据集合
        self.report_data = [0,0,0,0,0,0,[],[]]


def check_ports(ser_cable_numb):
    serial_ser_value = {
        "1": "FTDVKA2HA",
        "2": "FTGDWJ64A",
        "3": "FT9SP964A",
        "4": "FTHB6SSTA"
    }
    send_port_desc = "USB-SERIAL CH340"
    receive_port_desc = "USB Serial Port"
    ports = serial.tools.list_ports.comports()
    for i in range(len(ports)):
        logging.info("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device,ports[i].description,ports[i].hwid))
    if len(ports) <= 0:
        logging.info("无可用端口")
    elif len(ports) == 1:
        logging.info("只有一个可用端口:{}".format(ports[0].device))
    elif len(ports) >=2:
        if serial.tools.list_ports.grep(send_port_desc):
            send_com = next(serial.tools.list_ports.grep(send_port_desc)).device
        if serial.tools.list_ports.grep(receive_port_desc):
            receive_com = next(serial.tools.list_ports.grep(receive_port_desc)).device
    return send_com,receive_com

def serial_set(ser, ser_name, ser_baudrate):
    ser.port = ser_name
    ser.baudrate = ser_baudrate
    ser.bytesize = 8
    ser.parity = "N"
    ser.stopbits = 1
    ser.timeout = 1
    ser.open()

def hex_strs_to_bytes(strings):
    return bytes.fromhex(strings)

def write_logs_to_txt(file_path,logs):
    with open(file_path, "a+", encoding="utf-8") as fo:
        fo.write(logs)

def check_if_log_and_report_file_path_exists():
    global case_log_file_directory, full_log_file_directory, report_file_directory
    parent_path = os.path.dirname(os.getcwd())
    case_log_folder_name = "print_log"
    case_log_file_directory = os.path.join(parent_path, case_log_folder_name)
    full_log_folder_name = "print_log_full"
    full_log_file_directory = os.path.join(parent_path, full_log_folder_name)
    report_folder_name = "report"
    report_file_directory = os.path.join(parent_path, report_folder_name)

    if not os.path.exists(case_log_file_directory):
        os.mkdir(case_log_file_directory)
    if not os.path.exists(full_log_file_directory):
        os.mkdir(full_log_file_directory)
    if not os.path.exists(report_file_directory):
        os.mkdir(report_file_directory)

def build_print_log_and_report_file_path():
    global case_log_txt_path, full_log_txt_path, report_file_path, sheet_name
    case_info = ALL_TEST_CASE[choice_switch_case]
    time_info = re.sub(r"[-: ]", "_", str(datetime.now())[:19])
    case_log_file_name = "{}_{}_{}_{}.txt".format(case_info[0], case_info[1], case_info[2], time_info)
    case_log_txt_path = os.path.join(case_log_file_directory, case_log_file_name)
    full_log_file_name = "full_{}_{}_{}_{}.txt".format(case_info[0], case_info[1], case_info[2], time_info)
    full_log_txt_path = os.path.join(full_log_file_directory, full_log_file_name)
    report_file_name = "{}_{}_{}_{}.xlsx".format(case_info[0], case_info[1], case_info[2], time_info)
    report_file_path = os.path.join(report_file_directory, report_file_name)
    sheet_name = "{}".format(case_info[1])

    GL.report_data[0] = "{}_{}_{}".format(case_info[0], case_info[1], case_info[2])
    GL.report_data[1] = "{}".format(case_info[3])
    GL.report_data[3] = "{}".format(case_info[4])
    GL.report_data[4] = "{}".format(case_info[2])

def check_if_report_exists_and_write_data_to_report():
    report_title = [
        "report name",
        "group name",
        "group ch total",
        "ch type",
        "switch ch mode",
        "switch ch times",
        {"command": CH_INFO_KWS},
    ]

    alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
    if not os.path.exists(report_file_path):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.column_dimensions['A'].width = 16
        for i in range(len(report_title)):
            ws.row_dimensions[(i + 1)].height = 13.5
            if i < len(report_title) - 1:
                ws.cell(i + 1, 1).value = report_title[i]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(report_title) - 1:
                ws.cell(i + 1, 1).value = list(report_title[i].keys())[0]
                ws.cell(i + 1, 1).alignment = alignment
                for j in range(len(report_title[i]["command"])):
                    all_column_numb = column_index_from_string("A") + (j + 1)
                    all_column_char = get_column_letter(all_column_numb)
                    ws.column_dimensions[all_column_char].width = 16
                    ws.cell((i + 1), (1 + j + 1)).value = report_title[i]["command"][j]
                    ws.cell((i + 1), (1 + j + 1)).alignment = alignment

    elif os.path.exists(report_file_path):
        wb = load_workbook(report_file_path)
        sheets_name_list = wb.sheetnames
        logging.info(sheets_name_list)
        if sheet_name in sheets_name_list:
            ws = wb[sheet_name]
        elif sheet_name not in sheets_name_list:
            ws = wb.create_sheet(sheet_name)
        ws.column_dimensions['A'].width = 16
        for i in range(len(report_title)):
            ws.row_dimensions[(i + 1)].height = 13.5
            if i < len(report_title) - 1:
                ws.cell(i + 1, 1).value = report_title[i]
                ws.cell(i + 1, 1).alignment = alignment
            elif i == len(report_title) - 1:
                ws.cell(i + 1, 1).value = list(report_title[i].keys())[0]
                ws.cell(i + 1, 1).alignment = alignment
                for j in range(len(report_title[i]["command"])):
                    all_column_numb = column_index_from_string("A") + j
                    all_column_char = get_column_letter(all_column_numb)
                    ws.column_dimensions[all_column_char].width = 16
                    ws.cell((i + 1), (1 + j + 1)).value = report_title[i]["command"][j]
                    ws.cell((i + 1), (1 + j + 1)).alignment = alignment
    wb.save(report_file_path)

def delay_time(interval_time,expect_delay_time):
    global t
    expect_delay_time -= interval_time
    print(expect_delay_time)
    if expect_delay_time == 0:
        GL.control_stage_delay_state = True
    if expect_delay_time > 0:
        t = Timer(interval_time, delay_time, args=(interval_time, expect_delay_time)).start()

def build_ch_numbs_list(chs_total):
    # 新建指定个数的数值列表
    # ch_numbs = GL.channel_numbers
    ch_numbs_list = []
    for i in range(chs_total):
        ch_numbs_list.append(str(i + 1))
    return ch_numbs_list

def build_random_ch_commds_list(commds_list,number_of_time):
    # 在某指令集中随机抽取指定个数的指令集
    random_channel_commds_list = []
    for i in range(number_of_time):
        random_channel_commds_list.append(choice(commds_list))
    return random_channel_commds_list

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

def change_commds_to_numbs_list(commds_list):
    # 将指令集列表转化为数值列表
    channel_numbs_list = []
    for i in range(len(commds_list)):
        channel_numbs_list.append([])
        for j in range(len(commds_list[i])):
            channel_numbs_list[i].append(REVERSE_KEY[commds_list[i][j]])
    return channel_numbs_list

def change_numbs_to_str_list(numbs_list):
    # 将数值列表转化为字符串
    channel_str_list = []
    for m in range(len(numbs_list)):
        channel_str_list.append("")
        for n in range(len(numbs_list[m])):
            try:
                int(numbs_list[m][n])
            except:
                channel_str_list[m] = channel_str_list[m] + "_" + numbs_list[m][n]
            else:
                channel_str_list[m] = channel_str_list[m] + numbs_list[m][n]
    return channel_str_list

def commds_add_key_list(old_commds_list,key_name):
    # 在每个指令集后增加指定的单一指令
    new_commds_list = copy.deepcopy(old_commds_list)
    for i in range(len(new_commds_list)):
        new_commds_list[i].append(key_name)
    return new_commds_list

def build_single_commds_list(single_commd,number_of_time):
    # 创建指定次数的单一指令集
    single_commds_list = []
    for i in range(number_of_time):
        single_commds_list.append([])
        single_commds_list[i].append(single_commd)
    return single_commds_list

def build_random_move_focus_list(single_commd,number_of_time):
    # 创建指定次数的随机个数单一指令集
    random_mv_focus_list = []
    single_commd_list = [single_commd]
    for i in range(number_of_time):
        random_mv_focus_list.append(single_commd_list * random.randint(1,20))
    return random_mv_focus_list

def commds_add_random_move_focus_list(old_commds_list,single_commd):
    # 在已有的commds_list追加随机个数的单一指令集
    new_commds_list = copy.deepcopy(old_commds_list)
    # single_commd_list = [single_commd]
    for i in range(len(new_commds_list)):
        for j in range(random.randint(1,5)):
            new_commds_list[i].append(single_commd)
    return new_commds_list

def build_tv_numb_key_switch_list():
    scene_list = ["numb", "numb+ok", "numb+random"]
    random_switch_time = 1000
    numb_key_total_commd = []
    GL.all_group_prog_total = int(GL.TV_channel_groups["All"])
    chs_numb_list = build_ch_numbs_list(GL.all_group_prog_total)

    numb_key_chs_commd_list = change_numbs_to_commds_list(chs_numb_list)
    numb_key_chs_commd_add_ok_list = commds_add_key_list(numb_key_chs_commd_list,KEY["OK"])
    numb_key_random_chs_commd_list = build_random_ch_commds_list(numb_key_chs_commd_list,random_switch_time)

    numb_key_total_commd.append(numb_key_chs_commd_list)
    numb_key_total_commd.append(numb_key_chs_commd_add_ok_list)
    numb_key_total_commd.append(numb_key_random_chs_commd_list)
    return numb_key_total_commd

def build_all_test_case():
    GL.all_test_case = [
                        [ ALL_TEST_CASE[0],NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[0], EXIT_TO_SCREEN],

                        [ ALL_TEST_CASE[1],NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[1], EXIT_TO_SCREEN],

                        [ ALL_TEST_CASE[2],NUMB_KEY_PREPARATORY_WORK,
                          GL.numb_key_switch_commd[2], EXIT_TO_SCREEN],
    ]

def build_all_scene_commd_list():
    GL.numb_key_switch_commd = build_tv_numb_key_switch_list()

if __name__ == "__main__":
    LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s %(message)s"  # 配置输出日志的格式
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S %a"  # 配置输出时间的格式
    logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT)

    KEY = {
        "0": "A1 F1 22 DD 00",
        "1": "A1 F1 22 DD 01",
        "2": "A1 F1 22 DD 02",
        "3": "A1 F1 22 DD 03",
        "4": "A1 F1 22 DD 04",
        "5": "A1 F1 22 DD 05",
        "6": "A1 F1 22 DD 06",
        "7": "A1 F1 22 DD 07",
        "8": "A1 F1 22 DD 08",
        "9": "A1 F1 22 DD 09",
        "OK": "A1 F1 22 DD 15",
        "UP": "A1 F1 22 DD 11",
        "DOWN": "A1 F1 22 DD 14",
        "LEFT": "A1 F1 22 DD 12",
        "RIGHT": "A1 F1 22 DD 13",
        "PAGE_UP": "A1 F1 22 DD 41",
        "PAGE_DOWN": "A1 F1 22 DD 18",
        "MENU": "A1 F1 22 DD 0C",
        "EXIT": "A1 F1 22 DD 0D",
        "EPG": "A1 F1 22 DD 0E",
        "YELLOW": "A1 F1 22 DD 1B",
        "TV/R": "A1 F1 22 DD 42",
    }

    REVERSE_KEY = dict([val,key] for key,val in KEY.items())

    SWITCH_CHANNEL_KWS = [
        "[PTD]Prog_numb=",
        "[PTD]TP=",
        "[PTD]video_height=",
        "[PTD]Group_name="]

    CH_INFO_KWS = [
        "Prog_numb",
        "Prog_name",
        "TP",
        "TTX_flag",
        "SUB_flag",
        "Lock_flag",
        "Scramble_flag",
        "Prog_type",
        "V_codec",
        "A_codec",
        "video_height",
        "video_width",]

    GROUP_INGO_KWS = [
        "Group_name",
        "Prog_total"
    ]

    NUMB_KEY_PREPARATORY_WORK = []
    EXIT_TO_SCREEN = [KEY["EXIT"], KEY["EXIT"], KEY["EXIT"]]
    INTERVAL_TIME = [1.0, 5.0]

    ALL_TEST_CASE = [
        ["screen", "numb_key", "one_by_one_timeout", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "numb_key", "one_by_one_add_ok", "All", "TV", INTERVAL_TIME[1]],
        ["screen", "numb_key", "random", "All", "TV", INTERVAL_TIME[1]],
    ]

    GL = MyGlobal()
    # 指定切台用例
    choice_switch_case = int(sys.argv[1])

    # 检查打印日志和报告的目录,以及创建文件名称

    check_if_log_and_report_file_path_exists()
    build_print_log_and_report_file_path()
    check_if_report_exists_and_write_data_to_report()

    # 获取串口并配置串口信息
    send_ser_name,receive_ser_name = check_ports(GL.ser_cable_numb)
    send_ser = serial.Serial()
    receive_ser = serial.Serial()
    serial_set(send_ser,send_ser_name,9600)
    serial_set(receive_ser,receive_ser_name,115200)


    while GL.main_loop_state:
        data = receive_ser.readline()
        if data:
            tt = datetime.now()
            # data1 = data.decode("ISO-8859-1", "replace")
            # data2 = data.decode('ISO-8859-1', 'replace').replace('\ufffd', '')
            data1 = data.decode("GB18030", "replace")
            data2 = data.decode('GB18030', 'replace').replace('\ufffd', '').strip()
            data3 = re.compile('[\\x00-\\x08\\x0b-\\x0c\\x0e-\\x1f]').sub('', data2).strip()
            data4 = "[{}]     {}\n".format(str(tt), data2)
            print(data2)
            write_logs_to_txt(full_log_txt_path, data4)
            if not GL.get_group_channel_total_info_state:
                write_logs_to_txt(case_log_txt_path, data4)

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
                    if CH_INFO_KWS[2] in flag_info_split[i]:    # 提取频道所属TP
                        GL.channel_info[2] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[3] in flag_info_split[i]:    # 提取频道TTX_flag
                        GL.channel_info[3] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[4] in flag_info_split[i]:    # 提取频道SUB_flag
                        GL.channel_info[4] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[5] in flag_info_split[i]:    # 提取频道Lock_flag
                        GL.channel_info[5] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[6] in flag_info_split[i]:    # 提取频道Scramble_flag
                        GL.channel_info[6] = re.split(r"=", flag_info_split[i])[-1]
                    if CH_INFO_KWS[7] in flag_info_split[i]:    # 提取频道类别:TV/Radio
                        GL.channel_info[7] = re.split(r"=", flag_info_split[i])[-1]
                    if GL.channel_info[7] == "TV":
                        if CH_INFO_KWS[8] in flag_info_split[i]:    # 提取TV频道视频编码
                            GL.channel_info[8] = re.split(r"=", flag_info_split[i])[-1]
                        if CH_INFO_KWS[9] in flag_info_split[i]:    # 提取TV频道音频编码
                            GL.channel_info[9] = re.split(r"=", flag_info_split[i])[-1]
                    if GL.channel_info[7] == "Radio":
                        GL.channel_info[8] = "0"                    # 指定Radio频道视频编码为空
                        if CH_INFO_KWS[9] in flag_info_split[i]:    # 提取Radio频道音频编码
                            GL.channel_info[9] = re.split(r"=", flag_info_split[i])[-1]

            if SWITCH_CHANNEL_KWS[2] in data2:
                wide_high_info_split = re.split(r"[\],]", data2)
                for i in range(len(wide_high_info_split)):
                    if CH_INFO_KWS[10] in wide_high_info_split[i]:  # 提取TV频道画面高度
                        GL.channel_info[10] = re.split(r"=", wide_high_info_split[i])[-1]
                    if CH_INFO_KWS[11] in wide_high_info_split[i]:  # 提取TV频道画面宽度
                        GL.channel_info[11] = re.split(r"=", wide_high_info_split[i])[-1]

            if SWITCH_CHANNEL_KWS[3] in data2:
                group_info_split = re.split(r"[\],]", data2)
                for i in range(len(group_info_split)):
                    if GROUP_INGO_KWS[0] in group_info_split[i]:    # 提取频道所属组别
                        GL.prog_group_name = re.split(r"=", group_info_split[i])[-1]
                    if GROUP_INGO_KWS[1] in group_info_split[i]:    # 提取频道所属组别下的节目总数
                        GL.prog_group_total = re.split(r"=", group_info_split[i])[-1]

        if GL.get_group_channel_total_info_state:
            if not data:
                if GL.sub_stage == 0:   # 切台用于判断当前节目类别属性(TV/Radio)
                    send_ser.write(hex_strs_to_bytes(KEY["UP"]))
                    GL.sub_stage += 1
                elif GL.sub_stage == 1: # 默认切换到TV节目界面
                    if GL.channel_info[7] == "TV":
                        GL.sub_stage += 1
                    elif GL.channel_info[7] == "Radio":
                        send_ser.write(hex_strs_to_bytes(KEY["TV/R"]))
                elif GL.sub_stage == 2: # 调出频道列表,用于判断组别信息
                    send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                    GL.sub_stage += 1
                elif GL.sub_stage == 3: # 默认切换到TV的All分组
                    if GL.prog_group_name != "All":
                        send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                        if GL.prog_group_name not in GL.TV_channel_groups.keys():
                            GL.TV_channel_groups[GL.prog_group_name] = GL.prog_group_total
                    elif GL.prog_group_name == "All":
                        if GL.prog_group_name not in GL.TV_channel_groups.keys():
                            GL.TV_channel_groups[GL.prog_group_name] = GL.prog_group_total
                        GL.sub_stage += 1
                elif GL.sub_stage == 4: # 退出频道列表,回到大画面界面
                    send_ser.write(hex_strs_to_bytes(KEY["EXIT"]))
                    GL.sub_stage += 1
                elif GL.sub_stage == 5: # 创建所有场景的测试用例,测试数据,打印文件保存目录
                    GL.sub_stage == 0
                    logging.info(GL.TV_channel_groups.keys())
                    logging.info(GL.Radio_channel_groups.keys())
                    build_all_scene_commd_list()
                    build_all_test_case()
                    # build_print_log_and_report_file_path()
                    logging.info(GL.all_test_case[choice_switch_case][2])
                    GL.get_group_channel_total_info_state = False
                # if GL.sub_stage == 0:
                #     send_ser.write(hex_strs_to_bytes(KEY["UP"]))
                #     GL.sub_stage += 1
                # elif GL.sub_stage == 1:
                #     send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                #     GL.sub_stage += 1
                # elif GL.sub_stage == 2:
                #     if GL.channel_info[7] == "TV":
                #         if GL.prog_group_name not in GL.TV_channel_groups.keys():
                #             send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                #             GL.TV_channel_groups[GL.prog_group_name] = GL.prog_group_total
                #         elif GL.prog_group_name in GL.TV_channel_groups.keys():
                #             GL.sub_stage += 1
                #     elif GL.channel_info[7] == "Radio":
                #         if GL.prog_group_name not in GL.Radio_channel_groups.keys():
                #             send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                #             GL.Radio_channel_groups[GL.prog_group_name] = GL.prog_group_total
                #         elif GL.prog_group_name in GL.Radio_channel_groups.keys():
                #             GL.sub_stage += 1
                # elif GL.sub_stage == 3:
                #     send_ser.write(hex_strs_to_bytes(KEY["EXIT"]))
                #     GL.sub_stage += 1
                # elif GL.sub_stage == 4:
                #     send_ser.write(hex_strs_to_bytes(KEY["TV/R"]))
                #     GL.sub_stage += 1
                # elif GL.sub_stage == 5:
                #     send_ser.write(hex_strs_to_bytes(KEY["OK"]))
                #     GL.sub_stage += 1
                # elif GL.sub_stage == 6:
                #     send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                #     if GL.channel_info[7] == "Radio":
                #         if GL.prog_group_name not in GL.Radio_channel_groups.keys():
                #             GL.Radio_channel_groups[GL.prog_group_name] = GL.prog_group_total
                #         elif GL.prog_group_name in GL.Radio_channel_groups.keys():
                #             GL.sub_stage += 1
                #     elif GL.channel_info[7] == "TV":
                #         if GL.prog_group_name not in GL.TV_channel_groups.keys():
                #             GL.TV_channel_groups[GL.prog_group_name] = GL.prog_group_total
                #         elif GL.prog_group_name in GL.TV_channel_groups.keys():
                #             GL.sub_stage += 1
                # elif GL.sub_stage == 7:
                #     if GL.prog_group_name != "All":
                #         send_ser.write(hex_strs_to_bytes(KEY["RIGHT"]))
                #     elif GL.prog_group_name == "All":
                #         GL.sub_stage += 1
                # elif GL.sub_stage == 8:
                #     send_ser.write(hex_strs_to_bytes(KEY["EXIT"]))
                #     if GL.channel_info[7] == "Radio":
                #         send_ser.write(hex_strs_to_bytes(KEY["TV/R"]))
                #     elif GL.channel_info[7] == "TV":
                #         GL.sub_stage += 1
                # elif GL.sub_stage == 9:
                #     GL.get_group_channel_total_info_state = False
                #     GL.sub_stage == 0
                #     logging.info(GL.TV_channel_groups.keys())
                #     logging.info(GL.Radio_channel_groups.keys())
                #     build_all_scene_commd_list()
                #     build_all_test_case()

        elif not GL.get_group_channel_total_info_state:
            if GL.current_stage == 0:   # 判断节目类型:TV/Radio
                if GL.channel_info[7] != GL.all_test_case[choice_switch_case][0][4]:
                    send_ser.write(hex_strs_to_bytes(KEY["TV/R"]))
                elif GL.channel_info[7] == GL.all_test_case[choice_switch_case][0][4]:
                    GL.current_stage += 1

            elif GL.current_stage == 1: # 判断是否有准备工作,比如进入某界面
                if len(GL.all_test_case[choice_switch_case][1]) == 0:
                    GL.current_stage += 1
                elif len(GL.all_test_case[choice_switch_case][1]) > 0:
                    GL.commd_global_length = len(GL.all_test_case[choice_switch_case][1])
                    if GL.commd_global_pos < GL.commd_global_length:
                        preparatory_work_commd = GL.all_test_case[choice_switch_case][1]
                        send_ser.write(hex_strs_to_bytes(preparatory_work_commd[GL.commd_global_pos]))
                        GL.commd_global_pos += 1
                    elif GL.commd_global_pos == GL.commd_global_length:
                        GL.current_stage += 1
                        GL.commd_global_length = 0
                        GL.commd_global_pos = 0

            elif GL.current_stage == 2: # 发送所选用例的切台指令
                GL.commd_global_length = len(GL.all_test_case[choice_switch_case][2])
                if GL.commd_global_pos < GL.commd_global_length and GL.control_stage_delay_state:
                    GL.control_stage_delay_state = False
                    send_data = GL.all_test_case[choice_switch_case][2][GL.commd_global_pos]
                    for i in range(len(send_data)):
                        send_ser.write(hex_strs_to_bytes(send_data[i]))
                        time.sleep(0.2)
                    if GL.all_test_case[choice_switch_case][0][-1] == INTERVAL_TIME[0]:
                        delay_time(0.5, INTERVAL_TIME[0])
                    elif GL.all_test_case[choice_switch_case][0][-1] == INTERVAL_TIME[1]:
                        delay_time(1.0,INTERVAL_TIME[1])
                    GL.commd_global_pos += 1
                elif GL.commd_global_pos == GL.commd_global_length:
                    GL.current_stage += 1
                    GL.commd_global_length = 0
                    GL.commd_global_pos = 0

            elif GL.current_stage == 3: # 发送切台后的退回大画面的指令
                GL.commd_global_length = len(GL.all_test_case[choice_switch_case][3])
                if GL.commd_global_pos < GL.commd_global_length and GL.control_stage_delay_state:
                    GL.control_stage_delay_state = False
                    exit_commd = GL.all_test_case[choice_switch_case][3][GL.commd_global_pos]
                    send_ser.write(hex_strs_to_bytes(exit_commd))
                    GL.commd_global_pos += 1
                    delay_time(0.5, INTERVAL_TIME[0])
                elif GL.commd_global_pos == GL.commd_global_length:
                    GL.current_stage += 1
                    GL.commd_global_length = 0
                    GL.commd_global_pos = 0
            elif GL.current_stage == 4: # 写报告
                if ALL_TEST_CASE[choice_switch_case][4] == "TV":
                    GL.report_data[2] = int(GL.TV_channel_groups["All"])
                elif ALL_TEST_CASE[choice_switch_case][4] == "Radio":
                    GL.report_data[2] = int(GL.Radio_channel_groups["All"])

                GL.report_data[5] = len(GL.all_test_case[choice_switch_case][2])
                # GL.report_data[6] =
                check_if_report_exists_and_write_data_to_report()

            elif GL.current_stage == 5: # 结束程序

                GL.main_loop_state = False
