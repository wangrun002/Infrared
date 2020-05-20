#!/usr/bin/python3
# -*- coding: utf-8 -*-

from serial_setting import *
from multiprocessing import Process, Manager
from datetime import datetime
from random import randint
import os
import time
import logging
import re



def logging_info_setting():
    # 配置logging输出格式
    LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s %(message)s"  # 配置输出日志的格式
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S %a"  # 配置输出时间的格式
    logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT)


def hex_strs_to_bytes(strings):
    # 将红外命令字符串转换为字节串
    return bytes.fromhex(strings)


def write_log_data_to_txt(file_path, write_data):
    with open(file_path, "a+", encoding="utf-8") as fo:
        fo.write(write_data)


def build_send_and_receive_serial():
    # 创建发送和接受串口
    ser_name = list(check_ports())  # send_ser_name, receive_ser_name
    send_ser = serial.Serial(ser_name[0], 9600)
    receive_ser = serial.Serial(ser_name[1], 115200, timeout=1)
    return send_ser, receive_ser


def send_commd(commd):
    # 红外发送端发送指令
    send_serial.write(hex_strs_to_bytes(commd))
    send_serial.flush()
    logging.info("红外发送：{}".format(REVERSE_KEY[commd]))
    infrared_send_cmd.append(REVERSE_KEY[commd])
    time.sleep(1.0)


def send_more_commds(commd_list):
    # 用于发送一连串的指令
    for commd in commd_list:
        send_commd(commd)
    time.sleep(1)   # 增加函数切换时的的等待，避免可能出现send_commd函数中的等待时间没有执行的情况


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


def build_log_and_report_file_path():
    # 用于创建打印和报告文件路径
    # 构建存放数据的总目录，以及构建存放打印和报告的目录
    parent_path = os.path.dirname(os.getcwd())
    test_data_directory_name = "test_data"
    test_data_directory_path = os.path.join(parent_path, test_data_directory_name)
    log_directory_name = "print_log"
    log_directory_path = os.path.join(test_data_directory_path, log_directory_name)
    report_directory_name = "report"
    report_directory_path = os.path.join(test_data_directory_path, report_directory_name)
    # 判断目录是否存在，否则创建目录
    if not os.path.exists(test_data_directory_path):
        os.mkdir(test_data_directory_path)
    if not os.path.exists(log_directory_path):
        os.mkdir(log_directory_path)
    if not os.path.exists(report_directory_path):
        os.mkdir(report_directory_path)
    # 创建打印和报告文件的名称和路径
    time_info = re.sub(r"[-: ]", "_", str(datetime.now())[:19])
    fmt_name = "{}_{}_{}_{}_event_{}_triggered".format(
        TEST_CASE_INFO[0], TEST_CASE_INFO[1], TEST_CASE_INFO[2],
        TEST_CASE_INFO[3], TEST_CASE_INFO[4], TEST_CASE_INFO[5])
    log_file_name = "Log_{}_{}.txt".format(fmt_name, time_info)
    log_file_path = os.path.join(log_directory_path, log_file_name)
    report_file_name = "{}_{}.xlsx".format(fmt_name, time_info)
    report_file_path = os.path.join(report_directory_path, report_file_name)
    sheet_name = "{}_{}".format(TEST_CASE_INFO[2], TEST_CASE_INFO[4])
    return log_file_path, report_file_path, sheet_name


def clear_timer_setting_all_events():
    # 清除Timer_setting界面所有的事件
    enter_timer_setting_interface = [KEY["MENU"], KEY["LEFT"], KEY["DOWN"], KEY["OK"]]
    delete_all_res_events = [KEY["BLUE"], KEY["OK"]]
    exit_to_screen = [KEY["EXIT"], KEY["EXIT"], KEY["EXIT"]]
    # 进入定时器设置界面
    send_more_commds(enter_timer_setting_interface)
    # 对定时器设置界面的事件判断和清除
    while not state["res_event_numb_state"]:
        logging.info("还没有获取到预约事件个数")
        time.sleep(1)
    else:
        if rsv_kws["res_event_numb"] != 0:
            send_more_commds(delete_all_res_events)
        elif rsv_kws["res_event_numb"] == 0:
            logging.info("没有预约事件存在")
            time.sleep(1)
        else:
            logging.debug("警告：预约事件个数获取错误！！！")
        state["res_event_numb_state"] = False
    # 退回大画面
    send_more_commds(exit_to_screen)


def check_sys_time_mode():
    # 检测系统事件模式
    enter_time_setting_interface = [KEY["MENU"], KEY["LEFT"], KEY["OK"]]
    change_sys_time_mode = [KEY["RIGHT"], KEY["EXIT"], KEY["OK"]]
    exit_to_screen = [KEY["EXIT"], KEY["EXIT"], KEY["EXIT"]]
    # 进入时间设置界面
    send_more_commds(enter_time_setting_interface)
    # 对当前系统时间模式进行判断
    while not state["sys_time_mode_state"]:
        logging.info("还没有获取到系统时间模式信息")
        time.sleep(1)
    else:
        if rsv_kws["sys_time_mode"] == "Auto":
            send_more_commds(change_sys_time_mode)
        elif rsv_kws["sys_time_mode"] == "Manual":
            logging.info("系统时间模式已经为手动模式")
        else:
            logging.debug("警告：系统时间模式获取错误！！！")
        state["sys_time_mode_state"] = False
    # 退回大画面
    send_more_commds(exit_to_screen)


def get_current_system_time():
    # 获取当前系统时间
    while not state["current_sys_time_state"]:
        logging.info("还没有获取到系统时间信息")
        time.sleep(1)
    else:
        logging.info(f"当前系统时间为:{rsv_kws['current_sys_time']}")


def get_exist_event_info():
    # 获取当前已预约的事件信息
    if len(res_event_list) == 0:
        logging.info("当前没有预约事件")
        logging.info(f"预约事件列表为:{res_event_list}")
    else:
        logging.info("当前所有预约事件如下")
        for event in res_event_list:
            logging.info(event)


def choice_ch_for_res_event_type():
    # 根据预约事件类型来选择节目
    group_dict = {}
    choice_ch_numb = []
    # 根据所选case切换到对应类型节目的界面
    while channel_info[5] != TEST_CASE_INFO[2]:
        send_commd(KEY["TV/R"])
        if channel_info[3] == "1":
            send_commd(KEY["EXIT"])
    # 调出频道列表,用于判断组别信息
    send_commd(KEY["OK"])
    # 切到指定分组和采集分组下节目总数信息
    while rsv_kws["prog_group_name"] != TEST_CASE_INFO[1]:
        send_commd(KEY["RIGHT"])
        if channel_info[3] == "1":
            send_commd(KEY["EXIT"])
    else:
        if rsv_kws["prog_group_name"] == '':
            logging.info("警告：没有All分组信息")
        else:
            group_dict[rsv_kws["prog_group_name"]] = rsv_kws["prog_group_total"]
            logging.info(f"分组信息{group_dict}")
    # 退出频道列表,回到大画面界面
    send_commd(KEY["EXIT"])
    # 根据用例指定的事件类型来选择节目
    if TEST_CASE_INFO[4] == "Play":
        choice_ch_numb.append(randint(1, int(group_dict[TEST_CASE_INFO[1]])))
        choice_ch_cmd = change_numbs_to_commds_list(choice_ch_numb)
        for i in range(len(choice_ch_cmd)):
            for j in choice_ch_cmd[i]:
                send_commd(j)
        send_commd(KEY["OK"])
        time.sleep(2)
        if channel_info[3] == "1":
            send_commd(KEY["EXIT"])
        logging.info(f"所选节目频道号和所切到的节目频道号为:{choice_ch_numb}--{channel_info[0]}")
        logging.info(channel_info)


def calculate_expected_event_start_time():
    # 计算期望的预约事件的时间
    expected_res_time = ['', '', '', '', '']        # 期望的预约事件时间信息[年，月，日，时，分]
    swap_data = [0, 0, 0, 0, 0]                     # 用于交换处理的时间信息[年，月，日，时，分]
    time_interval = 2
    leap_year_month = 29
    nonleap_year_month = 28
    solar_month = [1, 3, 5, 7, 8, 10, 12]
    lunar_month = [4, 6, 9, 11]
    sys_time = rsv_kws['current_sys_time']
    sys_time_split = re.split(r"[\s:/]", sys_time)
    sys_year = int(sys_time_split[0])
    sys_month = int(sys_time_split[1])
    sys_day = int(sys_time_split[2])
    sys_hour = int(sys_time_split[3])
    sys_minute = int(sys_time_split[4])

    swap_data[4] = sys_minute + time_interval
    # 计算分钟和进位到小时
    if swap_data[4] < 60:
        expected_res_time[4] = "{0:02d}".format(swap_data[4])
        swap_data[3] = sys_hour
    elif swap_data[4] >= 60:
        expected_res_time[4] = "{0:02d}".format(swap_data[4] - 60)
        swap_data[3] = sys_hour + 1
    # 计算小时和进位到天数
    if swap_data[3] < 24:
        expected_res_time[3] = "{0:02d}".format(swap_data[3])
        swap_data[2] = sys_day
    elif swap_data[3] >= 24:
        expected_res_time[3] = "{0:02d}".format(swap_data[3] - 24)
        swap_data[2] = sys_day + 1
    # 按照闰年、平年，月份来计算天数和是否进位月份
    if sys_month == 2:
        logging.info("当前月份为二月")
        if sys_year % 100 == 0 and sys_year % 400 == 0:
            logging.info("当前年份为世纪闰年，二月有29天")
            if swap_data[2] <= leap_year_month:
                expected_res_time[2] = "{0:02d}".format(swap_data[2])
                expected_res_time[1] = "{0:02d}".format(sys_month)
                expected_res_time[0] = "{0:02d}".format(sys_year)
            elif swap_data[2] > leap_year_month:
                expected_res_time[2] = "{0:02d}".format(swap_data[2] - leap_year_month)
                expected_res_time[1] = "{0:02d}".format(sys_month + 1)
                expected_res_time[0] = "{0:02d}".format(sys_year)
        elif sys_year % 100 != 0 and sys_year % 4 == 0:
            logging.info("当前年份为普通闰年，二月有29天")
            if swap_data[2] <= leap_year_month:
                expected_res_time[2] = "{0:02d}".format(swap_data[2])
                expected_res_time[1] = "{0:02d}".format(sys_month)
                expected_res_time[0] = "{0:02d}".format(sys_year)
            elif swap_data[2] > leap_year_month:
                expected_res_time[2] = "{0:02d}".format(swap_data[2] - leap_year_month)
                expected_res_time[1] = "{0:02d}".format(sys_month + 1)
                expected_res_time[0] = "{0:02d}".format(sys_year)
        else:
            logging.info("当前年份为平年，二月有28天")
            if sys_month == 2:
                if swap_data[2] <= nonleap_year_month:
                    expected_res_time[2] = "{0:02d}".format(swap_data[2])
                    expected_res_time[1] = "{0:02d}".format(sys_month)
                    expected_res_time[0] = "{0:02d}".format(sys_year)
                elif swap_data[2] > nonleap_year_month:
                    expected_res_time[2] = "{0:02d}".format(swap_data[2] - nonleap_year_month)
                    expected_res_time[1] = "{0:02d}".format(sys_month + 1)
                    expected_res_time[0] = "{0:02d}".format(sys_year)

    elif sys_month in solar_month:
        logging.info("当前月份为大月，大月有31天")
        if swap_data[2] <= 31:
            expected_res_time[2] = "{0:02d}".format(swap_data[2])
            expected_res_time[1] = "{0:02d}".format(sys_month)
            expected_res_time[0] = "{0:02d}".format(sys_year)
        elif swap_data[2] > 31:
            expected_res_time[2] = "{0:02d}".format(swap_data[2] - 31)
            swap_data[1] = sys_month + 1
            if swap_data[1] <= 12:
                expected_res_time[1] = "{0:02d}".format(swap_data[1])
                expected_res_time[0] = "{0:02d}".format(sys_year)
            elif swap_data[1] > 12:
                expected_res_time[1] = "{0:02d}".format(swap_data[1] - 12)
                expected_res_time[0] = "{0:02d}".format(sys_year + 1)

    elif sys_month in lunar_month:
        logging.info("当前月份为小月，小月有30天")
        if swap_data[2] <= 30:
            expected_res_time[2] = "{0:02d}".format(swap_data[2])
            expected_res_time[1] = "{0:02d}".format(sys_month)
            expected_res_time[0] = "{0:02d}".format(sys_year)
        elif swap_data[2] > 30:
            expected_res_time[2] = "{0:02d}".format(swap_data[2] - 30)
            swap_data[1] = sys_month + 1
            if swap_data[1] <= 12:
                expected_res_time[1] = "{0:02d}".format(swap_data[1])
                expected_res_time[0] = "{0:02d}".format(sys_year)
            elif swap_data[1] > 12:
                expected_res_time[1] = "{0:02d}".format(swap_data[1] - 12)
                expected_res_time[0] = "{0:02d}".format(sys_year + 1)

    str_expected_res_time = ''.join(expected_res_time)
    logging.info(f"期望的完整的预约事件时间为{str_expected_res_time}")
    return str_expected_res_time


def create_expected_event_info():
    # 创建期望的事件信息
    expected_event_info = ['', '', '', '', '']      # [起始时间，事件响应类型，节目名称，持续时间，事件触发模式]
    if TEST_CASE_INFO[4] == "Play":
        choice_ch_for_res_event_type()
        if TEST_CASE_INFO[3] == "Once":
            expected_event_full_time = calculate_expected_event_start_time()
            expected_event_info[0] = expected_event_full_time
            expected_event_info[1] = TEST_CASE_INFO[4]
            expected_event_info[2] = channel_info[1]
            expected_event_info[3] = "--:--"
            expected_event_info[4] = TEST_CASE_INFO[3]

    elif TEST_CASE_INFO[4] == "PVR":
        pass
    elif TEST_CASE_INFO[4] == "Power Off":
        pass
    return expected_event_info


def edit_res_event_info():
    # 生成预期的预约事件
    expected_res_event_info = create_expected_event_info()
    # 编辑预约事件信息


def add_res_event():
    # 新增预约事件
    # 进入Timer_Setting界面
    enter_timer_setting_interface = [KEY["MENU"], KEY["LEFT"], KEY["DOWN"], KEY["OK"]]
    send_more_commds(enter_timer_setting_interface)
    # 获取当前系统时间
    get_current_system_time()

    # 进入事件编辑界面，设置预约事件参数
    send_commd(KEY["GREEN"])





def receive_serial_process(
        prs_data, infrared_send_cmd, rsv_kws, res_event_list, state, current_triggered_event_info, channel_info):
    logging_info_setting()
    rsv_key = {
        "POWER": "0xbbaf", "TV/R": "0xbbbd", "MUTE": "0xbbf7",
        "1": "0xbb7f", "2": "0xbbbf", "3": "0xbb3f",
        "4": "0xbbdf", "5": "0xbb5f", "6": "0xbb9f",
        "7": "0xbb1f", "8": "0xbbef", "9": "0xbb6f",
        "FAV": "0xbb87", "0": "0xbbff", "SAT": "0xbb97",
        "MENU": "0xbbcf", "EPG": "0xbb8f", "INFO": "0xbb07", "EXIT": "0xbb4f",
        "UP": "0xbb77", "DOWN": "0xbbd7",
        "LEFT": "0xbbb7", "RIGHT": "0xbb37", "OK": "0xbb57",
        "P/N": "0xbb0f", "SLEEP": "0xbb17", "PAGE_UP": "0xbb7d", "PAGE_DOWN": "0xbbe7",
        "RED": "0xbb67", "GREEN": "0xbba7", "YELLOW": "0xbb27", "BLUE": "0xbbc7",
        "F1": "0xbb9d", "F2": "0xbb5d", "F3": "0xbbdd", "RECALL": "0xbb3d",
        "REWIND": "0xbb47", "FF": "0xbb1d", "PLAY": "0xbb2f", "RECORD": "0xbbfd",
        "PREVIOUS": "0xbbad", "NEXT": "0xbb6d", "TIMESHIFT": "0xbbed", "STOP": "0xbb4d"
    }
    reverse_rsv_key = dict([val, key] for key, val in rsv_key.items())

    res_kws = [
        "[PTD]Time_mode=",          # 0     获取系统时间模式
        "[PTD]System_time=",        # 1     系统时间
        "[PTD]Res_event_numb=",     # 2     预约事件数量
        "[PTD]Res_event:",          # 3     预约事件信息
        "[PTD]Res_triggered:",      # 4     预约事件触发和当前响应事件的信息
        "[PTD]Res_confirm_jump",    # 5     预约事件确认跳转
        "[PTD]Res_cancel_jump",     # 6     预约事件取消跳转
        "[PTD]REC_start",           # 7     录制开始
        "[PTD]REC_end",             # 8     录制结束
        "[PTD]No_storage_device",   # 9     没有存储设备
        "[PTD]No_enough_space"      # 10    没有足够的空间
    ]

    switch_ch_kws = [
        "[PTD]Prog_numb=",
        "[PTD]TP=",
        "[PTD]video_height=",
        "[PTD]Group_name=",
        "Swtich Video interval"]

    ch_info_kws = [
        "Prog_numb",
        "Prog_name",
        "TP",
        "Lock_flag",
        "Scramble_flag",
        "Prog_type",
        "Group_name"]

    group_info_kws = [
        "Group_name",
        "Prog_total"
    ]

    edit_event_kws = [
        "[PTD]Mode=",
        "[PTD]Type=",
        "[PTD]Start Date=",
        "[PTD]Start Time=",
        "[PTD]Duration=",
        "[PTD]=Channel="
    ]

    other_kws = [
        "[PTD]Infrared_key_values:",    # 获取红外接收关键字
    ]

    infrared_rsv_cmd = []
    receive_serial = prs_data["receive_serial"]

    while True:
        data = receive_serial.readline()
        if data:
            tt = datetime.now()
            data1 = data.decode("GB18030", "ignore")
            data2 = re.compile('[\\x00-\\x08\\x0b-\\x0c\\x0e-\\x1f]').sub('', data1).strip()
            data3 = "[{}]     {}\n".format(str(tt), data2)
            print(data2)
            write_log_data_to_txt(prs_data["log_file_path"], data3)

            if other_kws[0] in data2:   # 红外接收打印
                rsv_cmd = re.split(":", data2)[-1]
                infrared_rsv_cmd.append(rsv_cmd)
                if rsv_cmd not in reverse_rsv_key.keys():
                    logging.info("红外键值{}不在当前字典中，被其他遥控影响".format(rsv_cmd))
                else:
                    logging.info("红外键值(发送和接受):({})--({})".format(
                        infrared_send_cmd[-1], reverse_rsv_key[infrared_rsv_cmd[-1]]))
                    logging.info("红外次数统计(发送和接受):{}--{}".format(
                        len(infrared_send_cmd), len(infrared_rsv_cmd)))

            if res_kws[0] in data2:     # 获取系统时间模式（自动还是手动）
                state["sys_time_mode_state"] = True
                rsv_kws["sys_time_mode"] = re.split(r"=", data2)[-1]

            if res_kws[1] in data2:     # 获取当前系统时间
                state["current_sys_time_state"] = True
                rsv_kws["current_sys_time"] = re.split(r"=", data2)[-1]

            if res_kws[2] in data2:     # 获取预约事件数量
                state["res_event_numb_state"] = True
                rsv_kws["res_event_numb"] = re.split(r"=", data2)[-1]

            if res_kws[3] in data2:     # 获取预约事件信息
                event_split_info = re.split(r"t:|,", data2)
                event_info = ['', '', '', '', '']
                for info in event_split_info:
                    if "Start_time" in info:
                        event_info[0] = re.split(r"=", info)[-1]
                    if "Event_type" in info:
                        event_info[1] = re.split(r"=", info)[-1]
                    if "Ch_name" in info:
                        event_info[2] = re.split(r"=", info)[-1]
                    if "Duration" in info:
                        event_info[3] = re.split(r"=", info)[-1]
                    if "Event_mode" in info:
                        event_info[4] = re.split(r"=", info)[-1]
                res_event_list.append(event_info)

            if res_kws[4] in data2:     # 获取预约事件跳转触发信息，以及当前响应事件的信息
                state["res_event_triggered_state"] = False
                current_event_split_info = re.split(r"d:|,", data2)
                current_event_info = ['', '', '', '', '']
                for info in current_event_split_info:
                    if "Start_time" in info:
                        current_event_info[0] = re.split(r"=", info)[-1]
                    if "Event_type" in info:
                        current_event_info[1] = re.split(r"=", info)[-1]
                    if "Ch_name" in info:
                        current_event_info[2] = re.split(r"=", info)[-1]
                    if "Duration" in info:
                        current_event_info[3] = re.split(r"=", info)[-1]
                    if "Event_mode" in info:
                        current_event_info[4] = re.split(r"=", info)[-1]
                current_triggered_event_info.extend(current_event_info)

            if res_kws[5] in data2:     # 获取预约事件确认跳转信息
                state["res_event_confirm_jump_state"] = True

            if res_kws[6] in data2:     # 获取预约事件取消跳转信息
                state["res_event_cancel_jump_state"] = True

            if res_kws[7] in data2:     # 获取PVR预约事件录制开始信息
                state["rec_start_state"] = True

            if res_kws[8] in data2:     # 获取PVR预约事件录制结束信息
                state["rec_end_state"] = True

            if res_kws[9] in data2:     # 存储设备没有插入的打印信息
                state["no_storage_device_state"] = True

            if res_kws[10] in data2:    # 存储设备没有足够空间的打印信息
                state["no_enough_space_state"] = True

            if switch_ch_kws[0] in data2:
                ch_info_split = re.split(r"[\],]", data2)
                for i in range(len(ch_info_split)):
                    if ch_info_kws[0] in ch_info_split[i]:  # 提取频道号
                        channel_info[0] = re.split("=", ch_info_split[i])[-1]
                    if ch_info_kws[1] in ch_info_split[i]:  # 提取频道名称
                        channel_info[1] = re.split("=", ch_info_split[i])[-1]

            if switch_ch_kws[1] in data2:
                flag_info_split = re.split(r"[\],]", data2)
                for i in range(len(flag_info_split)):
                    if ch_info_kws[2] in flag_info_split[i]:  # 提取频道所属TP
                        channel_info[2] = re.split(r"=", flag_info_split[i])[-1].replace(" ", "")
                    if ch_info_kws[3] in flag_info_split[i]:  # 提取频道Lock_flag
                        channel_info[3] = re.split(r"=", flag_info_split[i])[-1]
                    if ch_info_kws[4] in flag_info_split[i]:  # 提取频道Scramble_flag
                        channel_info[4] = re.split(r"=", flag_info_split[i])[-1]
                    if ch_info_kws[5] in flag_info_split[i]:  # 提取频道类别:TV/Radio
                        channel_info[5] = re.split(r"=", flag_info_split[i])[-1]

            if switch_ch_kws[3] in data2:
                group_info_split = re.split(r"[\],]", data2)
                for i in range(len(group_info_split)):
                    if group_info_kws[0] in group_info_split[i]:  # 提取频道所属组别
                        rsv_kws["prog_group_name"] = re.split(r"=", group_info_split[i])[-1]
                        channel_info[6] = rsv_kws["prog_group_name"]
                    if group_info_kws[1] in group_info_split[i]:  # 提取频道所属组别下的节目总数
                        rsv_kws["prog_group_total"] = re.split(r"=", group_info_split[i])[-1]

            if edit_event_kws[0] in data2:
                rsv_kws["edit_event_focus_pos"] = "Mode"
                rsv_kws["edit_event_mode"] = re.split(r"=", data2)[-1]

            if edit_event_kws[1] in data2:
                rsv_kws["edit_event_focus_pos"] = "Type"
                rsv_kws["edit_event_type"] = re.split(r"=", data2)[-1]

            if edit_event_kws[2] in data2:
                rsv_kws["edit_event_focus_pos"] = "Start Date"
                rsv_kws["edit_event_date"] = re.split(r"=", data2)[-1]

            if edit_event_kws[3] in data2:
                rsv_kws["edit_event_focus_pos"] = "Start Time"
                rsv_kws["edit_event_time"] = re.split(r"=", data2)[-1]

            if edit_event_kws[4] in data2:
                rsv_kws["edit_event_focus_pos"] = "Duration"
                rsv_kws["edit_event_duration"] = re.split(r"=", data2)[-1]

            if edit_event_kws[5] in data2:
                rsv_kws["edit_event_focus_pos"] = "Channel"
                rsv_kws["edit_event_ch"] = re.split(r"=", data2)[-1]





if __name__ == "__main__":
    logging_info_setting()

    KEY = {
        "POWER": "A1 F1 22 DD 0A", "TV/R": "A1 F1 22 DD 42", "MUTE": "A1 F1 22 DD 10",
        "1": "A1 F1 22 DD 01", "2": "A1 F1 22 DD 02", "3": "A1 F1 22 DD 03",
        "4": "A1 F1 22 DD 04", "5": "A1 F1 22 DD 05", "6": "A1 F1 22 DD 06",
        "7": "A1 F1 22 DD 07", "8": "A1 F1 22 DD 08", "9": "A1 F1 22 DD 09",
        "FAV": "A1 F1 22 DD 1E", "0": "A1 F1 22 DD 00", "SAT": "A1 F1 22 DD 16",
        "MENU": "A1 F1 22 DD 0C", "EPG": "A1 F1 22 DD 0E", "INFO": "A1 F1 22 DD 1F", "EXIT": "A1 F1 22 DD 0D",
        "UP": "A1 F1 22 DD 11", "DOWN": "A1 F1 22 DD 14",
        "LEFT": "A1 F1 22 DD 12", "RIGHT": "A1 F1 22 DD 13", "OK": "A1 F1 22 DD 15",
        "P/N": "A1 F1 22 DD 0F", "SLEEP": "A1 F1 22 DD 17", "PAGE_UP": "A1 F1 22 DD 41", "PAGE_DOWN": "A1 F1 22 DD 18",
        "RED": "A1 F1 22 DD 19", "GREEN": "A1 F1 22 DD 1A", "YELLOW": "A1 F1 22 DD 1B", "BLUE": "A1 F1 22 DD 1C",
        "F1": "A1 F1 22 DD 46", "F2": "A1 F1 22 DD 45", "F3": "A1 F1 22 DD 44", "RECALL": "A1 F1 22 DD 43",
        "REWIND": "A1 F1 22 DD 1D", "FF": "A1 F1 22 DD 47", "PLAY": "A1 F1 22 DD 0B", "RECORD": "A1 F1 22 DD 40",
        "PREVIOUS": "A1 F1 22 DD 4A", "NEXT": "A1 F1 22 DD 49", "TIME_SHIFT": "A1 F1 22 DD 48", "STOP": "A1 F1 22 DD 4D"
    }
    REVERSE_KEY = dict([val, key] for key, val in KEY.items())
    TEST_CASE_INFO = ["23", "All", "TV", "Once", "Play", "Screen"]

    file_path = build_log_and_report_file_path()
    serial_object = build_send_and_receive_serial()
    send_serial = serial_object[0]

    infrared_send_cmd = Manager().list([])
    res_event_list = Manager().list([])
    current_triggered_event_info = Manager().list([])
    channel_info = Manager().list(['', '', '', '', '', '', ''])     # [频道号,频道名称,tp,lock,scramble,频道类型,组别]
    rsv_kws = Manager().dict({
        "sys_time_mode": '', "current_sys_time": '', "res_event_numb": '', "prog_group_name": '',
        "prog_group_total": '', "edit_event_focus_pos": '', "edit_event_mode": '', "edit_event_type": '',
        "edit_event_date": '', "edit_event_time": '', "edit_event_duration": '', "edit_event_ch": ''
    })

    state = Manager().dict({
        "res_event_numb_state": False, "res_event_triggered_state": False, "res_event_confirm_jump_state": False,
        "res_event_cancel_jump_state": False, "rec_start_state": False, "rec_end_state": False,
        "no_storage_device_state": False, "no_enough_space_state": False, "sys_time_mode_state": False,
        "current_sys_time_state": False
    })

    prs_data = Manager().dict({
        "log_file_path": file_path[0], "receive_serial": serial_object[1],
    })

    rsv_p = Process(target=receive_serial_process, args=(
        prs_data, infrared_send_cmd, rsv_kws, res_event_list, state, current_triggered_event_info, channel_info))
    rsv_p.start()

    clear_timer_setting_all_events()
    check_sys_time_mode()
    choice_ch_for_res_event_type()
    add_res_event()