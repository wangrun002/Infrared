#!/usr/bin/python3
# -*- coding: utf-8 -*-

import serial
import serial.tools.list_ports
import platform
import logging

ser_cable_num = 7


def check_ports():
    send_com, receive_com = '', ''
    send_port_desc, receive_port_desc = '', ''
    ports_info = []
    if platform.system() == "Windows":
        serial_ser = {
            "1": "FTDVKA2HA",
            "2": "FTGDWJ64A",
            "3": "FT9SP964A",
            "4": "FTHB6SSTA",
            "5": "FTDVKPRSA",
            "6": "FTHI8UIHA",
            "7": "FTHG05TTA",
        }
        send_port_desc = "USB-SERIAL CH340"
        receive_port_desc = serial_ser[str(ser_cable_num)]
    elif platform.system() == "Linux":
        serial_ser = {
            "1": "FTDVKA2H",
            "2": "FTGDWJ64",
            "3": "FT9SP964",
            "4": "FTHB6SST",
            "5": "FTDVKPRS",
            "6": "FTHI8UIH",
            "7": "FTHG05TT",
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
    return send_com, receive_com


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
        "PREVIOUS": "A1 F1 22 DD 4A", "NEXT": "A1 F1 22 DD 49", "TIMESHIFT": "A1 F1 22 DD 48", "STOP": "A1 F1 22 DD 4D",
    }
REVERSE_KEY = dict([val, key] for key, val in KEY.items())

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

all_test_case = [
    ['00', 'All', 'TV', 'Once', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['01', 'All', 'TV', 'Once', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['02', 'All', 'TV', 'Once', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['03', 'All', 'TV', 'Daily', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['04', 'All', 'TV', 'Daily', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['05', 'All', 'TV', 'Daily', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['06', 'All', 'TV', 'Mon.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['07', 'All', 'TV', 'Mon.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['08', 'All', 'TV', 'Mon.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['09', 'All', 'TV', 'Tues.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['10', 'All', 'TV', 'Tues.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['11', 'All', 'TV', 'Tues.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['12', 'All', 'TV', 'Wed.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['13', 'All', 'TV', 'Wed.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['14', 'All', 'TV', 'Wed.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['15', 'All', 'TV', 'Thurs.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['16', 'All', 'TV', 'Thurs.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['17', 'All', 'TV', 'Thurs.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['18', 'All', 'TV', 'Fri.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['19', 'All', 'TV', 'Fri.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['20', 'All', 'TV', 'Fri.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['21', 'All', 'TV', 'Sat.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['22', 'All', 'TV', 'Sat.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['23', 'All', 'TV', 'Sat.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['24', 'All', 'TV', 'Sun.', 'Play', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['25', 'All', 'TV', 'Sun.', 'Play', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['26', 'All', 'TV', 'Sun.', 'Play', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['27', 'All', 'TV', 'Once', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['28', 'All', 'TV', 'Once', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['29', 'All', 'TV', 'Once', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['30', 'All', 'TV', 'Daily', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['31', 'All', 'TV', 'Daily', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['32', 'All', 'TV', 'Daily', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['33', 'All', 'TV', 'Mon.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['34', 'All', 'TV', 'Mon.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['35', 'All', 'TV', 'Mon.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['36', 'All', 'TV', 'Tues.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['37', 'All', 'TV', 'Tues.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['38', 'All', 'TV', 'Tues.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['39', 'All', 'TV', 'Wed.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['40', 'All', 'TV', 'Wed.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['41', 'All', 'TV', 'Wed.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['42', 'All', 'TV', 'Thurs.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['43', 'All', 'TV', 'Thurs.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['44', 'All', 'TV', 'Thurs.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['45', 'All', 'TV', 'Fri.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['46', 'All', 'TV', 'Fri.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['47', 'All', 'TV', 'Fri.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['48', 'All', 'TV', 'Sat.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['49', 'All', 'TV', 'Sat.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['50', 'All', 'TV', 'Sat.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['51', 'All', 'TV', 'Sun.', 'PVR', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['52', 'All', 'TV', 'Sun.', 'PVR', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['53', 'All', 'TV', 'Sun.', 'PVR', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['54', 'All', 'TV', 'Once', 'Power Off', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['55', 'All', 'TV', 'Once', 'Power Off', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['56', 'All', 'TV', 'Once', 'Power Off', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['57', 'All', 'TV', 'Daily', 'Power Off', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['58', 'All', 'TV', 'Daily', 'Power Off', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['59', 'All', 'TV', 'Daily', 'Power Off', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['60', 'All', 'TV', 'Mon.', 'Power Off', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['61', 'All', 'TV', 'Mon.', 'Power Off', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['62', 'All', 'TV', 'Mon.', 'Power Off', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['63', 'All', 'Radio', 'Once', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['64', 'All', 'TV', 'Once', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['65', 'All', 'TV', 'Daily', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['66', 'All', 'TV', 'Mon.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['67', 'All', 'TV', 'Tues.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['68', 'All', 'TV', 'Wed.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['69', 'All', 'TV', 'Thurs.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['70', 'All', 'TV', 'Fri.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['71', 'All', 'TV', 'Sat.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['72', 'All', 'TV', 'Sun.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['73', 'All', 'TV', 'Daily', 'Play', 'ChannelList', 'Manual_jump', 'other_interface_test_numb'],
    ['74', 'All', 'TV', 'Mon.', 'Play', 'Menu', 'Manual_jump', 'other_interface_test_numb'],
    ['75', 'All', 'TV', 'Fri.', 'Play', 'EPG', 'Manual_jump', 'other_interface_test_numb'],
    ['76', 'All', 'TV', 'Once', 'Play', 'ChannelEdit', 'Manual_jump', 'other_interface_test_numb']
]