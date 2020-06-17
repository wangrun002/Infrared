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
        # receive_port_desc = serial_ser[str(ser_cable_num)]
        receive_port_desc = "FT232R USB UART"
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
        # receive_port_desc = serial_ser[str(ser_cable_num)]
        receive_port_desc = "FT232R USB UART"
    ports = list(serial.tools.list_ports.comports())
    for i in range(len(ports)):
        logging.info("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device, ports[i].description, ports[i].hwid))
        # print("可用端口:名称:{} + 描述:{} + 硬件id:{}".format(ports[i].device, ports[i].description, ports[i].hwid))
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

epg_basic_case = [
    ["00", "GX", "TV", "EPG+Exit", "EnterExitEPGInterface"],
    ["01", "GX", "Radio", "EPG+Exit", "EnterExitEPGInterface"],
    ["02", "GX", "TV", "UP", "EPGSwitchChannel"],
    ["03", "GX", "TV", "DOWN", "EPGSwitchChannel"],
    ["04", "GX", "TV", "UP+Random", "EPGSwitchChannel"],
    ["05", "GX", "TV", "DOWN+Random", "EPGSwitchChannel"],
    ["06", "GX", "TV", "UP+DOWN+Random", "EPGSwitchChannel"],
    ["07", "GX", "Radio", "UP", "EPGSwitchChannel"],
    ["08", "GX", "Radio", "DOWN", "EPGSwitchChannel"],
    ["09", "GX", "Radio", "UP+Random", "EPGSwitchChannel"],
    ["10", "GX", "Radio", "DOWN+Random", "EPGSwitchChannel"],
    ["11", "GX", "Radio", "UP+DOWN+Random", "EPGSwitchChannel"],
    ["12", "GX", "TV", "LEFT", "SwitchEPGEvent"],
    ["13", "GX", "TV", "RIGHT", "SwitchEPGEvent"],
    ["14", "GX", "TV", "LEFT+Random", "SwitchEPGEvent"],
    ["15", "GX", "TV", "RIGHT+Random", "SwitchEPGEvent"],
    ["16", "GX", "TV", "RIGHT+LEFT+Random", "SwitchEPGEvent"],
    ["17", "GX", "TV", "Day+", "SwitchEPGEvent"],
    ["18", "GX", "TV", "Day-", "SwitchEPGEvent"],
    ["19", "GX", "TV", "Day++RIGHT+Random", "SwitchEPGEvent"],
    ["20", "GX", "TV", "Day++LEFT+Random", "SwitchEPGEvent"],
    ["21", "GX", "TV", "Day++LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["22", "GX", "TV", "Day-+RIGHT+Random", "SwitchEPGEvent"],
    ["23", "GX", "TV", "Day-+LEFT+Random", "SwitchEPGEvent"],
    ["24", "GX", "TV", "Day-+LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["25", "GX", "TV", "Day+orDay-orLEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["26", "GX", "Radio", "LEFT", "SwitchEPGEvent"],
    ["27", "GX", "Radio", "RIGHT", "SwitchEPGEvent"],
    ["28", "GX", "Radio", "LEFT+Random", "SwitchEPGEvent"],
    ["29", "GX", "Radio", "RIGHT+Random", "SwitchEPGEvent"],
    ["30", "GX", "Radio", "RIGHT+LEFT+Random", "SwitchEPGEvent"],
    ["31", "GX", "Radio", "Day+", "SwitchEPGEvent"],
    ["32", "GX", "Radio", "Day-", "SwitchEPGEvent"],
    ["33", "GX", "Radio", "Day++RIGHT+Random", "SwitchEPGEvent"],
    ["34", "GX", "Radio", "Day++LEFT+Random", "SwitchEPGEvent"],
    ["35", "GX", "Radio", "Day++LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["36", "GX", "Radio", "Day-+RIGHT+Random", "SwitchEPGEvent"],
    ["37", "GX", "Radio", "Day-+LEFT+Random", "SwitchEPGEvent"],
    ["38", "GX", "Radio", "Day-+LEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["39", "GX", "Radio", "Day+orDay-orLEFTorRIGHT+Random", "SwitchEPGEvent"],
    ["40", "GX", "TV", "LEFT+OK", "SwitchAndDetailEPGEvent"],
    ["41", "GX", "TV", "RIGHT+OK", "SwitchAndDetailEPGEvent"],
    ["42", "GX", "TV", "LEFT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["43", "GX", "TV", "RIGHT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["44", "GX", "Radio", "LEFT+OK", "SwitchAndDetailEPGEvent"],
    ["45", "GX", "Radio", "RIGHT+OK", "SwitchAndDetailEPGEvent"],
    ["46", "GX", "Radio", "LEFT+Random+OK", "SwitchAndDetailEPGEvent"],
    ["47", "GX", "Radio", "RIGHT+Random+OK", "SwitchAndDetailEPGEvent"]
]
