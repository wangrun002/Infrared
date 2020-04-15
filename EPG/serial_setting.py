#!/usr/bin/python3
# -*- coding: utf-8 -*-

import serial
import serial.tools.list_ports
import platform
import logging

ser_cable_num = 7

def check_ports():
    global send_com,receive_com
    ports_info = []
    if platform.system() == "Windows":
        # ser_cable_num = 7
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
        # ser_cable_num = 7
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
    return send_com,receive_com