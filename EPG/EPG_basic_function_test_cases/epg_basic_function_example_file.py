#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import shutil
import platform

choice_case_numb = 43

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path, "MainProgram", "EPG_basic_function_test.py")
test_file_path = os.path.join(os.getcwd(), "EPG_basic_function_test.py")

shutil.copy(main_file_path, os.getcwd())
if platform.system() == "Windows":
    os.system("python %s %d" % (test_file_path, choice_case_numb))
elif platform.system() == "Linux":
    os.system("python3 %s %d" % (test_file_path, choice_case_numb))
os.unlink(test_file_path)

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
