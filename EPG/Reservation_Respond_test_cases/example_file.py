#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import shutil
import platform

choice_case_numb = 0

parent_path = os.path.dirname(os.getcwd())
main_file_path = os.path.join(parent_path, "MainProgram","Reservation_and_triggered_event.py")
test_file_path = os.path.join(os.getcwd(), "Reservation_and_triggered_event.py")

shutil.copy(main_file_path, os.getcwd())
if platform.system() == "Windows":
    os.system("python %s %d" % (test_file_path, choice_case_numb))
elif platform.system() == "Linux":
    os.system("python3 %s %d" % (test_file_path, choice_case_numb))
os.unlink(test_file_path)

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
    ['54', 'All', 'TV', 'Once', 'Power 0ff', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['55', 'All', 'TV', 'Once', 'Power 0ff', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['56', 'All', 'TV', 'Once', 'Power 0ff', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['57', 'All', 'TV', 'Daily', 'Power 0ff', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['58', 'All', 'TV', 'Daily', 'Power 0ff', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['59', 'All', 'TV', 'Daily', 'Power 0ff', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
    ['60', 'All', 'TV', 'Mon.', 'Power 0ff', 'TVScreenDiffCH', 'Manual_jump', 'screen_test_numb'],
    ['61', 'All', 'TV', 'Mon.', 'Power 0ff', 'TVScreenDiffCH', 'Auto_jump', 'screen_test_numb'],
    ['62', 'All', 'TV', 'Mon.', 'Power 0ff', 'TVScreenDiffCH', 'Cancel_jump', 'screen_test_numb'],
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

