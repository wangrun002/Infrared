#!/usr/bin/python3
# -*- coding: utf-8 -*-

screen_test_numb = 20
other_interface_test_numb = 50
total_list = []
n = 0
res_type = ["Play", "PVR", "Power 0ff"]
res_mode = ["Once", "Daily", "Mon.", "Tues.", "Wed.", "Thurs.", "Fri.", "Sat.", "Sun."]
jump_mode = ["Manual_jump", "Auto_jump", "Cancel_jump"]

for i in res_type:
    for j in res_mode:
        for k in jump_mode:
            if i == "Power 0ff":
                if j == "Once" or j == "Daily" or j == "Mon.":
                    single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "", "screen_test_numb"]
                    single_list[0] = "{0:02d}".format(n)
                    single_list[3] = j
                    single_list[4] = i
                    single_list[6] = k
                    total_list.append(single_list)
            else:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = j
                single_list[4] = i
                single_list[6] = k
                total_list.append(single_list)
            n += 1


total_list.append(['63', 'All', 'Radio', 'Once', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])

total_list.append(['64', 'All', 'TV', 'Once', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['65', 'All', 'TV', 'Daily', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['66', 'All', 'TV', 'Mon.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['67', 'All', 'TV', 'Tues.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['68', 'All', 'TV', 'Wed.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['69', 'All', 'TV', 'Thurs.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['70', 'All', 'TV', 'Fri.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['71', 'All', 'TV', 'Sat.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])
total_list.append(['72', 'All', 'TV', 'Sun.', 'PVR', 'RadioScreenDiffCH', 'Manual_jump', 'screen_test_numb'])

total_list.append(['73', 'All', 'TV', 'Daily', 'Play', 'ChannelList', 'Manual_jump', 'other_interface_test_numb'])

total_list.append(['74', 'All', 'TV', 'Mon.', 'Play', 'Menu', 'Manual_jump', 'other_interface_test_numb'])

total_list.append(['75', 'All', 'TV', 'Fri.', 'Play', 'EPG', 'Manual_jump', 'other_interface_test_numb'])

total_list.append(['76', 'All', 'TV', 'Once', 'Play', 'ChannelEdit', 'Manual_jump', 'other_interface_test_numb'])





print(len(total_list))
for m in range(len(total_list)):
    print(f"{total_list[m]},")