#!/usr/bin/python3
# -*- coding: utf-8 -*-

screen_test_numb = 20
other_interface_test_numb = 50
total_list = []
n = 0
res_type = ["Play", "PVR", "Power Off", "Power On"]
res_mode = ["Once", "Daily", "Weekly"]
scenes_list = [
    "Same(time+type+mode)",
    "Same(time+mode)+Diff(type)",
    "Same(time+type)+Diff(mode)",
    "Same(time)+Diff(type+mode)",
    ]
# "Same(time+type+mode)"
for i in res_type:
    for j in res_mode:
        single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                       "Same[time+type+mode]", "", "", "screen_test_numb"]
        single_list[0] = "{0:02d}".format(n)
        single_list[3] = j
        single_list[4] = i
        single_list[8] = j
        single_list[9] = i
        total_list.append(single_list)
        n += 1

# "Same(time+mode)+Diff(type)"
for i in res_mode:
    for j in res_type:
        for k in res_type:
            if k != j:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "Same[time+mode]+Diff[type]", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = i
                single_list[4] = j
                single_list[8] = i
                single_list[9] = k
                total_list.append(single_list)
                n += 1

# "Same(time+type)+Diff(mode)"
for i in res_type:
    for j in res_mode:
        for k in res_mode:
            if k != j:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "Same[time+type]+Diff[mode]", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = j
                single_list[4] = i
                single_list[8] = k
                single_list[9] = i
                total_list.append(single_list)
                n += 1

# "Same(time)+Diff(type+mode)"
for i in res_type:
    for j in res_type:
        for k in res_mode:
            for h in res_mode:
                if j != i and h != k:
                    single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                   "Same[time]+Diff[type+mode]", "", "", "screen_test_numb"]
                    single_list[0] = "{0:02d}".format(n)
                    single_list[3] = k
                    single_list[4] = i
                    single_list[8] = h
                    single_list[9] = j
                    total_list.append(single_list)
                    n += 1

# "Same[type+mode+dur]+Diff[time]"
for i in res_type:
    for j in res_mode:
        if i == "PVR":
            single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                           "Same[type+dur+mode]+Diff[time]", "", "", "screen_test_numb"]
            single_list[0] = "{0:02d}".format(n)
            single_list[3] = j
            single_list[4] = i
            single_list[8] = j
            single_list[9] = i
            total_list.append(single_list)
            n += 1

# "Same[type+dur]+Diff[time+mode]"
for i in res_type:
    for j in res_mode:
        for k in res_mode:
            if i == "PVR" and k != j:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "Same[type+dur]+Diff[time+mode]", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = j
                single_list[4] = i
                single_list[8] = k
                single_list[9] = i
                total_list.append(single_list)
                n += 1

# "Same[mode]+Diff[time+type+dur]"
for j in res_type:
    for k in res_type:
        for i in res_mode:
            if j == "PVR" and k != j:
                single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                               "Same[mode]+Diff[time+type+dur]", "", "", "screen_test_numb"]
                single_list[0] = "{0:02d}".format(n)
                single_list[3] = i
                single_list[4] = j
                single_list[8] = i
                single_list[9] = k
                total_list.append(single_list)
                n += 1

# "Diff[time+type+dur+mode]"
for k in res_type:
    for h in res_type:
        for i in res_mode:
            for j in res_mode:
                if j != i and k == "PVR" and h != k:
                    single_list = ["", "All", "TV", "", "", "TVScreenDiffCH", "Manual_jump",
                                   "Diff[time+type+dur+mode]", "", "", "screen_test_numb"]
                    single_list[0] = "{0:02d}".format(n)
                    single_list[3] = i
                    single_list[4] = k
                    single_list[8] = j
                    single_list[9] = h
                    total_list.append(single_list)
                    n += 1

# print(len(total_list))
for m in range(len(total_list)):
    print(f"{total_list[m]},")
