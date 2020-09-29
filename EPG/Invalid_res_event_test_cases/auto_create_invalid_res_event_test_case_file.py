#!/usr/bin/python3
# -*- coding: utf-8 -*-

invalid_res_case = [
    ['00', 'GX', 'TV', 'Play', 'Once', 'Expired', 'EPG', 'epg_test_numb'],
    ['01', 'GX', 'TV', 'PVR', 'Once', 'Expired', 'EPG', 'epg_test_numb'],
    ['02', 'GX', 'Radio', 'Play', 'Once', 'Expired', 'EPG', 'epg_test_numb'],
    ['03', 'GX', 'Radio', 'PVR', 'Once', 'Expired', 'EPG', 'epg_test_numb'],
    ['04', 'GX', 'TV', 'Play', 'Once', 'NowPlaying', 'EPG', 'epg_test_numb'],
    ['05', 'GX', 'TV', 'PVR', 'Once', 'NowPlaying', 'EPG', 'epg_test_numb'],
    ['06', 'GX', 'Radio', 'Play', 'Once', 'NowPlaying', 'EPG', 'epg_test_numb'],
    ['07', 'GX', 'Radio', 'PVR', 'Once', 'NowPlaying', 'EPG', 'epg_test_numb'],
    ['08', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_lower_limit', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['09', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_upper_limit', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['10', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_lower_limit', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['11', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_upper_limit', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['12', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_before_save_time_range', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['13', 'GX', 'TV', 'Play', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_after_save_time_range', 'ZeroTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['14', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_lower_limit', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['15', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_upper_limit', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['16', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_lower_limit', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['17', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_upper_limit', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['18', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_before_save_time_range', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['19', 'GX', 'TV', 'PVR', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_after_save_time_range', 'ZeroTimezone',
     'Summertime', 'epg_test_numb'],
    ['20', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_lower_limit',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['21', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_upper_limit',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['22', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_lower_limit',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['23', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_upper_limit',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['24', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_before_save_time_range',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['25', 'GX', 'TV', 'Power Off', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_after_save_time_range',
     'OtherTimezone', 'NoSummertime', 'epg_test_numb'],
    ['26', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_lower_limit',
     'OtherTimezone', 'Summertime', 'epg_test_numb'],
    ['27', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_before_upper_limit',
     'OtherTimezone', 'Summertime', 'epg_test_numb'],
    ['28', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_lower_limit', 'OtherTimezone',
     'Summertime', 'epg_test_numb'],
    ['29', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Boundary_after_upper_limit', 'OtherTimezone',
     'Summertime', 'epg_test_numb'],
    ['30', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_before_save_time_range',
     'OtherTimezone', 'Summertime', 'epg_test_numb'],
    ['31', 'GX', 'TV', 'Power On', 'Once', 'OutOfSaveTimeRange', 'Timer', 'Random_after_save_time_range',
     'OtherTimezone', 'Summertime', 'epg_test_numb'],
    ['32', 'GX', 'TV', 'Play', 'Once', 'Expired', 'Timer', 'Boundary_lower_limit', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['33', 'GX', 'TV', 'Play', 'Once', 'Expired', 'Timer', 'Boundary_upper_limit', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['34', 'GX', 'TV', 'Play', 'Once', 'Expired', 'Timer', 'Random_expired_time_range', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['35', 'GX', 'TV', 'PVR', 'Once', 'Expired', 'Timer', 'Boundary_lower_limit', 'ZeroTimezone', 'Summertime',
     'epg_test_numb'],
    ['36', 'GX', 'TV', 'PVR', 'Once', 'Expired', 'Timer', 'Boundary_upper_limit', 'ZeroTimezone', 'Summertime',
     'epg_test_numb'],
    ['37', 'GX', 'TV', 'PVR', 'Once', 'Expired', 'Timer', 'Random_expired_time_range', 'ZeroTimezone', 'Summertime',
     'epg_test_numb'],
    ['38', 'GX', 'TV', 'Power Off', 'Once', 'Expired', 'Timer', 'Boundary_lower_limit', 'OtherTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['39', 'GX', 'TV', 'Power Off', 'Once', 'Expired', 'Timer', 'Boundary_upper_limit', 'OtherTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['40', 'GX', 'TV', 'Power Off', 'Once', 'Expired', 'Timer', 'Random_expired_time_range', 'OtherTimezone',
     'NoSummertime', 'epg_test_numb'],
    ['41', 'GX', 'TV', 'Power On', 'Once', 'Expired', 'Timer', 'Boundary_lower_limit', 'OtherTimezone', 'Summertime',
     'epg_test_numb'],
    ['42', 'GX', 'TV', 'Power On', 'Once', 'Expired', 'Timer', 'Boundary_upper_limit', 'OtherTimezone', 'Summertime',
     'epg_test_numb'],
    ['43', 'GX', 'TV', 'Power On', 'Once', 'Expired', 'Timer', 'Random_expired_time_range', 'OtherTimezone',
     'Summertime', 'epg_test_numb'],
    ['44', 'GX', 'TV', 'Play', 'Once', 'NowPlaying', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime', 'epg_test_numb'],
    ['45', 'GX', 'TV', 'PVR', 'Once', 'NowPlaying', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime', 'epg_test_numb'],
    ['46', 'GX', 'TV', 'Power Off', 'Once', 'NowPlaying', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['47', 'GX', 'TV', 'Power On', 'Once', 'NowPlaying', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['48', 'GX', 'TV', 'PVR', 'Once', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['49', 'GX', 'TV', 'PVR', 'Daily', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['50', 'GX', 'TV', 'PVR', 'Mon.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['51', 'GX', 'TV', 'PVR', 'Tues.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['52', 'GX', 'TV', 'PVR', 'Wed.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['53', 'GX', 'TV', 'PVR', 'Thurs.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['54', 'GX', 'TV', 'PVR', 'Fri.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['55', 'GX', 'TV', 'PVR', 'Sat.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb'],
    ['56', 'GX', 'TV', 'PVR', 'Sun.', 'InvalidDuration', 'Timer', 'None', 'ZeroTimezone', 'NoSummertime',
     'epg_test_numb']
]

read_data = []

with open("create_invalid_res_event_example_file.py", "r") as f:
    for line in f.readlines():
        print(line, end='')
        read_data.append(line)


def build_x_numb_python_file():
    for i in range(len(invalid_res_case)):
        python_file_path = ''
        if invalid_res_case[i][6] == "EPG":
            python_file_path = "Invalid_{}_{}_{}_{}_{}_{}_{}.py".format(
                invalid_res_case[i][0], invalid_res_case[i][1],
                invalid_res_case[i][2], invalid_res_case[i][3],
                invalid_res_case[i][4], invalid_res_case[i][5],
                invalid_res_case[i][6])
        elif invalid_res_case[i][6] == "Timer":
            if invalid_res_case[i][3] == "Power Off" or invalid_res_case[i][3] == "Power On":
                python_file_path = "Invalid_{}_{}_{}_{}_{}_{}_{}_{}_{}_{}.py".format(
                    invalid_res_case[i][0], invalid_res_case[i][1],
                    invalid_res_case[i][2], invalid_res_case[i][3].replace(" ", ''),
                    invalid_res_case[i][4], invalid_res_case[i][5],
                    invalid_res_case[i][6], invalid_res_case[i][7],
                    invalid_res_case[i][8], invalid_res_case[i][9])
            else:
                python_file_path = "Invalid_{}_{}_{}_{}_{}_{}_{}_{}_{}_{}.py".format(
                    invalid_res_case[i][0], invalid_res_case[i][1],
                    invalid_res_case[i][2], invalid_res_case[i][3],
                    invalid_res_case[i][4], invalid_res_case[i][5],
                    invalid_res_case[i][6], invalid_res_case[i][7],
                    invalid_res_case[i][8], invalid_res_case[i][9])
        with open(python_file_path, "a+") as fo:
            for j in range(len(read_data)):
                if "choice_case_numb =" in read_data[j]:
                    fo.write("choice_case_numb = {}\n".format(int(invalid_res_case[i][0])))
                elif "case_info" in read_data[j]:
                    fo.write(f"case_info = {invalid_res_case[i]}\n")
                else:
                    fo.write(read_data[j])


build_x_numb_python_file()
