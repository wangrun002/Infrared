# !/usr/bin/python3
# -*- coding: UTF-8 -*-

from redminelib import Redmine
from datetime import datetime, timedelta
import csv
import pandas as pd
import re
import os
import platform

specify_project_name = 'BU1-SDK-2007-01-UNIFY'      # issue所在项目名称
pj_version_name = 'SDK_V2.4.0'                      # 目标版本名称
filter_issue_subject_kws = 'Gemini 6702H5'          # 主题中筛选方案名称的关键字
filter_issue_created_time = '2020-11-16'            # 用于筛选新建和遗留问题的日期，格式一定要为: XXXX-XX-XX


class RedmineHandle(object):

    def __init__(self):
        redmine_url = 'http://git.nationalchip.com/redmine'
        api_key = '53df3675279864992b561cb4c8af488a2d36c1a0'
        # api_key = '994d302b537ddaddb170058885da1afb4f4acf57'
        self.redmine = Redmine(redmine_url, key=api_key)

    def print_all_project_base_info(self):
        # 打印所有项目的基本信息（id,name,identifier）
        projects = self.redmine.project.all()
        for project in projects:
            print(project.id, '---->', project.name, '---->', project.identifier)  # id, name, 标识符

    def get_project_id_by_name(self, project_name):
        # 根据项目名称获取项目的id信息
        projects = self.redmine.project.all()
        for project in projects:
            # print(project.id, '---->', project.name)
            if project.name == project_name:
                return project.id

    def get_specify_pj_version_by_name(self, project_id, version_name):
        version_id = ''
        print('当前项目的所有版本信息如下：')
        print('{:#^50}'.format('开始'))
        for version in self.redmine.project.get(project_id).versions:
            print(version.id, '---->', version.name)
            if version.name == version_name:
                version_id = version.id
        print('{:#^50}'.format('结束'))
        return version_id

    def get_project_info_by_project_id(self, project_id):
        return self.redmine.project.get(project_id)

    def get_issues(self, project_id, project_version_id, status_id, tracker_id):
        return self.redmine.issue.filter(
            project_id=project_id, fixed_version_id=project_version_id, status_id=status_id, tracker_id=tracker_id)

    def get_all_statuses(self):
        for status in self.redmine.issue_status.all():
            print(status.id, status.name)

    def get_all_trackers(self):
        tracker_dict = {}
        for tracker in self.redmine.tracker.all():
            print(tracker.id, tracker.name)
            tracker_dict[f'{tracker.id}'] = tracker.name
        # print(tracker_dict)

    def get_issue_info(self, issue_id):
        return self.redmine.issue.get(issue_id, include=['journals'])

    def issue_data_to_csv(self, issues):
        csv_path = ''
        if platform.system() == 'Linux':
            csv_path = r'./%s/issues.csv' % output_dir
        elif platform.system() == 'Windows':
            csv_path = r'.\%s\issues.csv' % output_dir

        def get_custom_fields(issue, custom_fields_id):
            try:
                return issue.custom_fields.get(custom_fields_id).value
            except:
                return None

        def get_assigned_to_name(issue):
            try:
                return issue.assigned_to.name.split()[-1]
            except:
                return None

        def get_updated_on(issue):
            try:
                return issue.updated_on + timedelta(hours=8)
            except:
                return None

        def get_closed_on(issue):
            try:
                return issue.closed_on + timedelta(hours=8)
            except:
                return None

        with open(csv_path, 'w', newline='') as f:
            f_csv = csv.DictWriter(f, fieldnames=CSV_FIELDS.keys())
            f_csv.writeheader()

            for issue in issues:
                # print(issue.id, issue.subject)
                #
                # print(issue.id)                                     # bug编号
                # print(issue.tracker)                                # 追踪
                # print(issue.status.name)                            # bug状态名称
                # print(issue.priority.name)                          # 优先级
                # print(issue.author.name.split()[-1])                # 作者
                # print(assigned_to)                                  # 指派给　　(有空值问题)
                # print(issue.created_on + timedelta(hours=8))        # 创建时间
                # print(issue.updated_on + timedelta(hours=8))        # 最后更新时间
                # print(issue.closed_on + timedelta(hours=8))         # 关闭时间  (有空值问题)
                # # print(issue_info.due_date)                        # 计划完成时间    (有空值问题)
                # print(issue.category)                               # 类别（开发用的类别)   (有空值问题)
                # print(issue.custom_fields.get(20).value)            # 出现频率
                # print(issue.custom_fields.get(17).value)            # 严重性
                # print(issue.custom_fields.get(21).value)            # 处理状况      (有空值问题)
                # print(issue.custom_fields.get(79).value)            # 卫星模块分类
                # print(issue.custom_fields.get(83).value)            # 问题性质      (有空值问题)
                # print(issue.custom_fields.get(84).value)            # 问题归属      (有空值问题)
                # print(issue.custom_fields.get(105).value)           # 结论        (有空值问题)

                author = issue.author.name.split()[-1]
                assigned_to = get_assigned_to_name(issue)
                created_on = issue.created_on + timedelta(hours=8)
                updated_on = get_updated_on(issue)
                closed_on = get_closed_on(issue)
                frequency = get_custom_fields(issue, 20)
                seriousness = get_custom_fields(issue, 17)
                situation = get_custom_fields(issue, 21)
                modules = get_custom_fields(issue, 79)
                if author in Author_ptd:
                    if re.findall(f'{filter_issue_subject_kws}', issue.subject):
                        f_csv.writerow({
                            "#": issue.id,
                            "tracker": issue.tracker,
                            "status": issue.status.name,
                            "priority": issue.priority.name,
                            "author": author,
                            "assigned_to": assigned_to,
                            "created_on": created_on,
                            "updated_on": updated_on,
                            "closed_on": closed_on,
                            "frequency": frequency,
                            "seriousness": seriousness,
                            "situation": situation,
                            "modules": modules,
                            "subject": issue.subject
                        })


class Data(object):

    def __init__(self):
        encode_fmt = ''
        if platform.system() == 'Linux':
            encode_fmt = 'utf-8'
        elif platform.system() == 'Windows':
            encode_fmt = 'gb18030'
        self.df = pd.read_csv(f'./{output_dir}/issues.csv', encoding=encode_fmt)

    def all(self):
        return self.df

    def authors(self, data_frame, authors):
        return data_frame[data_frame['author'].isin(authors)]

    def status(self, data_frame, statuses):
        return data_frame[data_frame['status'].isin(statuses)]

    def situation(self, data_frame, situation):
        return data_frame[data_frame['situation'].isin(situation)]

    def new_created(self, data_frame, filter_time):
        return data_frame[pd.to_datetime(data_frame['created_on']) > pd.to_datetime(filter_time)]

    def legacy(self, data_frame, filter_time):
        return data_frame[pd.to_datetime(data_frame['created_on']) <= pd.to_datetime(filter_time)]

    def write_filter_data(self, f_path, w_data):
        with open(f_path, 'a+', encoding='utf-8') as fo:
            fo.write(w_data)

    def handle_filter_issue_by_module(self, issues_data, result_name):
        output_file_path = f'./{output_dir}/{result_name}.txt'
        title = f'过滤到的{result_name}的issues总数为：（{len(issues_data)}）个:\n\n'
        if len(issues_data) > 0:
            data_dict = {}
            for index_n in issues_data.index:
                if issues_data.loc[index_n, 'modules'] not in data_dict:
                    data_dict[issues_data.loc[index_n, 'modules']] = []
                data_dict[issues_data.loc[index_n, 'modules']].append(issues_data.loc[index_n, '#'])
            # for i in range(len(issues_data)):
            #     if list(issues_data['modules'])[i] not in data_dict:
            #         data_dict[list(issues_data['modules'])[i]] = []
            #     data_dict[list(issues_data['modules'])[i]].append(list(issues_data['#'])[i])

            # 写title
            self.write_filter_data(output_file_path, title)
            for key in list(data_dict.keys()):
                # print(f'{key}:')
                self.write_filter_data(output_file_path, f'{key}:共({len(data_dict[key])})个\n')
                for value in data_dict[key]:
                    # print('* {{issue(%d)}}' % value)
                    self.write_filter_data(output_file_path, '* {{issue(%d)}}\n' % value)
                # print('')
                self.write_filter_data(output_file_path, '\n')
        else:
            # 没有数据，直接写title
            self.write_filter_data(output_file_path, title)

    def handle_filter_issue_by_seriousness(self, issues_data, result_name):
        output_file_path = f'./{output_dir}/{result_name}.txt'
        title = f'过滤到的{result_name}的issues总数为：（{len(issues_data)}）个:\n\n'
        if len(issues_data) > 0:
            data_dict = {}
            for index_n in issues_data.index:
                if issues_data.loc[index_n, 'seriousness'] not in data_dict:
                    data_dict[issues_data.loc[index_n, 'seriousness']] = []
                data_dict[issues_data.loc[index_n, 'seriousness']].append(issues_data.loc[index_n, '#'])
            # for i in range(len(issues_data)):
            #     if list(issues_data['seriousness'])[i] not in data_dict:
            #         data_dict[list(issues_data['seriousness'])[i]] = []
            #     data_dict[list(issues_data['seriousness'])[i]].append(list(issues_data['#'])[i])

            # 写title
            self.write_filter_data(output_file_path, title)
            for key in list(data_dict.keys()):
                # print(f'{key}:')
                self.write_filter_data(output_file_path, f'{key}:共({len(data_dict[key])})个\n')
                for value in data_dict[key]:
                    # print('* {{issue(%d)}}' % value)
                    self.write_filter_data(output_file_path, '* {{issue(%d)}}\n' % value)
                # print('')
                self.write_filter_data(output_file_path, '\n')
        else:
            # 没有数据，直接写title
            self.write_filter_data(output_file_path, title)

    def handle_filter_issue_by_module_and_seriousness(self, issues_data, result_name):
        output_file_path = f'./{output_dir}/{result_name}.txt'
        title = f'过滤到的{result_name}的issues总数为：（{len(issues_data)}）个:\n\n'
        serious_name_list = ['死机', '严重', '一般', '建议']
        module_issues_dict = {}     # 各个模块的问题个数
        crash_and_serious_issues_dict = {}  # 各个模块的死机＋严重问题个数
        if len(issues_data) > 0:
            data_dict = {}
            for index_n in issues_data.index:
                if issues_data.loc[index_n, 'modules'] not in data_dict:
                    data_dict[issues_data.loc[index_n, 'modules']] = {}
                if issues_data.loc[index_n, 'seriousness'] not in data_dict[issues_data.loc[index_n, 'modules']]:
                    data_dict[issues_data.loc[index_n, 'modules']][issues_data.loc[index_n, 'seriousness']] = []
                data_dict[issues_data.loc[index_n, 'modules']][issues_data.loc[index_n, 'seriousness']].append(
                    issues_data.loc[index_n, '#']
                )

            # for i in range(len(issues_data)):
            #     if issues_data['modules'][i] not in data_dict:
            #         data_dict[issues_data['modules'][i]] = {}
            #     if issues_data['seriousness'][i] not in data_dict[issues_data['modules'][i]]:
            #         data_dict[issues_data['modules'][i]][issues_data['seriousness'][i]] = []
            #     data_dict[issues_data['modules'][i]][issues_data['seriousness'][i]].append(issues_data['#'][i])

            # 写title
            self.write_filter_data(output_file_path, title)
            # 写数据
            for mod in list(data_dict.keys()):
                # print(f'{key}:')
                module_issues_dict[mod] = sum(list(map(lambda x: len(x), list(data_dict[mod].values()))))
                self.write_filter_data(output_file_path, '%s:(共%d个)\n' % (
                    mod, sum(list(map(lambda x: len(x), list(data_dict[mod].values()))))))
                for serious_name in serious_name_list:
                    if serious_name in list(data_dict[mod].keys()):
                        self.write_filter_data(output_file_path, '%s:(%d)个    %s\n' % (
                            serious_name, len(data_dict[mod][serious_name]), data_dict[mod][serious_name]))
                    else:
                        self.write_filter_data(
                            output_file_path, '%s:(%d)个\n' % (serious_name, 0))
                # # print('')
                self.write_filter_data(output_file_path, '\n')
                crash_and_serious_numb_list = []
                for serious_name in list(data_dict[mod].keys()):
                    if serious_name in serious_name_list[:2]:
                        crash_and_serious_numb_list.append(len(data_dict[mod][serious_name]))
                crash_and_serious_issues_dict[mod] = sum(crash_and_serious_numb_list)

            # 生成问题从多到少分布报告
            sort_by_mod_numb_name = '问题从多到少分布'
            result_path = f'./{output_dir}/{sort_by_mod_numb_name}.txt'
            # 按模块的问题个数从多到少分布
            sorted_module_issues_list = sorted(module_issues_dict.items(), key=lambda x: x[1], reverse=True)
            title_msg = f'按模块的问题个数从多到少分布:\n\n'
            self.write_filter_data(result_path, title_msg)
            for mod_numb in sorted_module_issues_list:
                self.write_filter_data(result_path, '%s:%d\n' % (mod_numb[0], mod_numb[1]))
            self.write_filter_data(result_path, '\n{:#^50}\n\n'.format('分割线'))

            # 按模块的死机＋严重问题个数从多到少分布
            sorted_crash_and_serious_issues_list = sorted(
                crash_and_serious_issues_dict.items(), key=lambda x: x[1], reverse=True)
            title_msg = f'按模块的死机＋严重问题个数从多到少分布:\n\n'
            self.write_filter_data(result_path, title_msg)
            for mod_numb in sorted_crash_and_serious_issues_list:
                self.write_filter_data(result_path, '%s:%d\n' % (mod_numb[0], mod_numb[1]))

        else:
            # 写title
            self.write_filter_data(output_file_path, title)


if __name__ == '__main__':

    Issue_status_dict = {
        '1': '新建',
        '2': '已确认',
        '3': '已解决',
        '4': '反馈',
        '5': '已关闭',
        '6': '已拒绝',
    }

    custom_field_dict = {
        '20': '出现频率',
        '17': '严重性',
        '21': '处理状况',
        '79': '卫星模块分类',
        '83': '问题性质',
        '84': '问题归属',
        '105': '结论',
    }

    Tracker_dict = {
        '1': '错误',
        '2': '功能',
        '3': '支持',
        '5': '结点',
        '6': '视频',
        '8': '需求',
        '10': '学习',
        '11': '测试',
        '12': '资料发布',
        '13': '机顶盒',
        '14': '智能音箱',
        '15': '机器人',
        '16': 'UAC',
        '17': '结点测试1',
        '18': '功能测试1',
        '19': '知识',
        '20': '资源',
        '21': '目标',
        '22': '关键结果',
        '23': '评价'
    }

    CSV_FIELDS = {
        "#":            "ID",
        "tracker":      "跟踪",
        "status":       "状态",
        "priority":     "优先级",
        "author":       "作者",
        "assigned_to":  "指派给",
        "created_on":   "创建于",
        "updated_on":   "更新于",
        "closed_on":    "结束日期",
        "frequency":    "出现频率",
        "seriousness":  "严重性",
        "situation":    "处理状况",
        "modules":      "卫星模块分类",
        "subject":      "主题"
    }

    Author_ptd = ['陈材', '吴胜燕', '王丽君', '王润', '颜潮锋', '罗萍', '宋戈', '倪斌']
    # 创建输出目录
    now = datetime.now()
    output_dir = f'__{specify_project_name}__{filter_issue_subject_kws}__' \
                 f'{pj_version_name}__{now.strftime("%Y%m%d%H%M%S")}'
    if os.path.exists(output_dir):
        pass
    else:
        os.mkdir(output_dir)

    RH = RedmineHandle()

    # 打印redmine上所有的项目的基本信息
    RH.print_all_project_base_info()

    specify_project_id = RH.get_project_id_by_name(specify_project_name)
    print(f'指定的项目id为：{specify_project_id}')

    pj = RH.get_project_info_by_project_id(specify_project_id)

    pj_version_name = 'SDK_V2.4.0'
    pj_version_id = RH.get_specify_pj_version_by_name(specify_project_id, pj_version_name)
    # print(pj_version_id)

    # 筛选指定版本下的issue
    issues = RH.get_issues(specify_project_id, pj_version_id, 'open', 1)
    print(f'根据指定版本过滤到的打开的错误个数为：{len(issues)}')

    # 将issue信息逐个写入csv
    RH.issue_data_to_csv(issues)

    data = Data()
    # 打印项目中issue的所有状态
    # RH.get_all_statuses()
    # 打印项目中issue的所有跟踪
    # RH.get_all_trackers()

    # 获取某issue的信息
    # issue_info = RH.get_issue_info(271210)
    # print(issue_info)

    # 获取csv中所有的issues信息
    all_issues = data.all()
    # print(all_issues)
    # 从all_issues中筛选ptd组成员的issues
    issues_belong_to_ptd = data.authors(all_issues, Author_ptd)
    print(f'筛选属于PTD的打开的问题数为：{len(issues_belong_to_ptd)}')

    while True:
        choice_filter_issue_scene = input("\033[1;31m请选择过滤问题场景，获取回归测试数据请输入1，获取方案测试数据请输入2：\033[0m")
        if choice_filter_issue_scene.strip() == '1':
            # 从issues_belong_to_ptd中筛选(已解决已合并)的issues
            resolved_issues = data.status(issues_belong_to_ptd, ['已解决'])
            resolved_merged_issues = data.situation(resolved_issues, ['已合并'])
            resolved_merged_result_name = '(已解决已合并)问题按模块分布'
            data.handle_filter_issue_by_module(resolved_merged_issues, resolved_merged_result_name)
            # print(len(resolved_merged_issues['#']))
            # for i in resolved_merged_issues['#']:
            #     print('* {{issue(%d)}}' % i)

            # 从issues_belong_to_ptd中筛选(反馈)的issues
            feedback_issues = data.status(issues_belong_to_ptd, ['反馈'])
            feedback_result_name = '(反馈)问题按模块分布'
            data.handle_filter_issue_by_module(feedback_issues, feedback_result_name)

            # print(feedback_issues)
            # for j in feedback_issues['#']:
            #     print('* {{issue(%d)}}' % j)

            break
        elif choice_filter_issue_scene.strip() == '2':

            # 本轮新增问题按严重程度分布
            new_created_issues = data.new_created(issues_belong_to_ptd, filter_issue_created_time)
            # print(len(new_created_issues))
            # for j in new_created_issues['#']:
            #     print('* {{issue(%d)}}' % j)
            new_created_result_name = '(新增)问题按严重程度分布'
            data.handle_filter_issue_by_seriousness(new_created_issues, new_created_result_name)

            # 遗留问题按严重程度分布
            legacy_issues = data.legacy(issues_belong_to_ptd, filter_issue_created_time)
            # print(len(legacy_issues))
            # for j in legacy_issues['#']:
            #     print('* {{issue(%d)}}' % j)
            legacy_result_name = '(遗留)问题按严重程度分布'
            data.handle_filter_issue_by_seriousness(legacy_issues, legacy_result_name)

            # 按模块和严重程度输出
            module_and_seriousness_name = '所有问题按(模块和严重程度)分布'
            data.handle_filter_issue_by_module_and_seriousness(issues_belong_to_ptd, module_and_seriousness_name)

            # 所有打开问题按严重程度分布
            all_issues_result_name = '所有问题按(严重程度)分布'
            data.handle_filter_issue_by_seriousness(issues_belong_to_ptd, all_issues_result_name)
            break
        else:
            print('\033[1;31m场景选择输入错误，请输入正确的场景代号\033[0m')
