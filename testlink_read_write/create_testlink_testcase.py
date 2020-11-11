# !/usr/bin/python3
# -*- coding: UTF-8 -*-

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.styles.colors import RED, BLUE
from openpyxl.utils import get_column_letter, column_index_from_string
import testlink
import os
import logging
import re


def logging_info_setting():
    # 配置logging输出格式
    log_format = "%(asctime)s %(name)s %(levelname)s %(message)s"  # 配置输出日志的格式
    date_format = "%Y-%m-%d %H:%M:%S %a"  # 配置输出时间的格式
    logging.basicConfig(level=logging.DEBUG, format=log_format, datefmt=date_format)


# url = 'http://git.nationalchip.com/testlink/lib/api/xmlrpc.php'
# key = 'f5df4f1bd2bdd22403ec6b8b118d022c'
# tlc = testlink.TestlinkAPIClient(url, key)


class TestLinkHandle(object):

    def __init__(self):
        url = 'http://git.nationalchip.com/testlink/lib/api/xmlrpc.php'
        key = 'f5df4f1bd2bdd22403ec6b8b118d022c'
        self.tlc = testlink.TestlinkAPIClient(url, key)

    def get_all_projects(self):
        # 获取所有的项目列表
        r = []
        for i in self.tlc.getProjects():
            r.append({'id': i['id'], 'name': i['name']})
        return r

    def get_project_id_by_name(self, name):
        # 通过项目名称获取项目Ｉd
        for i in self.tlc.getProjects():
            if i['name'] == name:
                return i['id']
        return ''

    def get_suites(self, project_id):
        # 通过项目id获取项目测试用例集
        r = []
        for i in self.tlc.getFirstLevelTestSuitesForTestProject(project_id):
            r.append({'id': i['id'], 'name': i['name']})
            print({'id': i['id'], 'name': i['name'], 'node_table': i['node_table']})
            # logging.info(i)
        return r

    def delete_project(self, prefix):
        # 删除项目
        try:
            return self.tlc.deleteTestProject(prefix=prefix)[0]['status']
        except:
            return False

    def create_project(self, name, prefix):
        # 创建项目
        if self.get_project_id_by_name(name) == '':
            return self.tlc.createTestProject(testprojectname=name, testcaseprefix=prefix, active=True)[0]['id']
        else:
            self.delete_project(prefix)
            return self.tlc.createTestProject(testprojectname=name, testcaseprefix=prefix, active=True)[0]['id']

    def create_suite(self, project_id, suite_name, parent_id=None):
        # 创建用例集（每一个模块都是一个用例集）
        try:
            if parent_id is None:
                first_suite = self.tlc.createTestSuite(testprojectid=project_id, testsuitename=suite_name,
                                                       details=suite_name)
            else:
                first_suite = self.tlc.createTestSuite(testprojectid=project_id, testsuitename=suite_name,
                                                       details=suite_name, parentid=parent_id)
            return first_suite[0]['id']
        except:
            return False

    def creat_cases_steps(self, steps, test_case_name, test_suite_id, test_project_id,
                          author_login, summary, preconditions):
        # 创建用例和创建测试步骤
        try:
            if len(steps) < 0:
                self.tlc.initStep('', '', 2)
            for step in steps:
                self.tlc.appendStep(step['actions'], step['expected_results'], step['execution_type'])
            return self.tlc.createTestCase(
                testcasename=test_case_name, testsuiteid=test_suite_id,
                testprojectid=test_project_id, authorlogin=author_login, summary=summary, preconditions=preconditions)
        except:
            return False

    def check_module_suite_exist(self, project_id, cur_module_name):
        first_suite_info_dict = {}
        for i in self.tlc.getFirstLevelTestSuitesForTestProject(project_id):
            first_suite_info_dict[i['name']] = i['id']
        if cur_module_name in list(first_suite_info_dict.keys()):
            logging.debug('项目中存在当前模块')
            return first_suite_info_dict[cur_module_name]
        else:
            logging.debug('项目中不存在当前模块')
            return self.create_suite(project_id, cur_module_name)

    def get_speify_suite_sub_suite_info(self, suite_id):
        specify_suite_info_dict = {}
        suites = self.tlc.getTestSuitesForTestSuite(suite_id)
        print(f'#######{suites}')
        if len(suites) > 0:
            if 'id' in suites.keys():
                specify_suite_info_dict[suites['name']] = suites['id']
                return specify_suite_info_dict
            elif 'id' not in suites.keys():
                for id_key in suites:
                    specify_suite_info_dict[suites[id_key]['name']] = suites[id_key]['id']
                return specify_suite_info_dict
        elif len(suites) == 0:
            return specify_suite_info_dict

    def check_case_suite_exist(self, project_id, suite_id, suite_name_list):

        for i in range(len(suite_name_list)):
            specify_suite_info_dict = self.get_speify_suite_sub_suite_info(suite_id)
            print(specify_suite_info_dict)

            if suite_name_list[i] not in specify_suite_info_dict.keys():
                suite_name_list = suite_name_list[i:]
                parent_suite_id = suite_id

                for name in suite_name_list:
                    parent_suite_id = TLH.create_suite(project_id, name, parent_id=parent_suite_id)
                return parent_suite_id
            elif suite_name_list[i] in specify_suite_info_dict.keys():
                suite_id = specify_suite_info_dict[suite_name_list[i]]
                if i == len(suite_name_list) - 1:
                    return suite_id


class ExcelTestCaseHandle(object):

    def __init__(self):
        pass

    def get_single_testcase(self, testcase_file_path):
        # 从Ｅxcel中获取单条case的信息
        all_cases = []
        wb = load_workbook(testcase_file_path)
        sheets_name_list = wb.sheetnames
        print(sheets_name_list)
        if len(sheets_name_list) == 1:
            ws = wb[sheets_name_list[0]]
            # print(ws.max_row)
            for row in range(2, ws.max_row + 1):
                single_case = []
                for column in range(1, 7):
                    # print(ws.cell(row=i, column=j).value)
                    single_case.append(ws.cell(row=row, column=column).value)
                all_cases.append(single_case)
        return all_cases, sheets_name_list[0]

    def handle_multi_suite_name(self, suites_name):
        eval_suite_name_list = []
        suite_name_list = eval(suites_name)
        for name in suite_name_list:
            if len(name) == 1:
                eval_suite_name_list.append(name[0])
        return eval_suite_name_list

    def handle_testcase_step(self, step_action, step_result):

        def f(x):
            return f'<p>{x}</p>'

        def add_tag_to_object(old_object):
            new_object = []
            for obj in old_object:
                if '\n' in obj:
                    obj_split = list(filter(lambda x: x, re.split(r'\n', obj)))
                    # print(obj_split)
                    new_obj_split = list(map(f, obj_split))
                    # print(new_obj_split)
                    new_object.append('\n'.join(new_obj_split))
                else:
                    new_object.append(f(obj))
            return new_object

        steps = []
        # 处理步骤数＋步骤内容
        action_split = re.split(r'步骤\d+、', step_action)
        list_action_split = list(filter(lambda x: x, action_split))
        new_list_action_split = add_tag_to_object(list_action_split)
        # 处理步骤数＋期望的结果内容
        result_split = re.split(r'步骤\d+、', step_result)
        list_result_split = list(filter(lambda x: x, result_split))
        new_list_result_split = add_tag_to_object(list_result_split)

        if len(new_list_action_split) == len(new_list_result_split):
            for numb in range(len(new_list_action_split)):
                step_info = dict()
                step_info['step_number'] = numb + 1
                step_info['actions'] = new_list_action_split[numb]
                step_info['expected_results'] = new_list_result_split[numb]
                step_info['execution_type'] = 1
                steps.append(step_info)
            return steps
        else:
            return False


logging_info_setting()
TLH = TestLinkHandle()
ETCH = ExcelTestCaseHandle()
projects = TLH.get_all_projects()
for project in projects:
    print(project)

specify_project_name = 'test_for_using_testlink'
pj_id = TLH.get_project_id_by_name(specify_project_name)
print(f'{specify_project_name}的id为{pj_id}')

pj_first_suites = TLH.get_suites(pj_id)


# 单条case的所有信息和当前模块名称
all_cases_info, module_name = ETCH.get_single_testcase('播放控制1.xlsx')
# print(all_cases_info)
# print(len(all_cases_info))
for single_case_info in all_cases_info:
    print(single_case_info)
    # print(len(single_case_info))
    # for i in single_case_info:
    #     print(type(i))
    #     print(i)

    # 单条case中所有的套件名称
    handle_suite_name_list = ETCH.handle_multi_suite_name(single_case_info[0])
    # print(handle_suite_name_list)

    # 单条case中的用例标题名称
    single_case_name = single_case_info[1]
    # print(single_case_name)

    # 单条case中用例的摘要
    single_case_summary = single_case_info[2]

    # 单条case中用例的前提
    single_case_preconditions = single_case_info[3]


    # def f(x):
    #     return f'<p>{x}</p>'
    # action_split = re.split(r'步骤\d+、', single_case_info[-2])
    # print(action_split)
    # print(list(filter(lambda x: x, action_split)))
    # print(len(list(filter(lambda x: x, action_split))))
    # action_split_list = list(filter(lambda x: x, action_split))
    # new_action_split_list = []
    # for action in action_split_list:
    #     if '\n' in action:
    #         action_split = list(filter(lambda x: x, re.split(r'\n', action)))
    #         print(action_split)
    #         new_action_split = list(map(f, action_split))
    #         print(new_action_split)
    #         new_action_split_list.append('\n'.join(new_action_split))
    #     else:
    #         new_action_split_list.append(f(action))
    # print(new_action_split_list)

    # result_split = re.split(r'步骤\d+、', single_case_info[-1])
    # print(result_split)
    # print(list(filter(lambda x: x, result_split)))
    # print(len(list(filter(lambda x: x, result_split))))
    # result_split_list = list(filter(lambda x: x, result_split))

    # 单个case中的所有步骤和结果
    single_case_steps_info = ETCH.handle_testcase_step(single_case_info[-2], single_case_info[-1])
    # print(single_case_steps_info)

    # 用例作者
    author_name = 'wangrun'

    module_suite_id = TLH.check_module_suite_exist(pj_id, module_name)
    print(module_suite_id)
    case_parent_suite_id = TLH.check_case_suite_exist(pj_id, module_suite_id, handle_suite_name_list)
    # # 创建模块名称套件和模块下的多重套件
    # parent_suite_id = TLH.create_suite(pj_id, module_name)
    # for name in handle_suite_name_list:
    #     parent_suite_id = TLH.create_suite(pj_id, name, parent_id=parent_suite_id)
    #     # if name == handle_suite_name_list[0]:
    #     #     parent_suite_id = TLH.create_suite(pj_id, name)
    #     # else:
    #     #     parent_suite_id = TLH.create_suite(pj_id, name, parent_id=parent_suite_id)
    #
    # 创建单个测试用例
    return_testcase_info = TLH.creat_cases_steps(single_case_steps_info, single_case_name, case_parent_suite_id, pj_id,
                                                 author_name, single_case_summary, single_case_preconditions)
    print(return_testcase_info)