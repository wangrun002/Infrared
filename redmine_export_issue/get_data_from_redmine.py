# !/usr/bin/python3
# -*- coding: UTF-8 -*-

from redminelib import Redmine


def print_all_project_base_info():
    projects = redmine.project.all()
    for project in projects:
        print(project.id, '---->', project.name, '---->', project.identifier)   # id, name, 标识符
        # print(project)


def get_project_id_by_name(name):
    projects = redmine.project.all()
    for project in projects:
        # print(project.id, '---->', project.name)
        if project.name == name:
            return project.id


redmine_url = 'http://git.nationalchip.com/redmine'
api_key = '53df3675279864992b561cb4c8af488a2d36c1a0'
# api_key = '994d302b537ddaddb170058885da1afb4f4acf57'

redmine = Redmine(redmine_url, key=api_key)

# 打印redmine上所有的项目的基本信息
# print_all_project_base_info()

specify_project_name = 'BU1-SDK-2007-01-UNIFY'
specify_project_id = get_project_id_by_name(specify_project_name)
print(specify_project_id)

pj = redmine.project.get(specify_project_id)
issues = pj.issues  # 默认情况下筛选打开的所有的issue
# print(len(issues))
for issue in issues:
    print(issue.keys())


