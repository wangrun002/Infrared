# !/usr/bin/python3
# -*- conding: UTF-8 -*-
# File Name:read_info_from_vimeo.py
# Created Time:2021-01-15 13:58:08

from config import access_token, client_id, client_secret
import vimeo
import requests
import json
import re

# client = vimeo.VimeoClient(
#     token=access_token,
#     key=client_id,
#     secret=client_secret
# )
#
# url = ''
#
# response = client.get(url)
#
# print(response.json())


class VimeoInfoHandle(object):

    def __init__(self, token, key, secret):
        self.base_url = 'https://api.vimeo.com/'
        self.token = token
        self.key = key
        self.secret = secret
        self.client = vimeo.VimeoClient(
            token=self.token,
            key=self.key,
            secret=self.secret
        )

    def get_html_to_json(self, path, timeout):
        # 组合url后get网页并将response转化为json
        api_url = f"{self.base_url}{path}"
        r = self.client.get(api_url, timeout=timeout)

        if r.status_code == requests.codes.ok:
            # data = json.dumps(r.json(), sort_keys=False, indent=4, separators=(',', ':'))
            data = r.json()

        else:
            data = None
        return data

    def get_channel_id_from_channel_name(self, channel_name):
        # 通过channel_name获取channel_id
        path = f'channels/{channel_name}'
        data = self.get_html_to_json(path, 10)
        data_fmt_json = json.dumps(data, sort_keys=False, indent=4, separators=(',', ':'))
        # print(data_fmt_json)
        try:
            channel_id_data = data["uri"]
            channel_id = re.split(r'/', channel_id_data)[-1]
        except KeyError:
            channel_id = None
        return channel_id

    def get_playList(self, channel_id):
        # 通过channel_id查找该频道下的video_id
        path = f'channels/{channel_id}/videos'
        data = self.get_html_to_json(path, 10)
        data_fmt_json = json.dumps(data, sort_keys=False, indent=4, separators=(',', ':'))
        # print(data_fmt_json)
        if not data:
            return {}

        datas_item = data['data']

        videos_info = {}
        for data_item in datas_item:
            print(data_item['uri'])
            video_id = re.split(r'/', data_item['uri'])[-1]
            videos_info[video_id] = []
            videos_info[video_id].append(data_item['name'])
            videos_info[video_id].append(data_item['duration'])
            videos_info[video_id].append(data_item['user']['name'])

        print(videos_info)
        for k, v in videos_info.items():
            print(k, v)
        return videos_info


def main():
    search_name = 'staffpicks'
    client = VimeoInfoHandle(access_token, client_id, client_secret)
    channel_id = client.get_channel_id_from_channel_name(search_name)
    print(channel_id)
    videos_info = client.get_playList(channel_id)


if __name__ == '__main__':
    main()

