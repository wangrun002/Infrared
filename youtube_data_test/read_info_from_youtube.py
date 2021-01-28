# !/usr/bin/python3
# -*- conding: UTF-8 -*-
# File Name:read_info_from_youtube.py
# Created Time:2021-01-06 14:17:30

from datetime import datetime
import requests
import logging


class YoutubeInfoHandle(object):

    def __init__(self, yt_key):
        self.base_url = "https://www.googleapis.com/youtube/v3/"
        self.api_key = yt_key

    def get_html_to_json(self, path):
        # 组合url后get网页并将response转化为json
        api_url = f"{self.base_url}{path}&key={self.api_key}"
        r = requests.get(api_url)

        if r.status_code == requests.codes.ok:
            data = r.json()
        else:
            data = None
        return data

    def get_channel_uploads_id(self, channel_id, part='contentDetails'):
        # 获取指定channel_id下的上传视频列表的id
        path = f"channels?part={part}&id={channel_id}"
        data = self.get_html_to_json(path)
        try:
            uploads_id = data["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]
        except KeyError:
            uploads_id = None
        return uploads_id

    def get_playlist(self, playlist_id, part="contentDetails", max_results=10):
        # 根据指定playlist_id获取该列表下的视频id
        path = f"playlistItems?part={part}&playlistId={playlist_id}&maxResults={max_results}"
        data = self.get_html_to_json(path)
        if not data:
            return []

        video_ids = []
        for data_item in data["items"]:
            video_ids.append(data_item["contentDetails"]["videoId"])
        return video_ids

    def get_video_info(self, video_id, part='snippet,statistics'):
        # 根据指定的视频id，获取视频的相关信息
        path = f'videos?part={part}&id={video_id}'
        data = self.get_html_to_json(path)
        if not data:
            return {}

        # 以下是按需提取出来的一些信息
        data_item = data["items"][0]
        try:
            # "2021-01-11T20:43:54Z"
            video_published_time = datetime.strptime(data_item["snippet"]["publishedAt"], "%Y-%m-%dT%H:%M:%SZ")
        except ValueError:  # 日期格式错误
            video_published_time = None

        video_url = f"https://www.youtube.com/watch?v={data_item['id']}"

        video_info = {
            "video_id": data_item['id'],
            "channel_title": data_item["snippet"]["channelTitle"],
            # "published_time": video_published_time,
            "video_title": data_item["snippet"]["title"],
            # "video_url": video_url
        }
        return video_info


def logging_info_setting():
    # 配置logging输出格式
    log_format = "%(asctime)s %(name)s %(levelname)s %(message)s"  # 配置输出日志的格式
    data_format = "%Y-%m-%d %H:%M:%S %a"  # 配置输出时间的格式
    logging.basicConfig(level=logging.DEBUG, format=log_format, datefmt=data_format)


def main():
    yt_key = 'AIzaSyA9fANtb9bDhmv6P0E7IWhU3xfJO4ebGV0'
    yt_channel_id = "UC2pmfLm7iq6Ov1UwYrWYkZA"
    hanle_yt_info = YoutubeInfoHandle(yt_key)

    uploads_id = hanle_yt_info.get_channel_uploads_id(yt_channel_id)
    print(f"根据channel_id:{yt_channel_id}获取到的上传视频列表id为：{uploads_id}")

    videos_id = hanle_yt_info.get_playlist(uploads_id, max_results=50)
    print(f"根据channel_id:{yt_channel_id}获取到的上传视频列表中的视频id列表为：\n{videos_id}")

    for video_id in videos_id:
        print("{:*^50}".format("分割线"))
        video_info = hanle_yt_info.get_video_info(video_id)
        print(video_info)


if __name__ == "__main__":
    # logging_info_setting()
    main()
