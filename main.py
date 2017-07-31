import os
import xlrd, xlwt
from xlutils.copy import copy

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class BmatJapan(object):

    def __init__(self):
        # Sheet/XLS
        self.n_rows = 0
        self.sheet_to_read = None
        self.list_xls = None
        self.sheet = None
        self.list_urls = []

        # Youtube
        self.DEVELOPER_KEY = 'AIzaSyCP4gsM87jyGOJSexWastRbUq1n1Rk92zQ'
        self.YOUTUBE_API_SERVICE_NAME = 'youtube'
        self.YOUTUBE_API_VERSION = 'v3'
        self.youtube = build(self.YOUTUBE_API_SERVICE_NAME, self.YOUTUBE_API_VERSION, developerKey=self.DEVELOPER_KEY,
                             cache_discovery=False)

    def load_channel_list(self):
        file_path = 'youtube_channel_list.xls'
        xls = xlrd.open_workbook(file_path, formatting_info=True)
        self.n_rows = xls.sheet_by_index(0).nrows
        self.sheet_to_read = xls.sheet_by_index(0)
        self.list_xls = copy(xls)
        self.sheet = self.list_xls.get_sheet(0)
        for row in range(1, self.n_rows):
            url = self.sheet_to_read.cell(row, 1).value + '/videos?flow=list&view=0'
            self.list_urls.append(url)
        print('yeah')
        # example = 'https://www.googleapis.com/youtube/v3/search?order=date&part=snippet&channelId=UCCdibO9eRKJIre22DKDNqug&maxResults=25&key=AIzaSyCP4gsM87jyGOJSexWastRbUq1n1Rk92zQ'
        # get a json and then interpret

    def get_all_tracks(self):
        self.load_channel_list()

bmat_japan = BmatJapan()
bmat_japan.get_all_tracks()
