# -*- coding: utf-8 -*-
import xlrd, xlwt
from xlutils.copy import copy
import requests

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class BmatJapan(object):

    def __init__(self):
        # Sheet/XLS
        self.n_rows = 0
        self.sheet_to_read = None
        self.list_xls = None
        self.sheet = None
        self.urls_list = []
        self.users_list = []
        self.track_list = []
        self.track_list_export = []

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
        # TODO: DEBUG VERSION
        # for row in range(5, 10):
            url = self.sheet_to_read.cell(row, 1).value
            user = self.sheet_to_read.cell(row, 0).value
            self.urls_list.append(url)
            self.users_list.append(user)

    def get_youtube_data(self):
        for youtube_url in self.urls_list:
            query = ''
            if youtube_url[24:31] == 'channel':
                query = 'id=' + youtube_url[32:]
            elif youtube_url[24:28] == 'user':
                query = 'forUsername=' + youtube_url[29:]
            query_url = 'https://www.googleapis.com/youtube/v3/channels?part=contentDetails&' \
                        + query + '&key=' + self.DEVELOPER_KEY
            response = requests.get(query_url)
            data = response.json()
            uploads_playlist_id = data[u'items'][0][u'contentDetails'][u'relatedPlaylists'][u'uploads']
            page_token = ''
            while True:
                query_url = 'https://www.googleapis.com/youtube/v3/playlistItems?' + page_token \
                            + 'part=snippet&playlistId=' + uploads_playlist_id \
                            + '&maxResults=50' + '&key=' + self.DEVELOPER_KEY
                response = requests.get(query_url)
                data = response.json()
                if len(data[u'items']) > 0:
                    for item in data['items']:
                        title = item['snippet']['title']
                        video_id = item['snippet']['resourceId']['videoId']
                        video_url = 'https://www.youtube.com/watch?v=' + video_id
                        self.track_list.append([title,
                                                video_url,
                                                self.users_list[self.urls_list.index(youtube_url)],
                                                youtube_url])
                if u'nextPageToken' in data.keys():
                    page_token = 'pageToken=' + data[u'nextPageToken'] + '&'
                else:
                    break

    def debug_get_titles(self):
        file_path = 'yt_tracks_wChannelInfo.xls'
        xls = xlrd.open_workbook(file_path, formatting_info=True)
        rows = xls.sheet_by_index(0).nrows
        sheet_to_read = xls.sheet_by_index(0)
        for row in range(1, rows):
            title = sheet_to_read.cell(row, 0).value
            artist = sheet_to_read.cell(row, 1).value
            url = sheet_to_read.cell(row, 2).value
            yt_channel = sheet_to_read.cell(row, 3).value
            yt_url = sheet_to_read.cell(row, 4).value
            self.track_list.append([title, artist, url, yt_channel, yt_url])

    def extract_guion(self, raw_title):
        title = None
        artist = None
        if ' - ' in raw_title:
            raw_title = raw_title.split('-')
            artist = raw_title[0]
            title = raw_title[1]
            tabus = [u'(Official', u'［Teaser']
            for tabu in tabus:
                if tabu in title:
                    title = title.split(tabu)[0]
                    break
        return title, artist

    def extract_japarenthesis(self, raw_title):
        title = None
        artist = None
        if u'「' in raw_title:
            raw_title = raw_title.split(u'「')
            artist = raw_title[0]
            if 'LIVE' in artist:
                artist = artist.split('LIVE')[0]
            elif 'NEW' in artist:
                artist = artist.split('NEW')[0]
            title = raw_title[1].split(u'」')[0]
            if 'LIVE' in title:
                title = raw_title[1].split(u'」')[1]

        return title, artist

    def extract_barra(self, raw_title):
        title = None
        artist = None
        if '/' in raw_title:
            raw_title = raw_title.split(u'/')
            artist = raw_title[0]
            title = raw_title[1]
        return title, artist

    def extract_48(self, raw_title):
        title = None
        artist = None
        if raw_title[3:5] == '48':
            if raw_title[5] == ' ' and '(' in raw_title:
                raw_title_ = raw_title[6:]
                title = raw_title_.split(' ')[0]
                artist = raw_title_[raw_title_.find('(')+1:raw_title_.find(')')]
        return title, artist

    def extract_japarenthesis2(self, raw_title):
        title = None
        artist = None
        if u'『' in raw_title:
            artist = raw_title.split(u'『')[0]
            title = raw_title[raw_title.find(u'『')+1:raw_title.find(u'』')]
        return title, artist

    def extract_guion_largo(self, raw_title):
        title = None
        artist = None
        if u' ー ' in raw_title:
            raw_title = raw_title.split(' ー ')
            artist = raw_title[0]
            title = raw_title[1]
        return title, artist

    def extract_title_data(self):
        for track in self.track_list:
            # Get Artist and Title
            raw_title = track[0]
            title = None
            artist = None
            if track[3] == u'Nulbarich':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Yuzu':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Spitz':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Akb48':
                title, artist = self.extract_48(raw_title)
            elif track[3] == u'Glay':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Mr. Children':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Sakanaction':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'The Yellow Monkey':
                artist, title = self.extract_barra(raw_title)
            elif track[3] == u'Namie Amuro':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'藤原さくら':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Dragon Ash':
                title, artist = self.extract_guion(raw_title)
                if title and artist is None:
                    title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Sekai No Owari':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'The Peggies':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Tokyo Ska Paradise Orchestra':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Aiko':
                artist = 'Aiko'
                title, artist_ = self.extract_japarenthesis2(raw_title)
            elif track[3] == u"B'Z":
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Bump Of Chicken':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Go!Go!Vanillas':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'I Dont Like Mondays':
                artist, title = self.extract_guion(raw_title)
            elif track[3] == u'Little Glee Monster':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Ryu Matsuyama':
                pass
            elif track[3] == u'Spicysol':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Yogee New Waves':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'星野源':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Ai':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Blue Encount':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Cocco':
                title, artist = self.extract_guion(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Five New Old':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Itowokashi':
                pass
            elif track[3] == u'K':
                title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Love Psychedelico':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Mrs. Green Apple':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Quruli':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Radwimps':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Sambomaster':
                title, artist = self.extract_guion_largo(raw_title)
            elif track[3] == u'Tuxedo':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Yuki':
                title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'平井 堅':
                title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'!!! (Chk Chik Chick)':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'9Mm Parabellum Bullet':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Akai Ko-En':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Androp':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'D.A.N.':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Dreams Come True':
                title, artist = self.extract_guion(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Dygl':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Exile The Second':
                title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Flumpool':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_barra(raw_title)
            elif track[3] == u'Hy':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Kyary Pamyu Pamyu':
                title, artist = self.extract_guion_largo(raw_title)
            elif track[3] == u'Lulu X':
                pass
            elif track[3] == u'Miwa':
                title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Mol-74':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Monkey Majik':
                pass
            elif track[3] == u'Moshimo':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Negoto':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Noriyuki Makihara':
                pass
            elif track[3] == u'Porno Graffitti':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Puffy':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Shin Rizumu':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Shogo Sakamoto':
                title, artist = self.extract_guion(raw_title)
            elif track[3] == u'Superfly':
                title, artist = self.extract_guion(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis(raw_title)
                    if title and artist is None:
                        title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Utada Hikaru':
                title, artist = self.extract_guion(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Wanima':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'Wednesday Campanella':
                title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'九十九':
                title, artist = self.extract_japarenthesis(raw_title)
            elif track[3] == u'怒髪天':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'私立恵比寿中学':
                title, artist = self.extract_japarenthesis(raw_title)
                if title and artist is None:
                    title, artist = self.extract_japarenthesis2(raw_title)
            elif track[3] == u'Sumika':
                title, artist = self.extract_barra(raw_title)
            if title and artist is not None:
                artist = artist.rstrip().lstrip().title()
                title = title.rstrip().lstrip().title()
            self.track_list_export.append([title, artist])
        print('i have all tracks')

    def export_tracks_data(self):
        # Open workbook and add sheet
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet('Youtube List')
        list_xls_ = xlwt.Workbook()
        sheet_ = list_xls_.add_sheet('Youtube List')

        # Headers
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Youtube URL')
        sheet_.write(0, 0, 'Video Title')
        sheet_.write(0, 1, 'Track Title')
        sheet_.write(0, 2, 'Track Artist')
        sheet_.write(0, 3, 'Youtube URL')
        sheet_.write(0, 4, 'Channel or User')
        sheet_.write(0, 5, 'Channel URL')
        row = 1

        # Add tracks
        for track in self.track_list:
            self.sheet.write(row, 0, track[0])
            self.sheet.write(row, 1, '')
            self.sheet.write(row, 2, track[1])
            sheet_.write(row, 0, track[0])
            sheet_.write(row, 1, self.track_list_export[self.track_list.index(track)][0])
            sheet_.write(row, 2, self.track_list_export[self.track_list.index(track)][1])
            sheet_.write(row, 3, track[1])
            sheet_.write(row, 4, track[2])
            sheet_.write(row, 5, track[3])
            row += 1

        # Write metadata
        self.list_xls.save('yt_tracks_.xls')
        list_xls_.save('yt_tracks_wChannelInfo_.xls')

    def get_all_tracks(self):
        # self.load_channel_list()
        # self.get_youtube_data()
        self.debug_get_titles()
        self.extract_title_data()
        self.export_tracks_data()

bmat_japan = BmatJapan()
bmat_japan.get_all_tracks()
