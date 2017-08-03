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

        # Channel List
        self.urls_list = []
        self.users_list = []
        self.wt_include_list = []
        self.wt_include_exclusive_list = []
        self.wt_exclude_list = []
        self.metadata_format_list = []

        # Final Track List
        self.track_list = []
        self.track_list_export = []

        # Youtube
        self.DEVELOPER_KEY = 'AIzaSyCP4gsM87jyGOJSexWastRbUq1n1Rk92zQ'
        self.YOUTUBE_API_SERVICE_NAME = 'youtube'
        self.YOUTUBE_API_VERSION = 'v3'
        self.youtube = build(self.YOUTUBE_API_SERVICE_NAME, self.YOUTUBE_API_VERSION, developerKey=self.DEVELOPER_KEY,
                             cache_discovery=False)

        # Words
        self.all_wt_include = [u'(Official Video)', u'(Official Music Video)', u'(Official Lyric Video)',
                               u'(official music video)', u'(Official Audio)',
                               u'(Music Video)', u'(MUSIC VIDEO)', u'(music video)', u'(Official MV)',
                               u'(MV)', u'(Full Ver.)', u'【MV full】', u'【MV】', u'［Radio Edit］',
                               u'［YouTube Ver.］', u'【Music Video】',
                               u'Official Music Video', u'Official Video', u'Official Lyric Video',
                               u'Official MV', u'official music video', u'Official Audio',
                               u'Music Video', u'MUSIC VIDEO', u'music video',
                               u'MV', u'Full Ver.', u'Full ver']
        self.all_wt_exclude = [u'Teaser', u'Digest Movie', u'LIVE Music Video', u'Recording Movie',
                               u'LIVE DVD', u'DVD&Blu-ray', u'SPOT', u'Short ver.', u'Tour', u'(Live at',
                               u'TOUR', u'Trailer', u'Blu-ray', u'Interview', u'tour', u'TALK SESSION',
                               u'launch party', u'DIGEST MOVIE', u'DVD', u'teaser', u'short ver.',
                               u'Live Video', u'trailer movie', u'short version', u'TVSPOT',
                               u'VIDEO CLIP SHORT', u'Digest movie', u'Web SPOT', u'(Live)']
        self.metadata_models = [u'Artist「Title」',
                                u'Artist - Title',
                                u'Artist / Title',
                                u'Title / Artist',
                                u'Artist -「Title」',
                                u'Artist-『Title』',
                                u'"Title" - Artist',
                                u'Title Artist',
                                u'Artist『Title』',
                                u'Artist "Title"',
                                u'Artist “Title"',
                                u'Artist /『Title』',
                                u'【Artist】Title']
        self.artists = [u'Mondo Grosso',
                        u'Nulbarich',
                        u'ゆず',
                        u'スピッツ',
                        u'AKB48[公式]',
                        u'GLAY',
                        u'Mr. Children',
                        u'サカナクション',
                        u'The Yellow Monkey',
                        u'安室奈美恵',
                        u'藤原さくら',
                        u'Dragon Ash',
                        u'Sekai No Owari',
                        u'The Peggies',
                        u'Tokyo Ska Paradise Orchestra',
                        u'Aiko',
                        u"B'Z",
                        u'Bump Of Chicken',
                        u'Go!Go!Vanillas',
                        u'I Dont Like Mondays',
                        u'Little Glee Monster',
                        u'Ryu Matsuyama',
                        u'Spicysol',
                        u'Yogee New Waves',
                        u'星野源',
                        u'Ai',
                        u'Blue Encount',
                        u'Cocco',
                        u'Five New Old',
                        u'イトヲカシ',
                        u'K',
                        u'Love Psychedelico',
                        u'Mrs. Green Apple',
                        u'くるり',
                        u'Radwimps',
                        u'サンボマスタ',
                        u'Suchmos',
                        u'Tuxedo',
                        u'Yuki',
                        u'三浦大知',
                        u'岡崎体育',
                        u'平井 堅',
                        u'!!! (Chk Chik Chick)',
                        u'9nm Parabellum Bullet',
                        u'赤い公園',
                        u'Androp',
                        u'D.A.N.',
                        u'Dreams Come True',
                        u'DYGL',
                        u'E-Girls',
                        u'Exile The Second',
                        u'Flumpool',
                        u'Hy',
                        u'Kyary Pamyu Pamyu',
                        u'Lulu X',
                        u'Miwa',
                        u'Mol-74',
                        u'Monkey Majik',
                        u'Moshimo',
                        u'ねごと',
                        u'Noriyuki Makihara',
                        u'Porno Graffitti',
                        u'Puffy',
                        u'シンリズム',
                        u'阪本奨悟',
                        u'Superfly',
                        u'宇多田ヒカル',
                        u'Wanima',
                        u'水曜日のカンパネラ',
                        u'九十九',
                        u'怒髪天',
                        u'私立恵比寿中学',
                        u'Sumika']

    def load_channel_list(self):
        file_path = 'youtube_channel_list.xls'
        xls = xlrd.open_workbook(file_path, formatting_info=True)
        self.n_rows = xls.sheet_by_index(0).nrows
        self.sheet_to_read = xls.sheet_by_index(0)
        self.list_xls = copy(xls)
        self.sheet = self.list_xls.get_sheet(0)
        for row in range(1, self.n_rows):
            self.users_list.append(self.sheet_to_read.cell(row, 0).value)
            self.urls_list.append(self.sheet_to_read.cell(row, 1).value)
            self.wt_include_list.append(self.sheet_to_read.cell(row, 2).value)
            self.wt_include_exclusive_list.append(self.sheet_to_read.cell(row, 3).value)
            self.wt_exclude_list.append(self.sheet_to_read.cell(row, 4).value)
            self.metadata_format_list.append(self.sheet_to_read.cell(row, 5).value)
        for metadata_format in self.metadata_format_list:
            if ',' in metadata_format:
                formats = metadata_format.split(',')
                new_formats = []
                for format in formats:
                    new_formats.append(format.rstrip().lstrip())
                index = self.metadata_format_list.index(metadata_format)
                self.metadata_format_list[index] = new_formats

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
                        query_url = 'https://www.googleapis.com/youtube/v3/videos?id=' + video_id  \
                                    + '&part=contentDetails&key=' + self.DEVELOPER_KEY
                        response = requests.get(query_url)
                        video_data = response.json()
                        duration = video_data['items'][0]['contentDetails']['duration'][2:]
                        self.track_list.append([title,
                                                video_url,
                                                duration,
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
            # artist = sheet_to_read.cell(row, 1).value
            url = sheet_to_read.cell(row, 2).value
            duration = 'NONE'
            yt_channel = sheet_to_read.cell(row, 3).value
            yt_url = sheet_to_read.cell(row, 4).value
            self.track_list.append([title, url, duration, yt_channel, yt_url])

    def extract_title_data(self):
        prev_user_index = -1
        for track in self.track_list:
            print('FOR TITLE -> ' + track[0])
            # Get User Index
            user_index = self.users_list.index(track[3])
            # Init Artist and Title
            title = ''
            if prev_user_index != user_index:
                artist = ''
            model_ids = []
            model_id = None
            # Get Metadata Model for this Track's User
            for model in self.metadata_models:
                if not isinstance(self.metadata_format_list[user_index], list):
                    if self.metadata_format_list[user_index] == model:
                        model_id = self.metadata_models.index(model)
                        break
                else:
                    for format in self.metadata_format_list[user_index]:
                        if format == model:
                            model_id = -1
                            model_ids.append(self.metadata_models.index(model))
            if model_id is not None:
                if model_id == -1:
                    format_index = 0
                    several_formats = True
                else:
                    several_formats = False

                while True:
                    if several_formats:
                        model_id = model_ids[format_index]

                    metadata_model = self.metadata_models[model_id]

                    if 'Title' in metadata_model.split('Artist')[1]:
                        id_artist = 0
                        id_title = 1
                    elif 'Title' in metadata_model.split('Artist')[0]:
                        id_title = 0
                        id_artist = 1

                    m_split = metadata_model.split('Artist')[id_title].split('Title')

                    raw_title = track[0]

                    get_title = True

                    for word in self.all_wt_exclude:
                        if word in raw_title:
                            get_title = False
                            break

                    if get_title:
                        if len(self.wt_include_exclusive_list[user_index]) > 0 and \
                                        self.wt_include_exclusive_list[user_index] not in raw_title:
                            get_title = False

                    if get_title:
                        artist = self.artists[user_index]
                        has_key_word = False
                        key_words = []
                        for word in self.all_wt_include:
                            if word in raw_title:
                                has_key_word = True
                                key_words.append(word)
                        if '"' in raw_title and '"Title"' in metadata_model:
                            if model_id == 6 and len(raw_title.split('"')) > 2:
                                title = raw_title.split('"')[1]
                                # if prev_user_index != user_index or artist == '':
                                #     artist = raw_title.split('"')[2]
                                #     if has_key_word:
                                #         for key_word in key_words:
                                #             if key_word in artist:
                                #                 artist = artist.replace(key_word, '').rstrip().lstrip()
                            elif model_id == 9 and len(raw_title.split('"')) > 1:
                                # if prev_user_index != user_index or artist == '':
                                #     artist = raw_title.split('"')[0].rstrip().lstrip()
                                title = raw_title.split('"')[1]
                        else:
                            # if prev_user_index != user_index or artist == '':
                            #     artist = raw_title.split(m_split[id_artist].rstrip().lstrip())
                            #     if (id_artist == 1 and len(artist) > 1) or id_artist == 0:
                            #         artist = artist[id_artist].rstrip()
                            #     else:
                            #         artist = None

                            if len(m_split) == 2:
                                if m_split[0] in raw_title and m_split[1] in raw_title:
                                    if m_split[1] != '':
                                        title = raw_title[raw_title.find(m_split[0])
                                                          + len(m_split[0]):raw_title.find(m_split[1])]
                                    else:
                                        title = raw_title[raw_title.find(m_split[0]) + len(m_split[0]):]
                                    if has_key_word:
                                        for key_word in key_words:
                                            if key_word in title:
                                                title = title.replace(key_word, '').rstrip().lstrip()

                    if several_formats and title == '' and format_index + 1 < len(model_ids):
                        format_index += 1
                    else:
                        break

            prev_user_index = user_index
            if title == '':
                artist_ = ''
            else:
                artist_ = artist
            self.track_list_export.append([title, artist_])

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
        sheet_.write(0, 4, 'Duration')
        sheet_.write(0, 5, 'Channel or User')
        sheet_.write(0, 6, 'Channel URL')
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
            sheet_.write(row, 6, track[4])
            row += 1

        # Write metadata
        self.list_xls.save('yt_tracks_.xls')
        list_xls_.save('yt_tracks_wChannelInfo_.xls')

    def get_all_tracks(self):
        self.load_channel_list()
        self.get_youtube_data()
        # self.debug_get_titles()
        self.extract_title_data()
        self.export_tracks_data()

bmat_japan = BmatJapan()
bmat_japan.get_all_tracks()
