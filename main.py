# -*- coding: utf-8 -*-
import xlrd, xlwt
from xlutils.copy import copy
import requests
import time
import argparse
import os

from googleapiclient.discovery import build


class BmatJapan(object):

    def __init__(self):
        # Sheet/XLS
        self.input_file_path = ''
        self.n_rows = 0
        self.n_cols = 0
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
        self.get_artist_from_input = False
        self.jp_artists = [u'Mondo Grosso',
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
        xls = xlrd.open_workbook(self.input_file_path, formatting_info=True)
        self.n_rows = xls.sheet_by_index(0).nrows
        self.n_cols = xls.sheet_by_index(0).ncols
        if self.n_rows > 1 and self.n_cols == 6:
            file_format_is_correct = True
            self.sheet_to_read = xls.sheet_by_index(0)
            self.list_xls = copy(xls)
            self.sheet = self.list_xls.get_sheet(0)
            for row in range(1, self.n_rows):
                self.users_list.append(self.sheet_to_read.cell(row, 0).value)
                self.urls_list.append(self.sheet_to_read.cell(row, 1).value)
                self.wt_include_list.append(self.sheet_to_read.cell(row, 2).value)
                self.wt_include_exclusive_list.append(self.sheet_to_read.cell(row, 3).value)
                self.wt_exclude_list.append(self.sheet_to_read.cell(row, 4).value)
                metadata_format = self.sheet_to_read.cell(row, 5).value
                if 'artist' in metadata_format:
                    metadata_format = metadata_format.replace('artist', 'Artist')
                if 'title' in metadata_format:
                    metadata_format = metadata_format.replace('title', 'Title')
                if '\\' in metadata_format:
                    metadata_format = metadata_format.replace('\\', '')
                self.metadata_format_list.append(metadata_format)
            for wt_include in self.wt_include_list:
                if wt_include != '':
                    if ',' in wt_include:
                        wts = wt_include.split(',')
                        for wt in wts:
                            if wt != '':
                                self.all_wt_include.append(wt)
                    else:
                        self.all_wt_include.append(wt_include)
            for wt_exclude in self.wt_exclude_list:
                if wt_exclude != '':
                    if ',' in wt_exclude:
                        wts = wt_exclude.split(',')
                        for wt in wts:
                            if wt != '':
                                self.all_wt_exclude.append(wt)
                    else:
                        self.all_wt_exclude.append(wt_exclude)
            for metadata_format in self.metadata_format_list:
                if ',' in metadata_format:
                    formats = metadata_format.split(',')
                    new_formats = []
                    for f in formats:
                        if 'artist' in f:
                            f = f.replace('artist', 'Artist')
                        if 'title' in f:
                            f = f.replace('title', 'Title')
                        new_formats.append(f.rstrip().lstrip())
                        if f not in self.metadata_models:
                            self.metadata_models.append(f)
                    index = self.metadata_format_list.index(metadata_format)
                    self.metadata_format_list[index] = new_formats
        else:
            file_format_is_correct = False

        return file_format_is_correct

    def get_youtube_data(self):
        for youtube_url in self.urls_list:
            if youtube_url != '':
                if 'channel' in youtube_url or 'user' in youtube_url:
                    print('user->' + youtube_url)
                    query = ''
                    if youtube_url[24:31] == 'channel':
                        query = 'id=' + youtube_url[32:]
                    elif youtube_url[24:28] == 'user':
                        query = 'forUsername=' + youtube_url[29:]
                    query_url = 'https://www.googleapis.com/youtube/v3/channels?part=contentDetails&' \
                                + query + '&key=' + self.DEVELOPER_KEY
                    response = requests.get(query_url)
                    data = response.json()
                    playlist_id = data[u'items'][0][u'contentDetails'][u'relatedPlaylists'][u'uploads']

                elif 'watch' in youtube_url and 'list' in youtube_url:
                    playlist_id = youtube_url.split('list=')[1]

                page_token = ''
                while True:
                    print('retrieving playlist page')
                    query_url = 'https://www.googleapis.com/youtube/v3/playlistItems?' + page_token \
                                + 'part=snippet&playlistId=' + playlist_id \
                                + '&maxResults=50' + '&key=' + self.DEVELOPER_KEY
                    while True:
                        try:
                            response = requests.get(query_url)
                            break
                        except:
                            print('problem retrieving playlist page')
                            time.sleep(10)
                            continue
                    data = response.json()
                    if len(data[u'items']) > 0:
                        for item in data['items']:
                            title = item['snippet']['title']
                            video_id = item['snippet']['resourceId']['videoId']
                            video_url = 'https://www.youtube.com/watch?v=' + video_id
                            query_url = 'https://www.googleapis.com/youtube/v3/videos?id=' + video_id \
                                        + '&part=contentDetails&key=' + self.DEVELOPER_KEY
                            while True:
                                try:
                                    print('retrieving video data with title -> ' + title)
                                    response = requests.get(query_url)
                                    break
                                except:
                                    print('problem retrieving video data')
                                    time.sleep(10)
                                    continue
                            video_data = response.json()
                            duration = video_data['items'][0]['contentDetails']['duration'][2:]
                            duration_in_secs = self.get_duration_in_secs(duration)
                            self.track_list.append([title,
                                                    video_url,
                                                    duration,
                                                    duration_in_secs,
                                                    self.users_list[self.urls_list.index(youtube_url)],
                                                    youtube_url])
                    if u'nextPageToken' in data.keys():
                        print('going to next page of this playlist')
                        page_token = 'pageToken=' + data[u'nextPageToken'] + '&'
                    else:
                        print('going to next user/channel')
                        break

    @staticmethod
    def get_duration_in_secs(duration):
        duration_in_secs = 0
        if 'H' in duration:
            hours_in_secs = int(duration.split('H')[0]) * 3600
            duration_in_secs += hours_in_secs
            if 'M' in duration or 'S' in duration:
                duration = duration.split('H')[1]
        if 'M' in duration:
            mins_in_secs = int(duration.split('M')[0]) * 60
            duration_in_secs += mins_in_secs
            if 'S' in duration:
                duration = duration.split('M')[1]
        if 'S' in duration:
            secs = int(duration.split('S')[0])
            duration_in_secs += secs
        return str(duration_in_secs)

    def debug_get_titles(self, file_name, get_dur_in_secs):
        file_path = file_name + '.xls'
        xls = xlrd.open_workbook(file_path, formatting_info=True)
        rows = xls.sheet_by_index(0).nrows
        sheet_to_read = xls.sheet_by_index(0)

        if get_dur_in_secs:
            list_xls = xlwt.Workbook()
            sheet = list_xls.add_sheet('Youtube List')
            sheet.write(0, 0, 'Video Title')
            sheet.write(0, 1, 'Track Title')
            sheet.write(0, 2, 'Track Artist')
            sheet.write(0, 3, 'Youtube URL')
            sheet.write(0, 4, 'Duration')
            sheet.write(0, 5, 'Duration (secs)')
            sheet.write(0, 6, 'Channel/User/Playlist')
            sheet.write(0, 7, 'Channel URL')

        for row in range(1, rows):
            video_title = sheet_to_read.cell(row, 0).value
            title = sheet_to_read.cell(row, 1).value
            artist = sheet_to_read.cell(row, 2).value
            url = sheet_to_read.cell(row, 3).value
            duration = sheet_to_read.cell(row, 4).value
            yt_channel = sheet_to_read.cell(row, 5).value
            yt_url = sheet_to_read.cell(row, 6).value
            if get_dur_in_secs:
                duration_in_secs = self.get_duration_in_secs(duration)
                track = [video_title, title, artist, url, duration, duration_in_secs, yt_channel, yt_url]
                for col in range(0, 8):
                    sheet.write(row, col, track[col])
                list_xls.save(file_name + '_.xls')
                print('Row: ' + str(row) + ' // Total Rows: ' + str(rows))
                print('Track: ' + video_title + ' // SAVED.')
            else:
                self.track_list.append([video_title, url, duration, yt_channel, yt_url])

    def export_prev_ver(self, file_name):
        list_xls_ = xlwt.Workbook()
        sheet_ = list_xls_.add_sheet('Youtube List')

        # Headers
        sheet_.write(0, 0, 'Video Title')
        sheet_.write(0, 1, 'Track Title')
        sheet_.write(0, 2, 'Track Artist')
        sheet_.write(0, 3, 'Youtube URL')
        sheet_.write(0, 4, 'Duration')
        sheet_.write(0, 5, 'Duration (secs)')
        sheet_.write(0, 6, 'Channel/User/Playlist')
        sheet_.write(0, 7, 'Channel URL')
        row = 1

        # Add tracks
        for track in self.track_list:
            sheet_.write(row, 0, track[0])
            sheet_.write(row, 1, '')
            sheet_.write(row, 2, '')
            sheet_.write(row, 3, track[1])
            sheet_.write(row, 4, track[2])
            sheet_.write(row, 5, track[3])
            sheet_.write(row, 6, track[4])
            sheet_.write(row, 7, track[5])
            row += 1

        # Write metadata
        list_xls_.save(file_name + '.xls')

    def extract_title_data(self):
        prev_user_index = -1
        for track in self.track_list:
            # Get User Index
            user_index = self.urls_list.index(track[5])
            # Init Artist and Title
            title = ''
            if prev_user_index != user_index:
                artist = ''
            format_id = 0
            while True:
                if not isinstance(self.metadata_format_list[user_index], list):
                    metadata_model = self.metadata_format_list[user_index]
                else:
                    metadata_model = self.metadata_format_list[user_index][format_id]

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
                    if word.lower() in raw_title.lower():
                        get_title = False
                        break

                if get_title:
                    if len(self.wt_include_exclusive_list[user_index]) > 0 and \
                                    self.wt_include_exclusive_list[user_index] not in raw_title:
                        get_title = False

                if get_title:
                    if self.get_artist_from_input:
                        artist = self.users_list[user_index]
                    has_key_word = False
                    key_words = []
                    for word in self.all_wt_include:
                        if word in raw_title:
                            has_key_word = True
                            key_words.append(word)
                    if ('"' in raw_title and '"Title"' in metadata_model) \
                            or ("'" in raw_title and "'Title'" in metadata_model):
                        if (id_title == 0 and len(raw_title.split('"')) > 2) \
                                or (id_title == 0 and len(raw_title.split("'")) > 2):
                            if '"' in raw_title:
                                title = raw_title.split('"')[1]
                            elif "'" in raw_title:
                                title = raw_title.split("'")[1]
                            if not self.get_artist_from_input:
                                if '"' in raw_title:
                                    artist = raw_title.split('"')[2]
                                elif "'" in raw_title:
                                    artist = raw_title.split("'")[2]
                                if has_key_word:
                                    for key_word in key_words:
                                        if key_word in artist:
                                            artist = artist.replace(key_word, '').rstrip().lstrip()
                        elif (id_title == 1 and len(raw_title.split('"')) > 1) \
                                or (id_title == 1 and len(raw_title.split("'")) > 1):
                            if not self.get_artist_from_input:
                                if '"' in raw_title:
                                    artist = raw_title.split('"')[0]
                                elif "'" in raw_title:
                                    artist = raw_title.split("'")[0]
                                if '-' in artist:
                                    artist = artist.replace('-', '')
                                artist = artist.rstrip().lstrip()
                            if '"' in raw_title:
                                title = raw_title.split('"')[1]
                            elif "'" in raw_title:
                                title = raw_title.split("'")[1]
                    else:
                        if not self.get_artist_from_input:
                            artist = raw_title.split(m_split[id_artist].rstrip().lstrip())
                            if (id_artist == 1 and len(artist) > 1) or id_artist == 0:
                                artist = artist[id_artist].rstrip()
                            else:
                                artist = None

                        if len(m_split) == 2:
                            if m_split[0] in raw_title and m_split[1] in raw_title:
                                if m_split[1] != '':
                                    title = raw_title[raw_title.find(m_split[0]) +
                                                      len(m_split[0]):raw_title.find(m_split[1])]
                                else:
                                    title = raw_title[raw_title.find(m_split[0]) +
                                                      len(m_split[0]):]
                                if has_key_word:
                                    for key_word in key_words:
                                        if key_word in title:
                                            title = title.replace(key_word, '').rstrip().lstrip()

                if isinstance(self.metadata_format_list[user_index], list) and \
                    len(self.metadata_format_list[user_index]) > 1 and title == '' \
                        and format_id + 1 < len(self.metadata_format_list[user_index]):
                    format_id += 1
                else:
                    break

            prev_user_index = user_index
            if title == '':
                artist_ = ''
                print('did not extract metadata for video title -> ' + track[0])
            else:
                artist_ = artist
                print('metadata extracted // track title -> ' + title + ' // track artist -> ' + artist_)
            self.track_list_export.append([title, artist_])

    def export_tracks_data(self, file_name):
        # Open workbook and add sheet
        list_xls_ = xlwt.Workbook()
        sheet_ = list_xls_.add_sheet('Youtube List')

        # Headers
        sheet_.write(0, 0, 'Video Title')
        sheet_.write(0, 1, 'Track Title')
        sheet_.write(0, 2, 'Track Artist')
        sheet_.write(0, 3, 'Youtube URL')
        sheet_.write(0, 4, 'Duration')
        sheet_.write(0, 5, 'Duration (secs)')
        sheet_.write(0, 6, 'Channel/User/Playlist')
        sheet_.write(0, 7, 'Channel URL')
        row = 1

        # Add tracks
        sheet_index = 1
        include_all_tracks = False
        for track in self.track_list:
            write_track = True
            title = self.track_list_export[self.track_list.index(track)][0]
            artist = self.track_list_export[self.track_list.index(track)][1]
            if not include_all_tracks:
                if title == '' or artist == '':
                    write_track = False
            if write_track:
                sheet_.write(row, 0, track[0])
                sheet_.write(row, 1, title)
                sheet_.write(row, 2, artist)
                sheet_.write(row, 3, track[1])
                sheet_.write(row, 4, track[2])
                sheet_.write(row, 5, track[3])
                sheet_.write(row, 6, track[4])
                sheet_.write(row, 7, track[5])
                row += 1
                if row == 60002:
                    sheet_index += 1
                    sheet_ = list_xls_.add_sheet('Youtube List (' + str(sheet_index) + ')')
                    sheet_.write(0, 0, 'Video Title')
                    sheet_.write(0, 1, 'Track Title')
                    sheet_.write(0, 2, 'Track Artist')
                    sheet_.write(0, 3, 'Youtube URL')
                    sheet_.write(0, 4, 'Duration')
                    sheet_.write(0, 5, 'Duration (secs)')
                    sheet_.write(0, 6, 'Channel/User/Playlist')
                    sheet_.write(0, 7, 'Channel URL')
                    row = 1

        # Write metadata
        list_xls_.save(file_name)

    def get_all_tracks(self, file_path, get_artist_from_input=False):
        self.input_file_path = file_path
        self.get_artist_from_input = get_artist_from_input

        file_is_correct = self.load_channel_list()
        if file_is_correct:
            self.get_youtube_data()
            self.extract_title_data()
            self.export_tracks_data(file_path[:-4] + '_output.xls')
        else:
            print('There was a problem with the format of the xls file.')


class InputParser(object):
    def __init__(self):
        self.parser = argparse.ArgumentParser(description='Get links and metadata from youtube channel/user/playlist.')
        self.file_path = None

    def add_arguments(self):
        # Arguments to parse
        self.parser.add_argument('-input_file_path', action='store', dest='file_path', default=None,
                                 help='path for input xls file')

    def parse_input(self):
        # Add arguments to parser
        self.add_arguments()
        # Parse arguments from input
        args = self.parser.parse_args()

        # Get Input File Path if exists
        if args.file_path is not None:
            if os.path.exists(args.file_path):
                self.file_path = args.file_path
            else:
                print('The path specified as input_file_path does not exist.')
        else:
            print('Please specify a valid input_file_path.')


args_input = True

if args_input:
    # Get input from user and parse arguments
    input_parser = InputParser()
    input_parser.parse_input()
    # Get tracks
    if input_parser.file_path is not None:
        bmat_japan = BmatJapan()
        bmat_japan.get_all_tracks(input_parser.file_path)
else:
    path = 'KoreanValidYTChannels.xls'
    bmat_japan = BmatJapan()
    bmat_japan.get_all_tracks(path)
