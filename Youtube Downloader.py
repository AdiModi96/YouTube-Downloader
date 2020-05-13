import sys
import colorama
from colorama import Fore, Back, Style
from pytube import YouTube
from prettytable import PrettyTable
from winreg import *
import requests
import urllib3
from moviepy.editor import *


# Example Link: Slow-Mo: https://www.youtube.com/watch?v=YGKRraZXp_4

# The Progressive Streams have audio and video together
# The DASH Streams have separate audio and video, need to merge them post the download

def displayInformation(video):
    try:
        # Title
        print(Fore.YELLOW + 'Video Title: ' + Style.RESET_ALL + Fore.GREEN + '{}'.format(video.title) + Style.RESET_ALL)
    except:
        print(Fore.YELLOW + 'Video Title: ' + Style.RESET_ALL + Fore.LIGHTRED_EX + 'Unavailable' + Style.RESET_ALL)

    try:
        # Video Length
        print(Fore.YELLOW + 'Video Duration: ' + Style.RESET_ALL + Fore.GREEN + '{} minutes and {} seconds'.format(
            int(int(video.length) / 60), int(video.length) % 60) + Style.RESET_ALL)
    except:
        print(Fore.YELLOW + 'Video Duration: ' + Style.RESET_ALL + Fore.LIGHTRED_EX + 'Unavailable' + Style.RESET_ALL)

    try:
        # Video Views
        print(Fore.YELLOW + 'Video Views: ' + Style.RESET_ALL + Fore.GREEN + '{}'.format(video.views) + Style.RESET_ALL,
              end=' | ')
    except:
        print(Fore.YELLOW + 'Video Views: ' + Style.RESET_ALL + Fore.LIGHTRED_EX + 'Unavailable' + Style.RESET_ALL,
              end=' | ')

    try:
        # Video Rating
        print(
            Fore.YELLOW + 'Video Rating: ' + Style.RESET_ALL + Fore.GREEN + '{}'.format(
                round(video.rating, 3)) + Style.RESET_ALL)
    except:
        print(Fore.YELLOW + 'Video Rating: ' + Style.RESET_ALL + Fore.LIGHTRED_EX + 'Unavailable' + Style.RESET_ALL)


def displayStreams(video):
    print(Fore.YELLOW + 'Video Streams: ' + Style.RESET_ALL)
    printTable = PrettyTable([Fore.YELLOW + 'File Type' + Style.RESET_ALL,
                              Fore.YELLOW + 'ID' + Style.RESET_ALL,
                              Fore.YELLOW + 'Resolution' + Style.RESET_ALL,
                              Fore.YELLOW + 'FPS' + Style.RESET_ALL,
                              Fore.YELLOW + 'Audio Bitrate' + Style.RESET_ALL,
                              Fore.YELLOW + 'Extension' + Style.RESET_ALL])

    for stream in video.streams.all():
        fileType = stream.type
        fileExtension = stream.subtype

        if fileType == 'video':
            itag = stream.itag
            resolution = stream.resolution
            fps = stream.fps
            if resolution == None or fps == None:
                continue
            audioBitRate = stream.abr
            if audioBitRate == None:
                fileType = 'Video Only'
            else:
                fileType = 'Audio+Video'

            printTable.add_row([Fore.CYAN + str(fileType) + Style.RESET_ALL,
                                Fore.CYAN + str(itag) + Style.RESET_ALL,
                                Fore.CYAN + str(resolution) + Style.RESET_ALL,
                                Fore.CYAN + str(fps) + Style.RESET_ALL,
                                Fore.CYAN + str(audioBitRate) + Style.RESET_ALL,
                                Fore.CYAN + str(fileExtension) + Style.RESET_ALL])

        elif fileType == 'audio':
            itag = stream.itag
            resolution = stream.resolution
            fps = stream.fps
            audioBitRate = stream.abr
            if audioBitRate == None:
                continue
            fileType = 'Audio'
            printTable.add_row([Fore.CYAN + str(fileType) + Style.RESET_ALL,
                                Fore.CYAN + str(itag) + Style.RESET_ALL,
                                Fore.CYAN + str(resolution) + Style.RESET_ALL,
                                Fore.CYAN + str(fps) + Style.RESET_ALL,
                                Fore.CYAN + str(audioBitRate) + Style.RESET_ALL,
                                Fore.CYAN + str(fileExtension) + Style.RESET_ALL])

    print(printTable)


def displayError(errorMessage):
    print(
        Fore.RED + 'Error: ' + Fore.LIGHTRED_EX + errorMessage + Style.RESET_ALL)


def downloadStream(streams, downloadFolderPath):
    filePointers = []
    for stream in streams:
        print(Fore.YELLOW + 'Downloading Stream with ID: {}'.format(stream.itag) + Style.RESET_ALL)
        filePointers.append(stream.download(downloadFolderPath, filename_prefix=stream.itag + '_'))
        print(Back.GREEN + 'Downloaded!' + Style.RESET_ALL)
    return filePointers


def downloadYouTubeVideo(videoLink, downloadFolderPath):
    try:
        # Collecting Video Information
        video = YouTube(videoLink)
    except Exception as ex:
        # print(ex)
        displayError(
            'Couldn\'t connect to the Internet OR Either the URL is incorrect OR the video doesn\'t have download permissions'
        )
        sys.exit(1)

    displayInformation(video)
    displayStreams(video)

    itags = []
    streams = []
    print('Instructions: ')
    print('\t- Enter 1 ID value to download one file.')
    print('\t- Enter 2 ID values, separated by space Order: video audio, to download and merge.')
    while True:
        itagsString = input('Enter the ID of the stream(s) you want to download: ').split(' ')
        for itag in itagsString:
            if len(itag) != 0:
                itags.append(itag)

        if len(itags) >= 1 and len(itags) <= 2:
            for itag in itags:
                stream = video.streams.get_by_itag(itag)
                if stream == None:
                    print(Fore.LIGHTRED_EX + 'Incorrect ID: {}, Skipping'.format(itag) + Style.RESET_ALL)
                else:
                    streams.append(stream)
            break
        else:
            itags = []
            print(Fore.LIGHTRED_EX + 'Please Enter only 1 or 2 ID values' + Style.RESET_ALL)

    if downloadFolderPath == None:
        try:
            with OpenKey(HKEY_CURRENT_USER,
                         'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
                downloadFolderPath = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        except Exception as ex:
            downloadFolderPath = '.'

    elif not os.path.exists(downloadFolderPath):
        os.makedirs(downloadFolderPath)

    try:
        filePointers = downloadStream(streams, downloadFolderPath)
        if len(filePointers) == 2:
            videoFile = VideoFileClip(filePointers[0])
            audioFile = AudioFileClip(filePointers[1])
            videoFile.set_audio(audioFile)
            videoFile.write_videofile(
                os.path.join(downloadFolderPath, video.title.replace('|', '-') + '.mp4')
            )
            # os.remove(filePointers[0])
            # os.remove(filePointers[1])

    except Exception as ex:
        print(ex)
        displayError('Some unexpected error occurred! Maybe loss of internet connection.')


colorama.init()
if len(sys.argv) < 2:
    displayError('Please enter the URL of the video!')
    sys.exit(1)

link = sys.argv[1].strip()
if len(sys.argv) == 3:
    downloadFolderPath = sys.argv[2]
else:
    downloadFolderPath = None

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
session = requests.Session()
downloadYouTubeVideo(videoLink=link, downloadFolderPath=downloadFolderPath)
session.close()
