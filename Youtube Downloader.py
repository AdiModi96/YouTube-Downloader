import sys
import colorama
from colorama import Fore, Back, Style
from pytube import YouTube
from prettytable import PrettyTable
from winreg import *
from moviepy.editor import *


# Example Link: https://youtu.be/LXb3EKWsInQ


# The Progressive Streams have audio and video together
# The DASH Streams have separate audio and video, need to merge them post the download

def display_information(video):
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


def display_streams(video):
    print(Fore.YELLOW + 'Video Streams: ' + Style.RESET_ALL)
    print_table = PrettyTable([Fore.YELLOW + 'File Type' + Style.RESET_ALL,
                              Fore.YELLOW + 'ID' + Style.RESET_ALL,
                              Fore.YELLOW + 'Resolution' + Style.RESET_ALL,
                              Fore.YELLOW + 'FPS' + Style.RESET_ALL,
                              Fore.YELLOW + 'Audio Bitrate' + Style.RESET_ALL,
                              Fore.YELLOW + 'Extension' + Style.RESET_ALL])

    for stream in video.streams:
        file_type = stream.type
        file_extension = stream.subtype

        if file_type == 'video':
            itag = stream.itag
            resolution = stream.resolution
            fps = stream.fps
            if resolution == None or fps == None:
                continue
            audio_bit_rate = stream.abr
            if audio_bit_rate == None:
                file_type = 'Video Only'
            else:
                file_type = 'Audio+Video'

            print_table.add_row([Fore.CYAN + str(file_type) + Style.RESET_ALL,
                                Fore.CYAN + str(itag) + Style.RESET_ALL,
                                Fore.CYAN + str(resolution) + Style.RESET_ALL,
                                Fore.CYAN + str(fps) + Style.RESET_ALL,
                                Fore.CYAN + str(audio_bit_rate) + Style.RESET_ALL,
                                Fore.CYAN + str(file_extension) + Style.RESET_ALL])

        elif file_type == 'audio':
            itag = stream.itag
            resolution = stream.resolution
            fps = stream.fps
            audio_bit_rate = stream.abr
            if audio_bit_rate == None:
                continue
            file_type = 'Audio'
            print_table.add_row([Fore.CYAN + str(file_type) + Style.RESET_ALL,
                                Fore.CYAN + str(itag) + Style.RESET_ALL,
                                Fore.CYAN + str(resolution) + Style.RESET_ALL,
                                Fore.CYAN + str(fps) + Style.RESET_ALL,
                                Fore.CYAN + str(audio_bit_rate) + Style.RESET_ALL,
                                Fore.CYAN + str(file_extension) + Style.RESET_ALL])

    print(print_table)


def display_error(errorMessage):
    print(
        Fore.RED + 'Error: ' + Fore.LIGHTRED_EX + errorMessage + Style.RESET_ALL)


def download_stream(streams, download_folder_path):
    if len(streams) == 1:
        file_name = str(streams[0].title).replace('|', '-')
        file_path = os.path.join(download_folder_path, file_name + '.' + streams[0].subtype)

        print(Fore.YELLOW + 'Downloading Stream with ID: {}'.format(streams[0].itag) + Style.RESET_ALL)
        streams[0].download(output_path=download_folder_path, filename=file_name)
        print(Fore.GREEN + 'Downloaded!' + Style.RESET_ALL)
    else:
        file_name = str(streams[0].title).replace('|', '-')
        file_path = os.path.join(download_folder_path, file_name)

        print(Fore.YELLOW + 'Downloading Stream with ID: {}'.format(streams[0].itag) + Style.RESET_ALL)
        video_file_name = str(streams[0].itag)
        video_file_path = os.path.join(download_folder_path, video_file_name + '.' + streams[0].subtype)
        streams[0].download(output_path=download_folder_path, filename=video_file_name)
        print(Fore.GREEN + 'Downloaded!' + Style.RESET_ALL)

        print(Fore.YELLOW + 'Downloading Stream with ID: {}'.format(streams[1].itag) + Style.RESET_ALL)
        audio_file_name = str(streams[1].itag)
        audio_file_path = os.path.join(download_folder_path, audio_file_name + '.' + streams[0].subtype)
        streams[1].download(output_path=download_folder_path, filename=audio_file_name)
        print(Fore.GREEN + 'Downloaded!' + Style.RESET_ALL)

        print(Fore.YELLOW + 'Merging' + Style.RESET_ALL)
        video_file = VideoFileClip(video_file_path)
        audio_file = AudioFileClip(audio_file_path)
        video_file.set_audio(audio_file)
        video_file.write_videofile(file_path + '.' + streams[0].subtype)
        print(Fore.GREEN + 'Completed!' + Style.RESET_ALL)

        os.remove(video_file_path)
        os.remove(audio_file_path)

    print(Fore.GREEN + 'File Available at:' + Style.RESET_ALL, file_path)

def download_youtube_video(video, download_folder_path):
    display_information(video)
    display_streams(video)

    itags = []
    streams = []
    print('Instructions:' + '\n' +
          '● Enter 1 ID value to download one file.' + '\n' +
          '● Enter 2 ID values, separated by space Order: video audio, to download and merge.')

    itags_string = input('Enter the ID of the stream(s) you want to download: ').split(' ')
    for itag in itags_string:
        if len(itag) != 0:
            itags.append(itag)

    if len(itags) >= 1 and len(itags) <= 2:
        for itag in itags:
            stream = video.streams.get_by_itag(itag)
            if stream == None:
                print(Fore.LIGHTRED_EX + 'Incorrect ID: {}, Skipping'.format(itag) + Style.RESET_ALL)
            else:
                streams.append(stream)
    else:
        display_error('Please enter only 1 or 2 ID values')

    if download_folder_path == '':
        try:
            with OpenKey(HKEY_CURRENT_USER,
                         'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
                download_folder_path = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        except Exception:
            download_folder_path = '.'

    elif not os.path.exists(download_folder_path):
        os.makedirs(download_folder_path)

    try:
        download_stream(streams, download_folder_path)
    except Exception:
        display_error('Some unexpected error occurred! Maybe loss of internet connection.')


colorama.init()

# Collecting Video Link
video_link = input('Enter Link of the Video: ').strip()

# Creating YouTube video link object
try:
    video = YouTube(video_link)
except Exception:
    # Throwing error if unable to create object
    display_error(
        'One of the following problems could have occurred:' + '\n' +
        '● Couldn\'t connect to the Internet' + '\n' +
        '● The URL is incorrect' + '\n' +
        '● The video doesn\'t have download permission(s)'
    )
    sys.exit(1)
download_folder_path = input('Enter Folder Path to Download the Video: ').strip()

download_youtube_video(video=video, download_folder_path=download_folder_path)
