import os
import openpyxl
import pydub
import re
from pytube import YouTube
from pytube.exceptions import RegexMatchError
from moviepy.editor import VideoFileClip
from moviepy.editor import AudioFileClip
from queue import Queue
from threading import Thread


# consts
A = 0
B = 1
C = 2
D = 3
E = 4

download_queue = Queue()

def downloadAudio(yt_url, folder, timestamps):
    print(f"Downloading and trimming {yt_url} to {folder}")

    try:
        url = YouTube(yt_url)
        stream = url.streams.filter(file_extension='mp4').first()
        folder_path = os.path.join(download_dir, folder)

        os.makedirs(folder_path, exist_ok=True)
        file_path = stream.download(output_path=folder_path, filename=url.title)

        # convert audio extention to mp3
        # video_duration = url.length
        # output_path = os.path.join(folder_path, f"{url.title}.mp3")
        # clip = VideoFileClip(file_path)
        # audio_clip = clip.audio
        # audio_clip.write_audiofile(output_path)

        if timestamps:
            trimmed_output_path = os.path.join(folder_path, f'{url.title}_trim.mp3')
            trimAudio(file_path, trimmed_output_path, timestamps)
            
            # remove not trimmed audio file
            os.remove(file_path)
        else:
            trimmed_output_path = file_path

        print("done!")

    except RegexMatchError as error:
        print(f"Download error for {url}: {error}")


def timestampToSeconds(timestamp):
    # convert timestamp in str to sec int
    match = re.match(r'(\d+):(\d+)', timestamp)

    if match:
        minutes, seconds = map(int, match.groups())
        return (minutes * 60 + seconds)
    return 0

def trimAudio(file_path, output_path, timestamps):
    # get audio file
    audio = AudioFileClip(file_path)

    # process timestamps string
    start, end = re.findall(r'(\d+:\d+)', timestamps)

    # convert timestamps str to sec
    start_sec = timestampToSeconds(start)
    
    # Check if the end timestamp is greater than the video duration
    video_duration = audio.duration
    end_sec = min(timestampToSeconds(end), video_duration)

    # trim audio from the end first
    trimmed_audio = audio.subclip(0, end_sec)

    # trim audio from the start
    trimmed_audio = trimmed_audio.subclip(start_sec, None)

    # save trimmed file
    trimmed_audio.write_audiofile(output_path)


def processQueue():
    while True:
        task = download_queue.get()
    
        if task is None:
            break
        
        # unpack task
        yt_url, folder = task
        #debug
        print(f"Processing task: {yt_url} in {folder}")
        # download audio
        downloadAudio(yt_url, folder, timestamps)
        # task done
        download_queue.task_done()


# open excel file
workbook = openpyxl.load_workbook("/home/usuario/Documentos/test.xlsx") # dir
#workbook = openpyxl.load_workbook("/home/adzk/youtube-audio-download-trim/test.xlsx")
worksheet = workbook.active
download_dir = "/home/usuario/Documentos/musicas-formatura"
#download_dir = "/home/adzk/Documents/formatura"

#queue
worker_thread = Thread(target=processQueue)
worker_thread.start()

# reading xlsx data
for row in worksheet.iter_rows(min_row=2, values_only=True):
    folder = row[C]
    yt_url = row[D]
    timestamps = row[E]
    download_queue.put((yt_url, os.path.join(download_dir, folder)))

# end of the queue
download_queue.put(None)

# wait for the worker thread to finish
worker_thread.join()
