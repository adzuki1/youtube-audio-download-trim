import youtube_dl
from pytube import Youtube
import os
import openpyxl
from moviepy.video.io.ffmpeg_tools import ffmpeg_extract_subclip
from queue import Queue
from threading import Thread

# consts
A = 0
B = 1
C = 2
D = 3
E = 4

download_queue = Queue()

def downloadAudio(yt_url, folder):
    print(f"Downloading {yt_url} to {folder}")

    ytdl_settings = {
    "format": "bestaudio/best",
    "postprocessors": [
        {
            "key": "FFmpegExtractAudio",
            "preferredcodec": "mp3",
            "preferredquality": "192",
        }
    ],
    "outtmpl": f"{folder}/%(title)s.%(ext)s",
    "verbose": True,
    "ignoreErrors": True,
    }
    with youtube_dl.YoutubeDL(ytdl_settings) as ytdl:
        try:
            info_dict = ytdl.extract_info(yt_url, download=True)
        except youtube_dl.utils.DownloadError as e:
            print(f"Download error for {yt_url}: {e}")

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
        downloadAudio(yt_url, folder)
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
    # timestamps = row[E]
    download_queue.put((yt_url, os.path.join(download_dir, folder)))

# end of the queue
download_queue.put(None)

# wait for the worker thread to finish
worker_thread.join()

workbook.save("test2.xlsx")
