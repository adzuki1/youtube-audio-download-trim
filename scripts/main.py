import youtube_dl
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

# open excel file
workbook = openpyxl.load_workbook("/home/usuario/Documentos/test.xlsx") # dir
worksheet = workbook.active
download_dir = "/home/usuario/Documentos/musicas-formatura"

download_queue = Queue()

def downloadAudio(yt_url, folder):
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
    }
    with youtube_dl.YoutubeDL(ytdl_settings) as ytdl:
        info_dict = ytdl.extract_info(yt_url, download=True)

def processQueue():
    while True:
        task = download_queue.get()
    
        if task is None:
            break
        
        # unpack task
        yt_url, folder = task
        # download audio
        downloadAudio(yt_url, folder)
        # task done
        download_queue.task_done()


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