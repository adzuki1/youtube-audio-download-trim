from pytube import YouTube
from pytube.exceptions import RegexMatchError
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

    try:
        url = YouTube(yt_url)
        stream = url.streams.filter(only_audio=True).first()
        folder_path = os.path.join(download_dir, folder)

        os.makedirs(folder_path, exist_ok=True)
        stream.download(output_path=folder_path, filename=url.title)
        
        print("Done!")
    except RegexMatchError as error:
        print(f"Download error for {url}: {error}")

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

