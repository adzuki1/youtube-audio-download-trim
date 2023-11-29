import os
import openpyxl
import re
from pytube import YouTube
from pytube.exceptions import RegexMatchError
from moviepy.editor import AudioFileClip
from queue import Queue
from threading import Thread

# Globals:
# consts
A = 0
B = 1
C = 2
D = 3
E = 4

download_queue = Queue()

def downloadAudio(yt_url, download_dir, new_folder, timestamps):
    try:
        url = YouTube(yt_url)
        stream = url.streams.filter(file_extension='mp4').first()
        new_folder_path = os.path.join(download_dir, new_folder)

        os.makedirs(new_folder_path, exist_ok=True)
        file_path = stream.download(output_path=new_folder_path, filename=url.title)

        if timestamps:
            trimmed_output_path = os.path.join(new_folder_path, f'{url.title}_trim.mp3')
            trimAudio(file_path, trimmed_output_path, timestamps)
            
            # remove not trimmed audio file
            os.remove(file_path)
        else:
            trimmed_output_path = file_path

        print("done!\n")

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

    print(f"\nOriginal timestamps: {timestamps}\n")

    # process timestamps string
    start, end = re.findall(r'\d+:\d+', timestamps)

    # convert timestamps str to sec
    start_sec = timestampToSeconds(start)
    print(f"Start time {start} in seconds: {start_sec}\n")
    
    # Check if the end timestamp is greater than the video duration
    video_duration = audio.duration
    end_sec = min(timestampToSeconds(end), video_duration)
    print(f"End time {end} in seconds: {end_sec}\n")

    # trim audio from the end first
    trimmed_audio = audio.subclip(0, end_sec)
    # trim audio from the start
    trimmed_audio = trimmed_audio.subclip(start_sec, None)
    # save trimmed file
    trimmed_audio.write_audiofile(output_path)

def processQueue():
    counter = 0

    while True:
        task = download_queue.get()
        counter += 1
    
        if task is None:
            break

        # unpack task
        yt_url, download_dir, new_folder, timestamps = task

        # skip task if YouTube URL is empty
        if not yt_url:
            print(f"SKIPPING TASK {counter}. EMPTY URL.")
            download_queue.task_done()
            continue

        #debug
        print(f"\nProcessing task {counter}: {yt_url} in {new_folder}")
        # download audio
        downloadAudio(yt_url, download_dir, new_folder, timestamps)
        # task done
        download_queue.task_done()

def main():
    download_dir1 = "/home/usuario/Documentos/musicas-formatura/3001"
    download_dir2 = "/home/usuario/Documentos/musicas-formatura/3002"
    download_dir3 = "/home/usuario/Documentos/musicas-formatura/3003"

    class_dirs = [download_dir1, download_dir2, download_dir3]
    start_rows = [2, 28, 59]
    end_rows = [27, 58, 86]

    # open excel file
    workbook = openpyxl.load_workbook("/home/usuario/Documentos/formatura-downloads-sheet.xlsx") # dir
    worksheet = workbook.active

    # queue
    worker_threads = []

    for i in range(len(class_dirs)):
        start_row, end_row = start_rows[i], end_rows[i]

        # create a new thread for each range
        worker_thread = Thread(target=processQueue)
        worker_thread.start()
        worker_threads.append(worker_thread)

        # reading xlsx data
        for row in worksheet.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
            new_folder = row[B]
            yt_url = row[D]
            timestamps = row[E]

            # skip task if YouTube URL is empty
            if not yt_url:
                print(f"SKIPPING TASK. EMPTY URL.")
                continue

            # add task to the corresponding queue
            download_queue.put((yt_url, class_dirs[i], new_folder, timestamps))

        # end of the queue for this directory
        download_queue.put(None)

    # wait for all worker threads to finish
    for worker_thread in worker_threads:
        worker_thread.join()

if __name__ == "__main__":
    main()
