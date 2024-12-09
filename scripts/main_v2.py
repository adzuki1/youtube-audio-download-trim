import os
import openpyxl
import re
from moviepy.audio.io.AudioFileClip import AudioFileClip
from queue import Queue
from threading import Thread
import yt_dlp

# Globals:
# consts
A = 0
B = 1
C = 2
D = 3
E = 4

download_queue = Queue()
download_dir = "/musicas" # "/your/download/directory"


def downloadAudio(yt_url, new_folder, timestamps):
    try:
        # Ensure folder path exists
        new_folder_path = os.path.join(download_dir, new_folder)
        os.makedirs(new_folder_path, exist_ok=True)

        # yt-dlp options for downloading audio
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': os.path.join(new_folder_path, '%(title)s.%(ext)s'),
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
        }

        # Download the audio
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info_dict = ydl.extract_info(yt_url, download=True)
            mp3_file_path = ydl.prepare_filename(info_dict).replace('.webm', '.mp3').replace('.m4a', '.mp3')

        # Trim audio if timestamps are provided
        if timestamps:
            trimmed_output_path = os.path.join(new_folder_path, f"{info_dict['title']}_trim.mp3")
            trimAudio(mp3_file_path, trimmed_output_path, timestamps)

            # Remove the untrimmed MP3 file
            os.remove(mp3_file_path)
        else:
            trimmed_output_path = mp3_file_path

        print(f"Download completed: {trimmed_output_path}\n")

    except Exception as error:
        print(f"Download error for {yt_url}: {error}")


def timestampToSeconds(timestamp):
    # Convert timestamp in str to sec int
    match = re.match(r'(\d+):(\d+)', timestamp)

    if match:
        minutes, seconds = map(int, match.groups())
        return minutes * 60 + seconds

    return 0


def trimAudio(file_path, output_path, timestamps):
    # Open the audio file
    audio = AudioFileClip(file_path)

    print(f"\nOriginal timestamps: {timestamps}\n")

    # Process timestamps string
    start, end = re.findall(r'\d+:\d+', timestamps)

    # Convert timestamps str to sec
    start_sec = timestampToSeconds(start)
    print(f"Start time {start} in seconds: {start_sec}\n")

    # Check if the end timestamp is greater than the audio duration
    audio_duration = audio.duration
    end_sec = min(timestampToSeconds(end), audio_duration)
    print(f"End time {end} in seconds: {end_sec}\n")

    # Trim the audio
    trimmed_audio = audio.subclip(start_sec, end_sec)

    # Save trimmed MP3 file
    trimmed_audio.write_audiofile(output_path, codec="libmp3lame")
    trimmed_audio.close()
    audio.close()


def processQueue():
    counter = 0

    while True:
        task = download_queue.get()
        counter += 1

        if task is None:
            break

        # Unpack task
        yt_url, new_folder, timestamps = task
        # Debug
        print(f"\nProcessing task {counter}: {yt_url} in {new_folder}")
        # Download audio
        downloadAudio(yt_url, new_folder, timestamps)
        # Task done
        download_queue.task_done()


def main():

    # Open Excel file
    workbook = openpyxl.load_workbook("url-input4.xlsx")  # Update with your Excel file's path
    worksheet = workbook.active

    # Queue worker thread
    worker_thread = Thread(target=processQueue)
    worker_thread.start()

    # Read XLSX data
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        new_folder = row[C]
        yt_url = row[D]
        timestamps = row[E]
        download_queue.put((yt_url, new_folder, timestamps))

    # End of the queue
    download_queue.put(None)

    # Wait for the worker thread to finish
    worker_thread.join()


if __name__ == "__main__":
    main()

