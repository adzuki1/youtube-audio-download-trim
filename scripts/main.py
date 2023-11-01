import youtube_dl
import os
import openpyxl
from moviepy.video.io.ffmpeg_tools import ffmpeg_extract_subclip

A = 0
B = 1
C = 2
D = 3
E = 4

# open excel file
workbook = openpyxl.load_workbook("file.xlsx")
worksheet = workbook.active

# coletando os dados das colunas
for row in worksheet.iter_rows(min_row=2, values_only=True):
    folder = row[A]
    yt_url = row[D]
    timestamps = row[E]

def downloadAudio(yt_url):
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
    }
    with youtube_dl.YoutubeDL(ytdl_settings) as ytdl:
        info_dict = ytdl.extract_info(yt_url, download=True)