import youtube_dl

def downloadAudio(url):
    ydl_opts = {'format': 'bestaudio/best',
                'postprocessors': [
                    {'key': 'FFmpegExtractAudio',
                     'preferredcodec': 'mp3',
                     'preferredquality': '192',}],}
