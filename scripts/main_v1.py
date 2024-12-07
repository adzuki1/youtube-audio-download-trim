import os
import openpyxl
import re
import yt_dlp
from moviepy.audio.io.AudioFileClip import AudioFileClip

# Globals:
A = 0
B = 1
C = 2
D = 3
E = 4


def downloadAudio(yt_url, download_dir, new_folder, timestamps):
    try:
        new_folder_path = os.path.join(download_dir, str(new_folder))  # Ensure new_folder is a string
        os.makedirs(new_folder_path, exist_ok=True)

        # Configure yt-dlp options
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': os.path.join(new_folder_path, '%(title)s.%(ext)s'),
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
        }

        # Download the audio using yt-dlp
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info_dict = ydl.extract_info(yt_url, download=True)
            mp3_file_path = ydl.prepare_filename(info_dict).replace('.webm', '.mp3').replace('.m4a', '.mp3')

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
    # Get audio file
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

    # Trim audio using set_start() and set_end()
    trimmed_audio = audio.subclip(start_sec, end_sec)

    # Save trimmed MP3 file
    trimmed_audio.write_audiofile(output_path, codec="libmp3lame")
    trimmed_audio.close()
    audio.close()


def processTasks(class_dir, worksheet, start_row, end_row):
    for row in worksheet.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        new_folder = row[B]
        yt_url = row[D]
        timestamps = row[E]

        # Skip task if YouTube URL is empty
        if not yt_url:
            print("SKIPPING TASK. EMPTY URL.")
            continue

        # Debug
        print(f"\nProcessing task:

{yt_url} in {new_folder}\n")

        # Download audio with the given timestamps
        downloadAudio(yt_url, class_dir, new_folder, timestamps)

        # Update status column in the worksheet
        status_column = "Status"
        status_idx = None
        for idx, cell in enumerate(worksheet[1]):
            if cell.value == status_column:
                status_idx = idx + 1
                break

        if status_idx:
            worksheet.cell(row=row[0], column=status_idx).value = "Completed"

    # Save updated worksheet
    workbook.save(os.path.join(class_dir, "tasks.xlsx"))


# Main function
def main():
    class_dir = "musicas/3001"
    excel_file = os.path.join(class_dir, "tasks.xlsx")

    # Load the Excel workbook and worksheet
    workbook = openpyxl.load_workbook(excel_file)
    worksheet = workbook.active

    # Process tasks from row 2 to the last row
    start_row = 2
    end_row = worksheet.max_row

    processTasks(class_dir, worksheet, start_row, end_row)

    print("\nAll tasks processed successfully.")


if __name__ == "__main__":
    main()

