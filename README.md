
PROJECT STILL IN PRODUCTION
# PT-BR
 
 Soon

# EN
#  VERSION 1

This repository contains `main_v1.py`, a Python script for downloading audio from YouTube videos, trimming it based on timestamps, and organizing the output into specified folders. This script relies on an Excel sheet for task configuration and provides basic functionality for audio processing.

## Features
- Download audio files from YouTube videos in MP3 format.
- Trim audio files based on user-provided timestamps.
- Organize downloads into folders specified in an Excel sheet.
- Automated batch processing of tasks defined in the Excel file.

## Dependencies
The following libraries are required to run the script:
- `os` (built-in)
- `re` (built-in)
- `openpyxl`: Excel file handling
- `yt_dlp`: YouTube video/audio downloading
- `moviepy`: Audio processing and trimming

To install the required external libraries, run:
```bash
pip install openpyxl yt-dlp moviepy
```

or
```bash
pip install -r requirements_v1_v2.txt
```

## File Structure
- **`main_v1.py`**: The main script for executing the tasks.
- **`test.xlsx`**: An Excel file used to configure download tasks.
- **`DOWNLOADS`**: Default directory where output files are saved.

## Excel Sheet Configuration
The script processes tasks from an Excel file. The columns in the sheet should be structured as follows:

| Column | Description                    |
| ------ | ------------------------------ |
| A      | Task ID (optional)             |
| B      | Task Description (optional)    |
| C      | Folder Name                    |
| D      | YouTube URL                    |
| E      | Timestamps (e.g., `1:00-2:30`) |

### Example:
| Task ID | Task Description | Folder Name | YouTube URL            | Timestamps |
|---------|------------------|-------------|------------------------|------------|
| 1       | Example Task     | Folder1     | https://youtu.be/xyz123 | 0:30-1:00  |

## Usage
1. Prepare an Excel file (`test.xlsx`) with tasks following the structure described above.
2. Place the script (`main_v1.py`) and the Excel file in the same directory.
3. Run the script:
   ```bash
   python3 main_v1.py
   ```
4. Processed audio files will be saved in the specified folders under the `DOWNLOADS` directory.

## Functions Overview

- **`downloadAudio(yt_url, download_dir, new_folder, timestamps)`**: Downloads and optionally trims audio from a YouTube video.

- **`timestampToSeconds(timestamp)`**: Converts a timestamp string to seconds.

- **`trimAudio(file_path, output_path, timestamps)`**: Trims an audio file based on provided timestamps.

- **`processTasks(dl_directory, worksheet, start_row, end_row)`**: Processes a range of tasks defined in the Excel file.

- **`main()`**: Entry point of the script; handles the overall workflow.

## Notes
- Ensure `FFmpeg` is installed and accessible in your system's PATH for audio processing.
- Tasks with empty YouTube URLs will be skipped.
- If no timestamps are provided, the entire audio file will be downloaded without trimming.

---

