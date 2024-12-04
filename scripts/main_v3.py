# dependencies
import os
import pandas as pd
from pydub import AudioSegment
from youtube_dl import YoutubeDL

# Config
BASE_DIR = os.path.expanduser("~/path/to/main/folder")
EXCEL_FILE = "~/path/to/excel/file.xlsx"


# Create directories from Excel file, to save each file separately
def createDirectory(folder_name):


# Download as audio files from Youtube URLs
def downloadAudio(url, output_path):


# Parse timestamps and convert it to seconds, to trim files
def parseTimestamps(timestamp):


# Uses processed timestamps to trim audio files in an interval 
def trimAudio(input_path, output_path=None, timestamp=None):


def processTask(file_data, output_path):


# Main function, to open file and pass arguments 
def main():

	df = pd.read_excel(EXCEL_FILE)
	
	for _, row in df.iterrows():
		processTask() # arguments
	

if __name__ == "__main__":
	main()
	
	
	
	
