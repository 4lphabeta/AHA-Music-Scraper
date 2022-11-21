import time, datetime, webbrowser, openpyxl, sys
import pandas as pd
from pathlib import Path

#Setting the current date for file reading
x = datetime.datetime.now()
Date = x.strftime("%Y-%m-%d")

#Setting the user's download path and file path
downloads_path = str(Path.home() / "Downloads")
our_file = downloads_path + "\\aha-music-export_" + Date
our_file_csv = Path(our_file + ".csv")
our_file_xlsx = Path(our_file + ".xlsx")

#Check if csv file exists
if our_file_csv.is_file():
    #Convert the csv to xlsx
    print("Converting csv file to xlsx")
    conv_file = pd.read_csv (our_file_csv)
    conv_file.to_excel (our_file + ".xlsx", index = None, header=True)
    print("File converted successfully")
elif our_file_xlsx.is_file():
    #An already converted file exists but there is no csv, continue as normal
    print("xlsx file found despite finding no csv, continuing as normal")
else:
    print("csv file does not exist, please ensure that it is in your Downloads folder")
    sys.exit()

#Loading the Excel workbook
wb = openpyxl.load_workbook(our_file + ".xlsx")
ws = wb.active

print("Workbook loaded")
print("Iterating through urls")
  
#Iterate through rows and open each url
for row in ws.iter_rows(min_row=2):
    url = row[5].value
    if url is not None:
        webbrowser.open(url)
        time.sleep(0.5)     #Delay slightly to hopefully prevent crashes

print("All urls opened")
print("Script finished")
