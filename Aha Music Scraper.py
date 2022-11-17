import time, datetime, webbrowser, openpyxl
from pathlib import Path

#Setting the user's download path
downloads_path = str(Path.home() / "Downloads")

#Setting the current date for file reading
x = datetime.datetime.now()
Date = x.strftime("%Y-%m-%d")

#Load excel with its path
data_folder = downloads_path + "\\aha-music-export_" + Date + ".xlsx"
wb = openpyxl.load_workbook(data_folder)

ws = wb.active

print("Workbook loaded")
print("Iterating through urls")
  
#Iterate through rows and open each url
for row in ws.iter_rows(min_row=2):
    url = row[5].value
    webbrowser.open(url)

print("All urls opened")
print("Script finished")

time.sleep(3)   #Sleep for 3 seconds