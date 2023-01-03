import time, datetime, webbrowser, openpyxl, sys
import tkinter as tk
import pandas as pd
from pathlib import Path

# Setting the current date for file reading
x = datetime.datetime.now()
Date = x.strftime("%Y-%m-%d")

# Setting the user's download path and file path
downloads_path = str(Path.home() / "Downloads")
our_file = downloads_path + "\\aha-music-export_" + Date
our_file_csv = Path(our_file + ".csv")
our_file_xlsx = Path(our_file + ".xlsx")

csv_state = False
xlsx_state = False

song_index = 0
max_rows = 0    # I dream of a world where sheet.max_row works after deletion


class MyGUI:
    def __init__(self):
        # This function makes the user interface
        self.root = tk.Tk()

        self.root.geometry('450x200')
        self.root.title('AHA Music Scraper')

        # Top text box
        self.textbox = tk.Text(self.root, height=2, font=('Arial', 8), state='disabled')
        self.textbox.pack(padx=10, pady=10)

        # Check for file button
        self.btn_check_file = tk.Button(self.root, text='Check for file', font=('Arial', 12), command=self.find_file)
        self.btn_check_file.place(x=10, y=50, height=25, width=110)

        # Convert CSV button
        self.btn_convCSV = tk.Button(self.root, text='Convert to xlsx', font=('Arial', 12),
                                     command=self.convert_xlsx, state='disabled')
        self.btn_convCSV.place(x=120, y=50, height=25, width=120)

        # Start iterating through songs button
        self.btn_start = tk.Button(self.root, text='Start', font=('Arial', 12), command=self.read_songs,
                                   state='disabled')
        self.btn_start.place(x=270, y=50, height=25, width=60)

        # Open song URL button
        self.btn_open_song = tk.Button(self.root, text='Open URL', font=('Arial', 12), command=self.open_url,
                                       state='disabled')
        self.btn_open_song.place(x=340, y=50, height=25, width=100)

        # Next song button
        self.btn_next = tk.Button(self.root, text='Next song', font=('Arial', 12), command=self.next_song,
                                  state='disabled')
        self.btn_next.place(x=340, y=80, height=25, width=100)

        # Delete song from file button
        self.btn_rem_song = tk.Button(self.root, text='Delete song from file', font=('Arial', 10), fg='red',
                                      command=self.delete_song, state='disabled')
        self.btn_rem_song.place(x=310, y=170, height=20, width=130)

        self.label_csv = tk.Label(self.root, text='CSV:', font=('Arial', 8))
        self.label_csv.place(x=10, y=100)

        self.label_xlsx = tk.Label(self.root, text='XLSX: ', font=('Arial', 8))
        self.label_xlsx.place(x=10, y=120)

        self.label_csv_result = tk.Label(self.root, text='awaiting check...', font=('Arial', 8), fg='dark orange')
        self.label_csv_result.place(x=50, y=100)

        self.label_xlsx_result = tk.Label(self.root, text='awaiting check...', font=('Arial', 8), fg='dark orange')
        self.label_xlsx_result.place(x=50, y=120)

        self.root.mainloop()

    def write_to_textbox(self, new_text):
        self.textbox.config(state='normal')
        self.textbox.delete(1.0, 'end')
        self.textbox.insert(1.0, new_text)
        self.textbox.config(state='disabled')

    def enable_file_buttons(self):
        self.btn_next.config(state='normal')
        self.btn_open_song.config(state='normal')
        self.btn_rem_song.config(state='normal')

    def disable_file_buttons(self):
        self.btn_next.config(state='disabled')
        self.btn_open_song.config(state='disabled')
        self.btn_rem_song.config(state='disabled')

    def find_file(self):
        # Checks if a current date CSV or xlsx file exists in the user downloads folder.
        global csv_state, xlsx_state, song_index

        self.disable_file_buttons()
        song_index = 0

        if our_file_csv.is_file() and our_file_xlsx.is_file():
            # Both the CSV and xlsx file exist
            csv_state = True
            xlsx_state = True
            self.btn_convCSV.config(state='disabled')
            self.btn_start.config(state='normal')
            self.label_csv_result.config(text='File found', fg='dark green')
            self.label_xlsx_result.config(text='File found', fg='dark green')
            self.write_to_textbox('If you would like to replace the current xlsx file, please delete it or move it from'
                                  ' your \nDownloads folder.')
        elif our_file_csv.is_file():
            # CSV file exists but no xlsx file
            csv_state = True
            self.btn_convCSV.config(state='normal')
            self.label_csv_result.config(text='File found', fg='dark green')
            self.label_xlsx_result.config(text='File not found', fg='red')
        elif our_file_xlsx.is_file():
            # xlsx file exists but no CSV file
            xlsx_state = True
            self.btn_convCSV.config(state='disabled')
            self.btn_start.config(state='normal')
            self.label_xlsx_result.config(text='File found', fg='dark green')
            self.label_csv_result.config(text='File not found', fg='red')
        else:
            # Neither file exists
            print("Neither CSV or xlsx exists, please ensure that it is in your Downloads folder")
            csv_state = False
            xlsx_state = False
            self.btn_convCSV.config(state='disabled')
            self.label_csv_result.config(text='File not found', fg='red')
            self.label_xlsx_result.config(text='File not found', fg='red')
            self.write_to_textbox('Neither CSV or xlsx exists, please ensure that it is in your Downloads folder')

    def convert_xlsx(self):
        conv_file = pd.read_csv(our_file_csv)
        conv_file.to_excel(our_file + ".xlsx", index=None, header=True)
        self.find_file()

    def read_songs(self):
        # Load the Excel workbook into memory
        global song_index, max_rows
        wb = openpyxl.load_workbook(our_file + ".xlsx")
        sheet = wb.active

        if max_rows == 0:
            max_rows = sheet.max_row

        for row in sheet.iter_rows(min_row=1):
            if song_index > 1:
                break
            else:
                artist = row[1].value
                song_name = row[2].value
                url = row[5].value
                print(f'{artist}: {song_name}\n{url}')

            self.write_to_textbox(f'{artist}: {song_name}   {song_index}/{max_rows-1}\n{url}')
            self.enable_file_buttons()
            song_index += 1

        self.btn_start.config(state='disabled')

    def open_url(self):
        wb = openpyxl.load_workbook(our_file + ".xlsx")
        sheet = wb.active
        url = sheet.cell(row=song_index, column=6).value

        if url is not None:
            webbrowser.open(url)
        else:
            self.write_to_textbox('There is no URL for this song')

    def next_song(self):
        global song_index, max_rows
        song_index += 1

        wb = openpyxl.load_workbook(our_file + ".xlsx")
        sheet = wb.active

        if song_index == max_rows+1:
            self.write_to_textbox('No more songs in file')
            wb.save(our_file + ".xlsx")
            wb.close()
            self.disable_file_buttons()
            self.btn_start.config(state='normal')
            song_index = 0
        else:
            row = sheet[song_index]
            artist = row[1].value
            song_name = row[2].value
            url = row[5].value

            print(f'{artist}: {song_name}\n{url}')
            self.write_to_textbox(f'{artist}: {song_name}   {song_index-1}/{max_rows-1}\n{url}')

    def delete_song(self):
        global song_index, max_rows
        wb = openpyxl.load_workbook(our_file + ".xlsx")
        sheet = wb.active

        sheet.delete_rows(song_index)
        wb.save(our_file + ".xlsx")
        wb.close()

        song_index -= 1
        max_rows -= 1
        self.write_to_textbox('Song removed from file')


MyGUI()
