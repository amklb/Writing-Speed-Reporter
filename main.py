from win32gui import GetWindowText, GetForegroundWindow # Processes, paths, threading
from win32process import GetWindowThreadProcessId
import psutil
from pathlib import Path
import threading
from os import listdir
from os.path import isfile, join

from time import sleep # Time 
from datetime import datetime

import numpy as np # Data handling
import pandas as pd

from pynput import keyboard #Keyboard input

import py_cui # UI

import seaborn as sns #Plots
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from reportlab.pdfgen import canvas # PDFs
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont



class WritingSpeedApp():
    def __init__(self, master : py_cui.PyCUI):
        # Iniciate global variables
        self.events = []
        self.per_minute_events = []
        self.writing_speed_per_minute = []
        self.REFRESH_TIME = 1 * 60
        self.max_speed = 400
        self.min_speed = 70
        self.running_status = False
        self.status_text = "OFF"
        self.files = []
        self.selected_dates = []
        self.include_all_keys = False
        self.selected_tag = np.nan
        self.tags = []
        try:
            with open(r".\tags\tags.txt", "r") as f:
                for line in f:
                    self.tags.append(line)
        except:
            pass

        # Set up CUI
        self.master = master
        self.title_label = self.master.add_label(title="CPM Reporter", row=0, column=1)
        self.start_button = self.master.add_button("START", 1, 1, command=self.start)
        self.stop_button = self.master.add_button("STOP", 1, 2, command=self.stop)
        self.saved_button = self.master.add_button("SAVED", 2, 1, command=self.saved_popup)
        self.options_button = self.master.add_button("OPTIONS", 2, 2, command=self.open_options)
        self.status_label = self.master.add_label(title="OFF", row=0, column=2).set_color(py_cui.RED_ON_BLACK)
        self.tag_1 = self.master.add_label(title="Current tag:", row=3, column=1)
        self.tag_2 = self.master.add_label(title="NONE", row=3, column=2)

        # Set up variables for threads
        self.listener = None
        self.listener_thread = None
        self.aggregate_thread = None
        self.aggregate_running = False
        
        # Set up folders 
        Path(r".\saved").mkdir(exist_ok=True)
        Path(r".\graphs").mkdir(exist_ok=True)
        Path(r".\pdf").mkdir(exist_ok=True)
        Path(r".\tags").mkdir(exist_ok=True)

        #Set up aesthetics
        sns.set_style("darkgrid")
        sns.set_palette("flare")

    def get_process_name(self):
        #Get name of the current process
        try:
            window_process_id = GetWindowThreadProcessId(GetForegroundWindow())
            process_name = psutil.Process(window_process_id[-1]).name()
            return process_name
        except:
            return np.nan
        
    def on_press(self, key):
        
        # Put keyboard event into temporary df
        time = datetime.now()
        process = self.get_process_name()
        if self.include_all_keys or (type(key) == keyboard._win32.KeyCode) or (key == keyboard.Key.space):
            try:
                event_dict= {
                    "hour" : time.hour,
                    "minute" : time.minute,
                    "process" : process,
                }
                self.events.append(event_dict)
            except AttributeError:
                pass

    def aggregate_events(self):
        # Aggregate data from df list to get cpm; move to main df
        try:
            processed_events = self.events.copy()
            events_df = pd.DataFrame(processed_events)
            self.events.clear()
            events_df = events_df.groupby(["hour", "minute"]).agg(
                strokes_per_minute =("process", "count"), 
                process = ("process", lambda x: x.mode()[0]),
            )
            events_df["tag"] = self.selected_tag
            events_array = events_df.to_dict('records') 
            for event in events_array:
                event["timestamp"] = datetime.now()
                self.per_minute_events.append(event)
        except KeyError:
            pass
        

        
        # Run recursively
        sleep(self.REFRESH_TIME)
        if self.aggregate_running:
            self.aggregate_events()
        else:
            return 

    def generate_report(self, df, mode = 0):
        report_df = df
        timestamp = datetime.now()
        third_graph = False

        # Clean data and get info to be printed
        clean_df = report_df.query(f"strokes_per_minute >= {self.min_speed} and strokes_per_minute <= {self.max_speed}")
        peak_speed = np.max(clean_df["strokes_per_minute"])
        average_speed = np.mean(clean_df["strokes_per_minute"])
        total_characters = np.sum(report_df["strokes_per_minute"])

        # Generate graph with processes
        speed_by_process_data = clean_df.groupby(["process"]).agg(
        speed = ("strokes_per_minute", "mean")).reset_index()
        fig_process = sns.barplot(data=speed_by_process_data,
                                x="process",
                                y="speed",
                                palette="flare",
                                hue="process",
                                legend=False)
        fig_process.set_title("CPM by process")
        fig_process.set_ylabel("Characters Per Minute")
        fig_process.set_xlabel("Used Process")
        fig1 = fig_process.get_figure()
        fig1.savefig(r".\graphs\barplot.png")
        fig1.clf()
       

        # Generate graph with timeline
        
        if mode == 0:
            speed_by_time_data = report_df
            speed_by_time_data["time"] = report_df["timestamp"].dt.strftime('%H:%M')
            
        elif mode == 1:
            speed_by_time_data = clean_df
            speed_by_time_data["time"] = clean_df["timestamp"].dt.date

        fig_time = sns.lineplot(data=speed_by_time_data,
                                x="time",
                                y="strokes_per_minute",
                                errorbar=None,
                                )
        if mode == 1:
            fig_time.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        fig_time.tick_params(axis='x', rotation=90, labelsize=6)
        fig_time.set_title("CPM by time")
        fig_time.set_ylabel("Characters per Minute")
        fig_time.set_xlabel("Timeline")
        fig2 = fig_time.get_figure()
        fig2.savefig(r".\graphs\lineplot.png")
        fig2.clf()
    
        # Additional graph for tags on multi-session report only!
        if mode == 1:
            unique_tags = pd.unique(clean_df["tag"]) 
            if len(unique_tags) == 1 and unique_tags[0] == np.nan:
                pass
            else:
                third_graph = True
                speed_by_tag_data = clean_df.groupby(["tag"]).agg(
                speed = ("strokes_per_minute", "mean")).reset_index()
                fig_tag = sns.barplot(data=speed_by_tag_data,
                                        x="tag",
                                        y="speed",
                                        palette="flare",
                                        hue="tag",
                                        legend=False)
                fig_tag.set_title("CPM by tag")
                fig_tag.set_ylabel("Characters Per Minute")
                fig_tag.set_xlabel("Used Tag")
                fig3 = fig_process.get_figure()
                fig3.savefig(r".\graphs\barplot_2.png")
                fig3.clf()



        # Create Reportlab PDF
        try:
            filename = f".\\pdf\\report{timestamp.date()}_{timestamp.hour}-{timestamp.minute}.pdf"
            doc =  canvas.Canvas(filename, pagesize=letter)
            doc.setLineWidth(.3)
            pdfmetrics.registerFont(TTFont("DejaVu-Sans", r".\font\DejaVuSans.ttf")) # Load font
            doc.setFont("DejaVu-Sans", 12) 
            doc.drawString(45,750,f"Date: {timestamp.date()}, {timestamp.hour}:{timestamp.minute}") #Put text
            doc.drawString(45, 735, f"Total characters: {total_characters} characters")
            doc.drawString(45, 720, f"Average speed: {average_speed: .2f} cpm")
            doc.drawString(45, 705, f"Peak Speed: {peak_speed} cpm")
            if third_graph == False:
                doc.drawImage(r".\graphs\barplot.png", 150, 400, 360, 280) # Put images
            else:
                doc.drawImage(r".\graphs\barplot.png", 10, 430, 300, 220) 
                doc.drawImage(r".\graphs\barplot_2.png", 300, 430, 300, 220) 
                
            doc.drawImage(r".\graphs\lineplot.png", 150, 100, 390, 280)
            doc.save()

            self.master.show_message_popup("Report Generated!", "Report saved in the PDF folder!")

        except:
            self.master.show_error_popup("Error!", "There was a problem generating your report:(")
        
    def save_record(self):
        # Save data into csv and return df
        report_df = pd.DataFrame(self.per_minute_events)
        if report_df.empty == True:
            self.master.show_warning_popup("Warning!", "Not enough data to generate report!")
        else:
            date = datetime.now()
            path = f".\\saved\\{date.year}-{date.month}-{date.day}_{date.hour}-{date.minute}.csv"
            report_df.to_csv(path_or_buf=path)
            self.generate_report(report_df)


    def start(self):
        # Start listening for keyboard
        self.running_status = True
        self.status_label = self.master.add_label(title="Running...", row=0, column=2).set_color(py_cui.GREEN_ON_BLACK)


        if self.listener and self.listener.running: #if thread already started
            self.master.show_warning_popup("Error", "Recording is already running!")
        
        # Start threads
        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener_thread = threading.Thread(target=self.listener.start, daemon=True)
        self.aggregate_thread = threading.Thread(target=self.aggregate_events, daemon=True)
        self.listener_thread.start()
        self.aggregate_thread.start()
        self.aggregate_running = True
        

    def stop(self):
        # Stop listening for keyboard and call to generate report
        
        if self.running_status == True:
            self.listener.stop()
            self.save_record()
            self.listener = None
            self.listener_thread = None
            self.aggregate_running = False
            self.aggregate_thread = None
            self.running_status = False
            self.status_label = self.master.add_label(title="OFF", row=0, column=2).set_color(py_cui.RED_ON_BLACK)
        else:
            self.master.show_error_popup("Error!", "Tracking is not running")
        
    def options(self, paramter):
        # Set min/max speed that will be considered in aggregating report
        if paramter == "minimum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 70 cpm):", command=self.set_min_speed)
        elif paramter == "maximum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 400 cpm):", command=self.set_max_speed)
        elif paramter == "include all keys":
            self.master.show_yes_no_popup("Include keys that aren't alphabetical, numerical or otherwise used in writing?", command=self.include_keys)
        elif paramter == "select tag for session":
            self.master.show_menu_popup("Select tag for this session", self.tags, command=self.select_tag, run_command_if_none=False)
        elif paramter == "add tags":
            self.master.show_text_box_popup("Add the tag:", command=self.add_tag)
        elif paramter == "delete tags":
            self.master.show_menu_popup("Select item for deletion", self.tags, command=self.delete_tag, run_command_if_none=False)
        elif paramter == "exclude tags from session":
            self.selected_tag = np.nan
            self.tag_2 = self.master.add_label(title="NONE", row=3, column=2)

    def delete_tag(self, tag):
        self.tags.remove(tag)
        self.rewrite_tags_file()

    def add_tag(self, tag):
        self.tags.append(tag)
        self.rewrite_tags_file()

    def select_tag(self, tag): 
        self.selected_tag = tag
        self.tag_2 = self.master.add_label(title=tag, row=3, column=2)

    def rewrite_tags_file(self):
        with open(r".\tags\tags.txt", "w") as f:
            for tag in self.tags:
                f.write(tag + "\n")

    def include_keys(self, bool_value):
        self.include_all_keys = bool_value

    def set_min_speed(self, speed):
        # Change min speed value
        try:
            self.min_speed = int(speed)
        except:
            self.master.show_error_popup("Error!", "Enter valid number!")
            sleep(1)
            self.options("minimum speed")
    
    def set_max_speed(self, speed):
        # Change max speed value
        try:
            self.max_speed = int(speed) 
        except:
            self.master.show_error_popup("Error!", "Enter valid number!")
            sleep(1)
            self.options("maximum speed")   
    
    def open_options(self):
        # Open options menu
        self.selected_dates.clear()
        self.master.show_menu_popup("Change parameters:",
                                    ["minimum speed", "maximum speed",
                                      "include all keys", "select tag for session", 
                                      "add tags", "delete tags",
                                      "exclude tags from session"],
                                    command=self.options,
                                    run_command_if_none=False)

    def saved_popup(self):
        # Starting date for multi-session report
        self.files = [f for f in listdir(r".\saved") if isfile(join(r".\saved", f))]
        self.master.show_menu_popup("Generate multi-day report, pick a start date:", self.files, command=self.pick_dates, run_command_if_none=False)

    def pick_dates(self, item):
        # End date for multi-session report
        start_index = self.files.index(item)
        self.selected_dates.append(start_index)
        self.master.show_menu_popup("Generate multi-day report, pick an end date:",
                                     self.files[start_index + 1:], command=self.multi_day_report,
                                       run_command_if_none=False)

    def multi_day_report(self, item):
        # Load saved files into df
        end_index = self.files.index(item)
        self.selected_dates.append(end_index)
        if len(self.selected_dates) == 2:
            report_df = report_df = pd.read_csv(f".\\saved\\{self.files[self.selected_dates[0]]}")
            for i in range((self.selected_dates[0] + 1), (self.selected_dates[1] + 1)):
                next_df = pd.read_csv(f".\\saved\\{self.files[i]}")
                report_df = pd.concat([report_df, next_df], ignore_index=True)
            report_df["timestamp"] = pd.to_datetime(report_df.timestamp)
            self.generate_report(report_df, mode=1)
        else: 
            self.master.show_error_popup("Error!", "There was a problem with picking dates, try again.")

if __name__ == "__main__":
    root = py_cui.PyCUI(6, 6)
    root.set_title("Writing Speed") 
    s = WritingSpeedApp(root)
    root.start()


