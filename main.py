

from win32gui import GetWindowText, GetForegroundWindow # Processes, paths, threading
from win32process import GetWindowThreadProcessId
import psutil
from pathlib import Path
import threading

from time import sleep # Time 
from datetime import datetime

import numpy as np # Data handling
import pandas as pd

from pynput import keyboard #Keyboard input

import py_cui # UI

import seaborn as sns #Plots
import matplotlib.pyplot as plt

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
        self.ESC_KEY = keyboard.Key.f10
        self.REFRESH_TIME = 1 * 20
        self.max_speed = 400
        self.min_speed = 70
        self.running_status = False
        self.status_text = "OFF"

        # Set up CUI
        self.master = master
        self.title_label = self.master.add_label(title="Writing", row=0, column=1)
        self.start_button = self.master.add_button("START", 3, 1, command=self.start)
        self.stop_button = self.master.add_button("STOP", 3, 2, command=self.stop)
        self.saved_button = self.master.add_button("SAVED", 4, 1, command=None)
        self.options_button = self.master.add_button("OPTIONS", 4, 2, command=self.open_options)
        self.status_label = self.master.add_label(title=self.status_text, row=1, column=1)

        # Set up variables for threads
        self.listener = None
        self.listener_thread = None
        self.aggregate_thread = None
        self.aggregate_running = False
        
        # Set up folders 
        Path(r".\saved").mkdir(exist_ok=True)
        Path(r".\graphs").mkdir(exist_ok=True)
        Path(r".\pdf").mkdir(exist_ok=True)

        #Set up aesthetics
        sns.set_palette("flare")
        sns.set_style("darkgrid")


    def get_process_name(self):
        #Get name of the current process
        try:
            window_process_id = GetWindowThreadProcessId(GetForegroundWindow())
            process_name = psutil.Process(window_process_id[-1]).name()
            return process_name
        except:
            return np.nan
        
    def on_relase(self, key):
        # Put keyboard event into temporary df
        time = datetime.now()
        process = self.get_process_name()
        try:
            event_dict= {
                "hour" : time.hour,
                "minute" : time.minute,
                "process" : process,
            }
            self.events.append(event_dict)
            if key == keyboard.Key.f10:
                self.generate_report()
            elif key == keyboard.Key.f8:
                self.save_record()
        except AttributeError:
            pass

    def aggregate_events(self):
        # Aggregate data from df list to get cmp; move to main df
        try:
            processed_events = self.events.copy()
            events_df = pd.DataFrame(processed_events)
            self.events.clear()
            events_df = events_df.groupby(["hour", "minute"]).agg(
                strokes_per_minute =("process", "count"), 
                process = ("process", lambda x: x.mode()[0]),
            )
            print(events_df.shape)
            print(events_df)
            events_array = events_df.to_dict('records') 
            for event in events_array:
                event["timestamp"] = datetime.now()
                self.per_minute_events.append(event)
        except KeyError:
            pass
        
        print("refreshing")
        
        # Run recursively
        sleep(self.REFRESH_TIME)
        if self.aggregate_running:
            self.aggregate_events()
        else:
            return 

    def generate_report(self):
        report_df = self.save_record()
        timestamp = datetime.now()

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
                                y="speed")
        fig_process.set_title("CPM by process")
        fig_process.set_ylabel("Characters Per Minute")
        fig_process.set_xlabel("Used Process")
        fig1 = fig_process.get_figure()
        fig1.savefig(r".\graphs\barplot.png")
        fig1.clf()
       

        # Generate graph with timeline
        speed_by_time_data = report_df
        speed_by_time_data["time"] = report_df["timestamp"].dt.strftime('%H:%M')
        fig_time = sns.lineplot(data=report_df,
                                x="time",
                                y="strokes_per_minute",
                                errorbar=None)
        fig_time.tick_params(axis='x', rotation=90, labelsize=6)
        fig_time.set_title("CPM by time")
        fig_time.set_ylabel("Characters per Minute")
        fig_time.set_xlabel("Timeline")
        fig2 = fig_time.get_figure()
        fig2.savefig(r".\graphs\lineplot.png")
        fig2.clf()
    

        # Create Reportlab PDF
        try:
            filename = f"report{timestamp.date()}_{timestamp.hour}-{timestamp.minute}.pdf"
            doc =  canvas.Canvas(filename, pagesize=letter)
            doc.setLineWidth(.3)
            pdfmetrics.registerFont(TTFont("DejaVu-Sans", r".\font\DejaVuSans.ttf"))
            doc.setFont("DejaVu-Sans", 12)
            doc.drawString(30,750,f"Date: {timestamp.date()}, {timestamp.hour}:{timestamp.minute}")
            doc.drawString(30, 735, f"Total characters: {total_characters} characters")
            doc.drawString(30, 720, f"Average speed: {average_speed} cmp")
            doc.drawString(30, 705, f"Peak Speed: {peak_speed} cpm")
            doc.drawImage(r".\graphs\barplot.png", 30, 450, 320, 240)
            doc.drawImage(r".\graphs\lineplot.png", 30, 200, 320, 240)
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
            return report_df


    def start(self):
        # Start listening for keyboard
        print("starting")
        self.running_status = True
        self.status_text = "Running..."

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
            self.status_text = "OFF"
            self.listener.stop()
            self.generate_report()
            self.listener = None
            self.listener_thread = None
            self.aggregate_running = False
            self.aggregate_thread = None
            self.running_status = False
        else:
            self.master.show_error_popup("Error!", "Tracking is not running")
        
    def options(self, paramter):
        # Set min/max speed that will be considered in aggregating report
        if paramter == "minimum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 70 cpm):", command=self.set_min_speed)
        elif paramter == "maximum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 400 cpm):", command=self.set_max_speed)

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
        self.master.show_menu_popup("Change parameters:",
                                    ["minimum speed", "maximum speed"],
                                    command=self.options,
                                    run_command_if_none=False)




if __name__ == "__main__":
    root = py_cui.PyCUI(7, 6)
    root.set_title("Writing Speed") 
    s = WritingSpeedApp(root)
    root.start()


