from win32gui import GetWindowText, GetForegroundWindow
from win32process import GetWindowThreadProcessId
from time import sleep
import numpy as np
import psutil
from pynput import keyboard
from datetime import datetime
from time import sleep
import pandas as pd
import py_cui
import seaborn as sns
import threading
from pylatex import Document, Figure, Section

class WritingSpeedApp():
    def __init__(self, master : py_cui.PyCUI):
        self.events = []
        self.per_minute_events = []
        self.writing_speed_per_minute = []
        self.ESC_KEY = keyboard.Key.f10
        self.REFRESH_TIME = 1 * 20
        self.max_speed = 400
        self.min_speed = 70
        self.running_status = False
        self.status_text = "OFF"


        self.master = master

        self.title_label = self.master.add_label(title="Writing", row=0, column=1)
        self.start_button = self.master.add_button("START", 3, 1, command=self.start)
        self.stop_button = self.master.add_button("STOP", 3, 2, command=self.stop)
        self.saved_button = self.master.add_button("SAVED", 4, 1, command=None)
        self.options_button = self.master.add_button("OPTIONS", 4, 2,
                                                     command=self.open_options)
        self.status_label = self.master.add_label(title=self.status_text, row=1, column=1)

        self.listener = None
        self.listener_thread = None
        self.aggregate_thread = None
        self.aggregate_running = False

    def get_process_name(self):
        try:
            window_process_id = GetWindowThreadProcessId(GetForegroundWindow())
            process_name = psutil.Process(window_process_id[-1]).name()
            return process_name
        except:
            return np.nan
        
    def on_press(self, key):
        print(key)
        time = datetime.now()
        process = self.get_process_name()
        try:
            print('alphanumeric key pressed')
            print(time.minute)
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
        fig2 = fig_time.get_figure()
        fig2.savefig(r".\graphs\lineplot.png")
        fig2.clf()
        

        # Creating a LaTeX document
        l_doc =Document(documentclass='report')

        with l_doc.create(Section("Today's writing speed report:")):
            l_doc.append(f"Date: {timestamp.date()}, {timestamp.hour}:{timestamp.minute}\n")
            l_doc.append(f"Total characters: {total_characters} characters\n")
            l_doc.append(f"Peak Speed: {peak_speed} cpm\n")
            l_doc.append(f"Average speed: {average_speed: .2f} cmp\n")
            with l_doc.create(Figure(position="H")) as fig1:
                fig1.add_image(r".\graphs\barplot.png", width="300px")
                fig1.add_caption("Average CPM per app used")
            with l_doc.create(Figure(position="H")) as fig2:
                fig2.add_image(r".\graphs\lineplot.png", width="300px")
                fig2.add_caption("CPM during writing time")

        l_doc.generate_pdf(f".\\saved\\{timestamp.date}_{datetime.hour}-{datetime.minute}", clean_tex=False, compiler="pdflatex")


        self.master.show_message_popup("Report Generated!", "Report saved in the PDF folder!")
        
        
    def save_record(self):
        report_df = pd.DataFrame(self.per_minute_events)
        self.master.show_error_popup("T", report_df.empty)
        print(report_df.empty)
        print(report_df)
        if report_df.empty == True:
            self.master.show_warning_popup("Warning!", "Not enough data to generate report!")
        else:
            print(report_df)
            date = datetime.now()
            path = f".\\saved\\record_{date.year}-{date.month}-{date.day}_{date.hour}-{date.minute}.csv"
            report_df.to_csv(path_or_buf=path)
            return report_df


    def start(self):
        print("starting")
        self.running_status = True
        self.status_text = "Running..."
        if self.listener and self.listener.running: #if thread already started
            self.master.show_warning_popup("Error", "Recording is already running!")
        
        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener_thread = threading.Thread(target=self.listener.start, daemon=True)
        self.aggregate_thread = threading.Thread(target=self.aggregate_events, daemon=True)
        self.listener_thread.start()
        self.aggregate_thread.start()
        self.aggregate_running = True
        

    def stop(self):
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
        print(paramter)
        if paramter == "minimum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 70 cpm):", command=self.set_min_speed)
        elif paramter == "maximum speed":
            self.master.show_text_box_popup("Minimum wiring speed considered (usually 400 cpm):", command=self.set_max_speed)

    def set_min_speed(self, speed):
        try:
            self.min_speed = int(speed)
        except:
            self.master.show_error_popup("Error!", "Enter valid number!")
            sleep(1)
            self.options("minimum speed")
    
    def set_max_speed(self, speed):
        try:
            self.max_speed = int(speed) 
        except:
            self.master.show_error_popup("Error!", "Enter valid number!")
            sleep(1)
            self.options("maximum speed")   
    
    def open_options(self):
        self.master.show_menu_popup("Change parameters:",
                                    ["minimum speed", "maximum speed"],
                                    command=self.options,
                                    run_command_if_none=False)




if __name__ == "__main__":
    root = py_cui.PyCUI(7, 6)
    root.set_title("Writing Speed") 
    s = WritingSpeedApp(root)
    root.start()


