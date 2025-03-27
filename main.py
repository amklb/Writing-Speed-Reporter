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

class WritingSpeedApp():
    def __init__(self, master : py_cui.PyCUI):
        self.events = []
        self.per_minute_events = []
        self.writing_speed_per_minute = []
        self.ESC_KEY = keyboard.Key.f10
        self.REFRESH_TIME = 1 * 20
        # self.master = master

        # self.title_label = self.master.add_label(title="Writing", row=0, column=1)
        # self.start_button = self.master.add_button("START", 3, 1, command=self.start)
        # self.stop_button = self.master.add_button("STOP", 3, 2, command=self.stop)
        # self.saved_button = self.master.add_button("SAVED", 4, 1, command=None)
        self.start()

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
                "process" : process
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
                process = ("process", lambda x: x.mode()[0])
            ).reset_index()
            print(events_df.shape)
            print(events_df)
            events_array = events_df.to_dict('records') 
            for array in events_array:
                self.per_minute_events.append(array)
        except KeyError:
            pass
        
        print("refreshing")
        
        sleep(self.REFRESH_TIME)
        self.aggregate_events()

    def generate_report(self):
        report_df = pd.DataFrame(self.per_minute_events)
        print(report_df)
        speed_by_process_data = report_df.groupby(["process"]).agg(
            speed = ("strokes_per_minute", "mean")
        ).reset_index()
        fig_process = sns.barplot(data=speed_by_process_data,
                                  x="process",
                                  y="speed")
        fig_process = fig_process.get_figure()
        fig_process.savefig("out.png")
        
    def save_record(self):
        report_df = pd.DataFrame(self.per_minute_events)
        print(report_df)
        date = datetime.now()
        report_df.to_csv(path_or_buf=f".\\saved\\record_{date.year}-{date.month}-{date.day}_{date.hour}-{date.minute}.csv")



    def start(self):
        print("starting")
        with keyboard.Listener(on_press=self.on_press) as self.listener:
            
            self.aggregate_events()
            self.listener.join()

    def stop(self):
        self.listener.stop()
        self.generate_report()

if __name__ == "__main__":
    root = py_cui.PyCUI(7, 6)
    root.set_title("Writing Speed") 
    s = WritingSpeedApp(root)
    #root.start()


