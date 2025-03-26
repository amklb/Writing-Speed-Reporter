from win32gui import GetWindowText, GetForegroundWindow
from win32process import GetWindowThreadProcessId
from time import sleep
import numpy as np
import psutil
from pynput import keyboard
from datetime import datetime
from time import sleep
import pandas as pd

events = []
per_minute_events = []
writing_speed_per_minute = []
ESC_KEY = keyboard.Key.f10
REFRESH_TIME = 1 * 60

def get_process_name():
    try:
        window_process_id = GetWindowThreadProcessId(GetForegroundWindow())
        process_name = psutil.Process(window_process_id[-1]).name()
        return process_name
    except:
        return np.nan
    
def on_press(key):
    print(key)
    time = datetime.now()
    process = get_process_name()
    try:
        print('alphanumeric key pressed')
        print(time.minute)
        event_dict= {
            "hour" : time.hour,
            "minute" : time.minute,
            "process" : process
        }
        events.append(event_dict)
        if key == ESC_KEY:
            return False
    except AttributeError:
        pass

def aggregate_events():
    try:
        processed_events = events.copy()
        events_df = pd.DataFrame(processed_events)
        events.clear()
        events_df = events_df.groupby(["hour", "minute"]).agg(
            strokes_per_minute =("process", "count"), 
            used_process = ("process", lambda x: x.mode()[0])
        ).reset_index()
        print(events_df.shape)
        print(events_df)
        events_array = events_df.to_numpy()
        for array in events_array:
            per_minute_events.append(array)
    except KeyError:
        pass
    
    print("refreshing")
    
    sleep(REFRESH_TIME)
    aggregate_events()

def generate_report():
    report_df = pd.DataFrame(per_minute_events)
    



if __name__ == "__main__":
    with keyboard.Listener(on_press=on_press) as listener:
       
        aggregate_events()
        listener.join()



