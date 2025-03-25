from win32gui import GetWindowText, GetForegroundWindow
from win32process import GetWindowThreadProcessId
from time import sleep
import numpy as np
import psutil
from pynput import keyboard
from datetime import datetime
from time import sleep
import pandas as pd

events = {"hour" : [],
           "minute" : [],
             "process" : []}
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
        events["hour"].append(time.hour)
        events["minute"].append(time.minute)
        events["process"].append(process)
        if key == ESC_KEY:
            return False
    except AttributeError:
        pass

def aggregate_events():
    try:
        processed_events = events.copy()
        events_df = pd.DataFrame(processed_events)
        events = {"hour" : [], "minute" : [], "process" : []}
        events_df.groupby("minute").agg(
            strokes_per_minute =("process", "count"), 
            used_process = ("process", lambda x: x.mode()[0])
        )
    except:
        pass
    print("refreshing")
    print(events_df)
    aggregate_events()





if __name__ == "__main__":
    time = 60
    while time > 0:
        with keyboard.Listener(on_press=on_press) as listener:
            time -= 1
            print(time)
            sleep(1)
            listener.join()
            
    aggregate_events()



