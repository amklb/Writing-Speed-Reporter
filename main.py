from win32gui import GetWindowText, GetForegroundWindow
from win32process import GetWindowThreadProcessId
from time import sleep
import numpy as np
import psutil

def get_process_name():
    try:
        window_process_id = GetWindowThreadProcessId(GetForegroundWindow())
        process_name = psutil.Process(window_process_id[-1]).name()
        process_name = process_name.split(".")[: -1]
        return process_name
    except:
        return np.nan