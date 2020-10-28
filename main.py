# -*- coding:utf-8 -*-

import win32serviceutil
import win32service
import win32event
import servicemanager
import socket

import datetime
import time
import os 
import threading
import logging
import random
from win32gui import GetWindowText, GetForegroundWindow
from pypresence import Presence

path = "c:\\drawing_time.log" # Total time log
client_id = 'client id here' # Discord app client id
RPC = Presence(client_id)  # Initialize the client class
RPC.connect() # Start the handshake loop

def get_active_window_title():
    return GetWindowText(GetForegroundWindow())

logging.basicConfig(
    filename = 'c:\\test-service.log',
    level = logging.DEBUG, 
    format="%(asctime)s:LINE[%(lineno)s] %(levelname)s %(message)s"
)

INTERVAL = 15

class MySvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "check-Drawing-RPC"
    _svc_display_name_ = "Check Drawing RPC"

    def __init__(self,args):
        if os.path.isfile(path):
            with open(path) as f:
                temp = f.read()
                if not temp:
                    temp = 0
        else:
            with open(path, mode="w"):
                pass
            temp = 0

        self.d_time = datetime.timedelta(seconds=int(temp))
        self.d_time_delta = datetime.timedelta(seconds=INTERVAL)

        win32serviceutil.ServiceFramework.__init__(self,args)
        self.stop_event = win32event.CreateEvent(None,0,0,None)
        socket.setdefaulttimeout(60)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        RPC.close() # Closes RPC connection 
        with open(path, mode="w") as f:
            f.write(str( int(self.d_time.total_seconds()) ))

        logging.info('Request to Stop Service...')
        self.stop_requested = True

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_,'')
        )

        self.main_loop()

    def get_h_m_s(self, td: datetime.timedelta):
        m, s = divmod(td.seconds, 60)
        h, m = divmod(m, 60)
        return h, m, s

    def main_task(self):
        if get_active_window_title == "CLIP STUDIO PAINT":
            self.d_time += self.d_time_delta
            h, m, s = self.get_h_m_s(self.d_time)
            message = "累計お絵かき時間: {h}時間 {m}分"
            RPC.update( pid=12345, 
                        details="Drawing", 
                        state=message, 
                        large_image="Clipstudio", 
                        large_text="Clip Studio Paint",
                        small_image="Drawing",
                        small_text="Drawing")

        else: # inactive
            h, m, s = self.get_h_m_s(self.d_time)
            message = "累計お絵かき時間: {h}時間 {m}分"
            RPC.update( pid=12345, 
                        details="Inactive", 
                        state=message, 
                        large_image="Clipstudio", 
                        large_text="Clip Studio Paint",
                        small_image="Inactive",
                        small_text="Inactive")

    # メインループ関数
    def main_loop(self):
        logging.info('Starting Service...')
        exec_time = time.time()

        while True: 
            if self.stop_requested:
                logging.info('Attempt to terminate.')
                break

            try: 
                # 実行日時を超えている場合
                if exec_time <= time.time():
                    self.main_task()
                    exec_time = exec_time + INTERVAL

            except Exception as e:
                logging.error("Error occured.")

            time.sleep(1)

        logging.info("Service Stopped")
        return

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(MySvc)