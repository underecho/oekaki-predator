# -*- coding:utf-8 -*-
#!/usr/bin/env python3.7

import asyncio
import datetime
import time
import os 
import logging
import random
from win32gui import GetWindowText, GetForegroundWindow
from pypresence import Presence

path = "c:\\Users\\under\\drawing_time.log" # Total time log
client_id = '771084728300339251' # Discord app client id

def get_active_window_title():
    return GetWindowText(GetForegroundWindow())

logging.basicConfig(
    filename = 'c:\\Users\\under\\test-service.log',
    level = logging.DEBUG, 
    format="%(asctime)s:LINE[%(lineno)s] %(levelname)s %(message)s"
)

INTERVAL = 15

class MySvc:

    def __init__(self):
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

        import selectors
        selector = selectors.SelectSelector()
        loop = asyncio.SelectorEventLoop(selector)
        self.RPC = Presence(client_id, loop=loop)  # Initialize the client class
        self.RPC.connect() # Start the handshake loop

        self.main_loop()

    def get_h_m_s(self, td: datetime.timedelta):
        m, s = divmod(td.seconds, 60)
        h, m = divmod(m, 60)
        return h, m, s

    def main_task(self):
        logging.info(get_active_window_title())
        if get_active_window_title() == "CLIP STUDIO PAINT":
            self.d_time += self.d_time_delta
            h, m, s = self.get_h_m_s(self.d_time)
            message = f"累計お絵かき時間: {h}:{m}"
            self.RPC.update( pid=12345, 
                        details="Drawing", 
                        state=message, 
                        large_image="clipstudio", 
                        large_text="Clip Studio Paint",
                        small_image="drawing",
                        small_text="Drawing")

        else: # inactive
            h, m, s = self.get_h_m_s(self.d_time)
            message = f"累計お絵かき時間: {h}:{m}"
            self.RPC.update( pid=12345, 
                        details="Inactive", 
                        state=message, 
                        large_image="clipstudio", 
                        large_text="Clip Studio Paint",
                        small_image="inactive",
                        small_text="Inactive")

    # メインループ関数
    def main_loop(self):
        logging.info('Starting Service...')
        exec_time = time.time()
        write_time = time.time()

        while True:
            try: 
                # 実行日時を超えている場合
                if exec_time <= time.time():
                    self.main_task()
                    exec_time = exec_time + INTERVAL

                if write_time <= time.time():
                    with open(path, mode="w") as f:
                        f.write(str( int(self.d_time.total_seconds()) ))
                    write_time += INTERVAL * 10

            except Exception as e:
                logging.error("Error occured.")

            time.sleep(3)

        logging.info("Service Stopped")
        return

if __name__ == '__main__':
    MySvc()