#!/usr/bin/python
# -*- coding: utf-8 -*-

# guider.py

import time
import threading
import win32com.client

class CameraThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.cam = None

    def run(self):
        win32com.client.pythoncom.CoInitialize()
        self.cam = win32com.client.Dispatch("ASCOM.SXGuide0.Camera")
        print(self.cam.Connected)
        self.cam.Connected = False
        print(self.cam.Connected)
        self.cam.Connected = True
        print(self.cam.Connected)
        time.sleep(0.1)
        self.cam.Connected = False
        print(self.cam.Connected)
        self.cam = None
        win32com.client.pythoncom.CoUninitialize()
            
if __name__ == '__main__':
    for i in range(10):
        camthread = CameraThread()
        try:
            camthread.run()
        except Exception, e:
            print i, 'failed'
            print e
        else:
            print i, 'ok'
        time.sleep(i*5)