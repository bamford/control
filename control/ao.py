import wx
import threading

from sxvao import SXVAO
from logevent import *

# ------------------------------------------------------------------------------
# Class to run AO unit on a separate thread
# When run, this connects to the AO unit and listens to a Queue
# for corrections to make, until the 'Q' command is received.
# The AO unit is disconnected before ending.
class AOThread(threading.Thread):
    def __init__(self, parent, corrections,
                 comport, timeout):
        threading.Thread.__init__(self)
        self.parent = parent
        self.corrections = corrections
        self.comport = comport
        self.timeout = timeout
        self.minsteptime = 0.1  # seconds
        self.AOunit = None

    def run(self):
        if not simulate:
            self.AOunit = SXVAO(self, self.comport, self.timeout)
            ok = self.AOunit.Connect()
        else:
            self.Log('Simulating AO')
            ok = True
        if not ok:
            self.Log('Failed to start AO')
        else:
            self.Log('Started AO')
            try:
                last_step_time = 0
                while True:
                    # avoid sending corrections too quickly to AO unit
                    dt = time.time() - last_step_time
                    time.sleep(max(0,  self.minsteptime - dt))
                    # get the last thing in the queue, in a way that
                    # avoids never doing anything is the queue is
                    # currently filling faster than we can empty it
                    c = self.corrections.get()
                    n = self.corrections.qsize()
                    while n > 0:
                        n -= 1
                        if c in ['Q', 'K']:
                            break
                        c = self.corrections.get()
                    # process received command
                    if c == 'Q':
                        # quit guiding
                        break
                    elif c == 'K':
                        # centre AO unit
                        if not simulate:
                            ok = self.AOunit.Centre()
                        if ok:
                            self.Log('Centred AO unit')
                        else:
                            self.Log('AO unit centring failed')
                        last_step_time = time.time()
                    else:
                        # expect (command, dx, dy) correction,
                        # don't do anything if they are both an
                        # insignificant fraction of a pixel
                        done = self.GetAndPerformCorrection(c)
                        if done:
                            last_step_time = time.time()
            finally:
                if self.AOunit is not None:
                    self.AOunit.Disconnect()
                self.Log('Stopped AO')

    def GetAndPerformCorrection(self, c):
        unknown = True
        ok = True
        try:
            command, dx, dy = c
            zero = abs(dx) < 1e-3 and abs(dy) < 1e-3
        except:
            pass
        else:
            if command == 'G':
                unknown = False
                if not zero:
                    if not simulate:
                        ok = self.AOunit.MakeCorrection(dx, dy)
                    if ok:
                        self.Log('Performed AO correction '
                            '({:.2f},{:.2f})'.format(dx, dy))
                    else:
                        self.Log('Failed to perform AO correction')
            elif command == 'M':
                unknown = False
                if not zero:
                    if not simulate:
                        ok = self.AOunit.MakeMountCorrection(dx, dy)
                    if ok:
                        self.Log('Performed AO mount correction '
                            '({:.2f},{:.2f})'.format(dx, dy))
                    else:
                        self.Log('Failed to perform AO mount correction')
        if unknown:
            self.Log('Unknown AO correction '
                     '({})'.format(c))
                
    def Log(self, text):
        wx.PostEvent(self.parent, LogEvent(text=text))

    def toggle_switch_xy(self):
        if self.AOunit is not None:
            self.AOunit.switch_xy = not self.AOunit.switch_xy
            self.Log('switch_xy = {}'.format(self.AOunit.switch_xy))

    def toggle_reverse_x(self):
        if self.AOunit is not None:
            self.AOunit.reverse_x = not self.AOunit.reverse_x
            self.Log('reverse_x = {}'.format(self.AOunit.reverse_x))

    def toggle_reverse_y(self):
        if self.AOunit is not None:
            self.AOunit.reverse_y = not self.AOunit.reverse_y
            self.Log('reverse_y = {}'.format(self.AOunit.reverse_y))

    def adjust_steps_per_pixel(self, factor):
        if self.AOunit is not None:
            self.AOunit.steps_per_pixel /= factor
            self.Log('steps_per_pixel = {:.2f}'.format(self.AOunit.steps_per_pixel))

    def toggle_mount_switch_xy(self):
        if self.AOunit is not None:
            self.AOunit.mount_switch_xy = not self.AOunit.mount_switch_xy
            self.Log('mount_switch_xy = {}'.format(self.AOunit.mount_switch_xy))

    def toggle_mount_reverse_x(self):
        if self.AOunit is not None:
            self.AOunit.mount_reverse_x = not self.AOunit.mount_reverse_x
            self.Log('mount_reverse_x = {}'.format(self.AOunit.mount_reverse_x))

    def toggle_mount_reverse_y(self):
        if self.AOunit is not None:
            self.AOunit.mount_reverse_y = not self.AOunit.mount_reverse_y
            self.Log('mount_reverse_y = {}'.format(self.AOunit.mount_reverse_y))

    def adjust_mount_steps_per_pixel(self, factor):
        if self.AOunit is not None:
            self.AOunit.mount_steps_per_pixel /= factor
            self.Log('mount_steps_per_pixel = {}'.format(self.AOunit.mount_steps_per_pixel))
        
