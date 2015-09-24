#!/usr/bin/python
# -*- coding: utf-8 -*-

# guider.py

from __future__ import print_function

# simulate obtaining images for testing
simulate = False

import numpy as np
import scipy.stats
import time
from datetime import datetime
from Queue import Queue

import wx
import threading
import serial
if not simulate:
    try:
        import win32com.client
    except ImportError:
        simulate = True

from camera import TakeGuiderImageThread, EVT_IMAGEREADY
from ao import AOThread
from logevent import EVT_LOG

# ------------------------------------------------------------------------------
# The main Guider frame
class Guider(wx.Frame):

    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, title='Guider',
                          size=(600, 600), **kwargs)
        self.parent = args[0]
        self.SetMinSize((400, 400))
        self.__DoLayout()
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.Bind(EVT_LOG, self.panel.OnLog)
        self.Bind(EVT_IMAGEREADY, self.panel.OnImageReady)
        if self.parent is None:
            self.Show(True)

    def __DoLayout(self):
        self.panel = GuiderPanel(self)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(self.panel, 1, wx.EXPAND)
        self.SetSizer(sizer)

    def OnQuit(self, e):
        # When this frame is quit, if it was started by the Control frame,
        # then just hide the frame, otherwise close it completely
        if self.parent is None:
            self.panel.stop_camera.set()
            self.Destroy()
        else:
            self.parent.panel.ToggleGuider(e)

    def OnExit(self, e):
        self.panel.stop_camera.set()
        self.Destroy()


# ------------------------------------------------------------------------------
# The main Guider panel
class GuiderPanel(wx.Panel):

    def __init__(self, *args, **kwargs):
        wx.Panel.__init__(self, *args, **kwargs)
        self.main = args[0]
        # config start
        self.comport = 4
        self.timeout = 10  # seconds
        self.dark = None
        self.default_exptime = 1.0  # seconds
        self.guide_box_size = 25  # pixels
        self.min_guide_correction = 0.1  # pixels
        # config end
        self.guiding_on = False
        # These positions are stored in numpy image pixel coordinates,
        # so zero-indexed and on the native image scale
        self.guide_box_position = None
        self.guide_centroid = None
        self.image = None
        self.imagecount = None
        self.AOtrained = False
        self.InitPanel()
        wx.CallLater(50, self.InitAO)
        wx.CallLater(100, self.InitCamera)

    def InitPanel(self):
        MainBox = wx.BoxSizer(wx.VERTICAL)
        sb = wx.StaticBox(self)
        ImageBox = wx.StaticBoxSizer(sb, wx.HORIZONTAL)
        self.ImageDisplay = wx.StaticBitmap(self,
                                bitmap=wx.EmptyBitmap(600, 400))
        self.ImageDisplay.Bind(wx.EVT_LEFT_UP, self.OnClickImage)
        ImageBox.Add((1,1), 1, flag=wx.EXPAND)
        ImageBox.Add(self.ImageDisplay, 10,
                     flag=wx.ALIGN_CENTER | wx.ALL | wx.SHAPED,
                     border=10)
        ImageBox.Add((1,1), 1, flag=wx.EXPAND)
        MainBox.Add(ImageBox, 1, flag=wx.EXPAND)
        sb = wx.StaticBox(self)
        CameraButtonBox = wx.StaticBoxSizer(sb, wx.HORIZONTAL)
        self.InitCameraButtons(self, CameraButtonBox)
        MainBox.Add(CameraButtonBox, 0, flag=wx.EXPAND)
        sb = wx.StaticBox(self)
        GuidingButtonBox = wx.StaticBoxSizer(sb, wx.HORIZONTAL)
        self.InitGuidingButtons(self, GuidingButtonBox)
        MainBox.Add(GuidingButtonBox, 0, flag=wx.EXPAND)
        self.SetSizer(MainBox)

    def InitCameraButtons(self, panel, box):
        self.ToggleCameraButton = wx.Button(panel, label='Start Camera')
        self.ToggleCameraButton.Bind(wx.EVT_BUTTON, self.ToggleCamera)
        self.ToggleCameraButton.SetToolTip(wx.ToolTip(
            'Start/stop taking guider images'))
        self.ToggleCameraButton.Disable()
        box.Add(self.ToggleCameraButton, flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL,
                border=10)
        box.Add((20, 10))
        box.Add(wx.StaticText(panel, label='Exp.Time'),
                       flag=wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, border=5)
        self.ExpTimeCtrl = wx.TextCtrl(panel, size=(50,-1),)
        self.ExpTimeCtrl.ChangeValue('{:.3f}'.format(self.default_exptime))
        self.ExpTimeCtrl.SetToolTip(wx.ToolTip(
            'Exposure time for guider camera'))
        box.Add(self.ExpTimeCtrl, flag=wx.ALIGN_CENTER_VERTICAL)
        box.Add(wx.StaticText(panel, label='sec'),
                       flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border=5)

    def InitGuidingButtons(self, panel, box):
        self.TrainGuidingButton = wx.Button(panel, label='Train Guiding')
        self.TrainGuidingButton.Bind(wx.EVT_BUTTON, self.TrainGuiding)
        self.TrainGuidingButton.SetToolTip(wx.ToolTip(
            'Automatically train guiding system'))
        self.TrainGuidingButton.Disable()
        box.Add(self.TrainGuidingButton, 5, flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL,
                border=10)
        self.ToggleGuidingButton = wx.Button(panel, label='Start Guiding')
        self.ToggleGuidingButton.Bind(wx.EVT_BUTTON, self.ToggleGuiding)
        self.ToggleGuidingButton.SetToolTip(wx.ToolTip(
            'Start/stop guiding images'))
        self.ToggleGuidingButton.Disable()
        box.Add(self.ToggleGuidingButton, 5, flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL,
                border=10)
        self.MoveAOLButton = wx.Button(panel, label='L')
        self.MoveAOLButton.Bind(wx.EVT_BUTTON, self.MoveAOLeft)
        self.MoveAOLButton.SetToolTip(wx.ToolTip(
            'Move stars left'))
        box.Add(self.MoveAOLButton, 1, flag=wx.ALIGN_CENTER_VERTICAL,
                border=10)
        self.MoveAORButton = wx.Button(panel, label='R')
        self.MoveAORButton.Bind(wx.EVT_BUTTON, self.MoveAORight)
        self.MoveAORButton.SetToolTip(wx.ToolTip(
            'Move stars right'))
        box.Add(self.MoveAORButton, 1, flag=wx.ALIGN_CENTER_VERTICAL,
                border=10)
        self.MoveAOUButton = wx.Button(panel, label='U')
        self.MoveAOUButton.Bind(wx.EVT_BUTTON, self.MoveAOUp)
        self.MoveAOUButton.SetToolTip(wx.ToolTip(
            'Move stars up'))
        box.Add(self.MoveAOUButton, 1, flag=wx.ALIGN_CENTER_VERTICAL,
                border=10)
        self.MoveAODButton = wx.Button(panel, label='D')
        self.MoveAODButton.Bind(wx.EVT_BUTTON, self.MoveAODown)
        self.MoveAODButton.SetToolTip(wx.ToolTip(
            'Move stars down'))
        box.Add(self.MoveAODButton, 1, flag=wx.ALIGN_CENTER_VERTICAL,
                border=10)
        self.logger = wx.TextCtrl(panel, size=(300,90),
                        style=wx.TE_MULTILINE|wx.TE_READONLY)
        box.Add(self.logger, 1, flag=wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

    def ToggleCamera(self, e):
        if self.camera_on.is_set():
            self.StopCamera()
            self.ToggleCameraButton.SetLabel('Start Camera')
            self.TrainGuidingButton.Disable()
            self.ToggleGuidingButton.Disable()
        else:
            self.StartCamera()
            self.ToggleCameraButton.SetLabel('Stop Camera')
            if self.guide_box_position is not None:
                self.TrainGuidingButton.Enable()
                if self.AOtrained:
                    self.ToggleGuidingButton.Enable()

    def ToggleGuiding(self, e):
        if self.guiding_on:
            self.StopGuiding()
            self.guiding_on = False
            self.ToggleGuidingButton.SetLabel('Start Guiding')
            self.ToggleCameraButton.Enable()
        else:
            self.StartGuiding()
            self.guiding_on = True
            self.ToggleGuidingButton.SetLabel('Stop Guiding')
            self.ToggleCameraButton.Disable()

    def StartCamera(self):
        exptime = self.GetExpTime()
        self.ImageTaker.SetExpTime(exptime)
        self.camera_on.set()

    def StopCamera(self):
        self.camera_on.clear()

    def InitCamera(self):
        exptime = self.GetExpTime()
        self.stop_camera = threading.Event()
        self.camera_on = threading.Event()
        self.ImageTaker = TakeGuiderImageThread(self, self.stop_camera,
                                                self.camera_on, exptime)
        self.ToggleCameraButton.Enable()

    def InitAO(self):
        self.StartGuiding()
        self.AOcorrections.put('K')
        self.StopGuiding()

    def StartGuiding(self):
        self.AOcorrections = Queue()
        self.AO = AOThread(self, self.AOcorrections,
                           self.comport, self.timeout)
        self.AO.start()

    def StopGuiding(self):
        self.guiding_on = False
        self.AOcorrections.put('Q')

    def MoveAOLeft(self, e):
        self.AOcorrections.put(('G', -50.0, 0.0))

    def MoveAORight(self, e):
        self.AOcorrections.put(('G', 50.0, 0.0))

    def MoveAOUp(self, e):
        self.AOcorrections.put(('G', 0.0, 50.0))

    def MoveAODown(self, e):
        self.AOcorrections.put(('G', 0.0, 50.0))

    def TrainGuiding(self, e):
        self.ToggleCameraButton.Disable()
        self.StartGuiding()
        self.Log('Training AO')
        # Train AO unit
        # NEED TO WATCH OUT FOR CASE WHERE WE LOSE GUIDE STAR!
        # COMMENTED OUT ADJUST_STEPS FOR NOW!
        for delta in [0.3, 0.6]:
            dpix = self.guide_box_size * delta
            movebox = delta>0.5
            # check x versus y and get step factor
            dx, dy = self.AObracket(0, dpix, movebox)
            if abs(dx) < abs(dy):
                self.AO.toggle_switch_xy()
                dx, dy = dy, dx
            factor = abs(dx) / dpix
            #self.AO.adjust_steps_per_pixel(factor)
            # check x and y directions
            dx, dy = self.AObracket(dpix, dpix, movebox)
            if dx < 0:
                self.AO.toggle_reverse_x()
            if dy < 0:
                self.AO.toggle_reverse_y()
        # Train AO mount
        for delta in [0.3, 0.6]:
            dpix = self.guide_box_size * delta
            movebox = delta>0.5
            # check x versus y and get step factor
            dx, dy = self.AObracket(0, dpix, movebox, mount=True)
            if abs(dx) < abs(dy):
                self.AO.toggle_mount_switch_xy()
                dx, dy = dy, dx
            factor = abs(dx) / dpix
            #self.AO.adjust_mount_steps_per_pixel(factor)
            # check x and y direction
            dx, dy = self.AObracket(0, dpix, movebox, mount=True)
            if dx < 0:
                self.AO.toggle_mount_reverse_x()
            if dy < 0:
                self.AO.toggle_mount_reverse_y()
        self.AOtrained = True
        self.ToggleGuidingButton.Enable()
        self.ToggleCameraButton.Enable()
        self.Log('AO training complete')

    def AObracket(self, axis, dpix, movebox, mount=False):
        if mount:
            command = 'M'
        else:
            command = 'G'
        self.AOcorrections.put((command, -dpix, 0.0))
        if movebox:
             pass # NEED TO IMPLEMENT THIS
        wx.Yield()
        time.sleep(0.5)
        self.WaitForNextImage()
        dx1, dy1 = self.CentroidBox()
        self.AOcorrections.put((command, 2.0*dpix, 0.0))
        wx.Yield()
        time.sleep(0.5)
        self.WaitForNextImage()
        dx2, dy2 = self.CentroidBox()
        self.AOcorrections.put((command, -dpix, 0.0))
        wx.Yield()
        time.sleep(0.5)
        return (dx2-dx1), (dy2-dy1)

    def WaitForNextImage(self):
        t = self.image_time
        for i in range(100):  # max 10 sec
            #wx.Yield()
            if t != self.image_time or t is None:
                break
            time.sleep(0.1)

    def UpdateImageDisplay(self):
        wd, hd = self.ImageDisplay.Size
        wi, hi = self.image.shape
        # scale image levels from 5th to 100th percentile
        imin, imax = np.percentile(self.image, (5.0, 100.0))
        # but do not exaggerate really low counts
        imax = max(imax, imin+16)
        image = ((self.image-imin)/(imax-imin) * 255).clip(0, 255)
        # convert to RGB (but still greyscale)
        image = image.T.astype('uint8')
        image = np.dstack((image, image, image))
        wxImg = wx.EmptyImage(wi, hi)
        wxImg.SetData( image.tostring() )
        # resize the image to fill sizer, preserving the aspect ratio
        ad = float(hd)/wd
        ai = float(hi)/wi
        if ad > ai:
            hi_new = hd
            wi_new = hd / ai
        else:
            wi_new = wd
            hi_new = wd * ai
        wxImg = wxImg.Scale(wi_new, hi_new)
        bitmap = wxImg.ConvertToBitmap()
        # Add guider box
        if self.guide_box_position is not None:
            x = wd * self.guide_box_position.x / float(wi) + 1
            y = hd * self.guide_box_position.y / float(hi) + 1
            size = self.guide_box_size * wd / float(wi)
            dc = wx.MemoryDC(bitmap)
            dc.SetPen(wx.Pen(wx.Colour(0, 255, 0, 127), 2))
            dc.SetBrush(wx.TRANSPARENT_BRUSH)
            xc, yc, size = self.GetRectCorner(x, y, size)
            dc.DrawRectangle(xc, yc, size, size)
            if self.guide_centroid is not None:
                x = wd * self.guide_centroid.x / wi + 1.5
                y = hd * self.guide_centroid.y / hi + 1.5
                dc.SetPen(wx.Pen(wx.Colour(255, 0, 0, 127), 2))
                dc.DrawCircle(int(round(x)), int(round(y)), size//3)
            dc.SelectObject(wx.NullBitmap)
        # Update display
        self.ImageDisplay.SetBitmap(bitmap)

    def GetRectCorner(self, x, y, size):
        xc, yc = [int(round(c - size / 2.0)) for c in (x, y)]
        size = int(round(size))
        return xc, yc, size

    def OnClickImage(self, event):
        if self.image is not None:
            wd, hd = self.ImageDisplay.Size
            wi, hi = self.image.shape
            pos = event.GetPosition()
            # convert position in ImageDisplay to position in image
            x = int(round(wi * (pos.x-1) / float(wd)))
            y = int(round(hi * (pos.y-1) / float(hd)))
            self.guide_box_position = wx.Point(x, y)
            # move box to centroid
            for i in range(3):
                dx, dy = self.CentroidBox()
                self.guide_box_position.x += int(round(dx))
                self.guide_box_position.y += int(round(dy))
                self.CentroidBox()
            if self.camera_on.is_set():
                self.TrainGuidingButton.Enable()
                if self.AOtrained:
                    self.ToggleGuidingButton.Enable()
            self.Log('Guide box centred at ({:d},{:d})'.format(
                self.guide_box_position.x, self.guide_box_position.y))

    def OnImageReady(self, event):
        self.image = event.image
        self.image_time = event.image_time
        if self.guiding_on:
            self.Guide()
        self.UpdateImageDisplay()

    def Guide(self):
        dx, dy = self.CentroidBox()
        dx = dx if (abs(dx) > self.min_guide_correction) else 0.0
        dy = dy if (abs(dy) > self.min_guide_correction) else 0.0
        self.AOcorrections.put(('G', dx, dy))

    def CentroidBox(self):
        xc, yc, size = self.GetRectCorner(self.guide_box_position.x,
                                          self.guide_box_position.y,
                                          self.guide_box_size)
        subimage = self.image[xc:xc+size,
                              yc:yc+size]
        dx, dy = self.Centroid(subimage)
        self.Log('Centroid within guide box is ({:.2f},{:.2f})'.format(dx, dy))
        x = self.guide_box_position.x + dx
        y = self.guide_box_position.y + dy
        self.Log('Centroid within image is ({:.2f},{:.2f})'.format(x, y))
        self.guide_centroid = wx.Point(x, y)
        return dx, dy

    def Centroid(self, image):
        dx, dy = [self.Centroid1d(np.sum(image, axis)) for axis in (1, 0)]
        return dx, dy

    def Centroid1d(self, array):
        n = len(array)
        p = np.arange(n) - (n-1)/2.0
        w = array - array.min()
        w **= 2
        c = np.average(p, weights=w)
        return c

    def GetExpTime(self):
        try:
            exptime = float(self.ExpTimeCtrl.GetValue())
        except ValueError:
            exptime = None
            self.Log('Exposure time invalid, setting to '
                    '{:.3f} sec'.format(self.default_exptime))
            self.ExpTimeCtrl.ChangeValue('{:.3f}'.format(self.default_exptime))
        return exptime

    def Log(self, text):
        self.logger.AppendText("{}\n".format(text))
        wx.CallAfter(self.logger.Refresh)

    def OnLog(self, event):
        self.Log(event.text)


def main():
    app = wx.App(False)
    Guider(None)
    app.MainLoop()


if __name__ == '__main__':
    main()
