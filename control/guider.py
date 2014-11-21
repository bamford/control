#!/usr/bin/python
# -*- coding: utf-8 -*-

# guider.py

# simulate obtaining images for testing
simulate = True

import numpy as np
import scipy.stats
import time
from Queue import Queue

import wx
import threading
import serial

from sxvao import SXVAO

# ------------------------------------------------------------------------------
# Event to signal that a new image is ready for use
myEVT_IMAGEREADY = wx.NewEventType()
EVT_IMAGEREADY = wx.PyEventBinder(myEVT_IMAGEREADY, 1)
class ImageReadyEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_IMAGEREADY, eid=wx.ID_ANY, image=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.image = image

# ------------------------------------------------------------------------------        
# Event to signal a new log entry is pending
myEVT_LOG = wx.NewEventType()
EVT_LOG = wx.PyEventBinder(myEVT_LOG, 1)
class LogEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_LOG, eid=wx.ID_ANY, text=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.text = text

# ------------------------------------------------------------------------------
# Class to obtain images on a separate thread
# When run, this connects to the camera and starts taking repeated images
# with the given exposure time until the stopevent flag is set.
# The camera is disconnected before ending.
class TakeImageThread(threading.Thread):
    def __init__(self, parent, stopevent, exptime):
        threading.Thread.__init__(self)
        self.parent = parent
        self.stopevent = stopevent
        self.exptime = exptime
        self.cam = None

    def run(self):
        self.InitCamera()
        self.Log('Started camera')
        while not self.stopevent.is_set():
            self.TakeImage(self.exptime)
        self.Log('Stopped camera')
        self.Disconnect()

    def InitCamera(self):
        if not simulate:
            self.cam = win32com.client.Dispatch("ASCOM.SXGuider0.Camera")
        else:
            self.Log("Simulating camera")
            self.cam = None
        self.Connect()
        
    def Connect(self):
        if self.cam is not None:
            self.cam.Connected = True
            if self.cam.Connected:
                self.Log("Connected to camera")
            else:
                self.Log("Unable to connect to camera")

    def Disconnect(self):
        if self.cam is not None:
            self.cam.Connected = False
            if not self.cam.Connected:
                self.Log("Disconnected from camera")
            else:
                self.Log("Unable to disconnect from camera")
            
    def Log(self, text):
        wx.PostEvent(self.parent, LogEvent(text=text))
        
    def TakeImage(self, exptime):
        if self.cam is not None:
            self.cam.StartExposure(exptime, True)
            time.sleep(exptime)
            while not self.cam.ImageReady:
                time.sleep(0.01)
            image = np.array(self.cam.ImageArray)
        else:
            time.sleep(exptime)
            size = 23
            shape = (600, 400)
            flux = 10000.0 * exptime
            g = scipy.stats.norm.pdf(np.arange(size), (size-1)/2.0, 2.0)
            star = np.dot(g[:, None], g[None, :])
            image = np.zeros(shape)
            x = shape[0]//2 - size//2
            y = shape[1]//2 - size//2
            image[x:x+size,y:y+size] += star * flux
            image = np.random.poisson(image)
            image += np.random.normal(800, 20, size=shape)
        wx.PostEvent(self.parent, ImageReadyEvent(image=image))

# ------------------------------------------------------------------------------
# Class to run AO unit on a separate thread
# When run, this connects to the AO unit and listens to a Queue
# for corrections to make, until the stopevent flag is set.
# The AO unit is disconnected before ending.
class AOThread(threading.Thread):
    def __init__(self, parent, stopevent, corrections,
                 comport, timeout):
        threading.Thread.__init__(self)
        self.parent = parent
        self.stopevent = stopevent
        self.corrections = corrections
        self.comport = comport
        self.timeout = timeout
        self.AO = None

    def run(self):
        if not simulate:
            self.AO = SXVAO(self.comport, self.timeout)
            ok = self.AO.Connect()
        else:
            self.Log('Simulating AO')
            ok = True
        if not ok:
            self.Log('Failed to start AO')
        else:
            self.Log('Started AO')
            while not self.stopevent.is_set():
                dx, dy = self.corrections.get()
                if abs(dx) > 1e-6 or abs(dy) > 1e-6:
                    if not simulate:
                        ok = self.AO.MakeCorrection(dx, dy)
                    if ok:
                        self.Log('Performed AO correction '
                                 '({:.2f},{:.2f})'.format(dx, dy))
                    else:
                        self.Log('Failed to perform AO correction')
                self.corrections.task_done()
            self.Log('Stopped AO')
            self.AO.Disconnect()

    def Log(self, text):
        wx.PostEvent(self.parent, LogEvent(text=text))

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
            self.Destroy()
        else:
            self.parent.panel.ToggleGuider(e)

        
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
        self.camera_on = False
        self.guiding_on = False
        # These positions are stored in numpy image pixel coordinates,
        # so zero-indexed and on the native image scale
        self.guide_box_position = None
        self.guide_centroid = None
        self.image = None
        self.InitPanel()

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
        self.ToggleGuidingButton = wx.Button(panel, label='Start Guiding')
        self.ToggleGuidingButton.Bind(wx.EVT_BUTTON, self.ToggleGuiding)
        self.ToggleGuidingButton.SetToolTip(wx.ToolTip(
            'Start/stop guiding images'))
        self.ToggleGuidingButton.Disable()
        box.Add(self.ToggleGuidingButton, flag=wx.ALIGN_CENTER_VERTICAL|wx.ALL,
                border=10)
        self.logger = wx.TextCtrl(panel, size=(300,90),
                        style=wx.TE_MULTILINE|wx.TE_READONLY)
        box.Add(self.logger, 1, flag=wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)

    def ToggleCamera(self, e):
        if self.camera_on:
            self.StopCamera()
            self.ToggleCameraButton.SetLabel('Start Camera')
            self.ToggleGuidingButton.Disable()
        else:
            self.StartCamera()
            self.ToggleCameraButton.SetLabel('Stop Camera')
            if self.guide_box_position is not None:
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
        self.camera_on = True
        self.stop_camera = threading.Event()
        self.ImageTaker = TakeImageThread(self, self.stop_camera, exptime)
        self.ImageTaker.start()

    def StopCamera(self):
        self.stop_camera.set()
        self.camera_on = False

    def StartGuiding(self):
        self.guiding_on = True
        self.stop_guiding = threading.Event()
        self.AOcorrections = Queue()
        self.AO = AOThread(self, self.stop_camera, self.AOcorrections,
                           self.comport, self.timeout)
        self.AO.start()

    def StopGuiding(self):
        self.stop_guiding.set()
        self.guiding_on = False
        
    def UpdateImageDisplay(self):
        wd, hd = self.ImageDisplay.Size
        wi, hi = self.image.shape
        # scale image levels from 5th to 100th percentile
        imin, imax = np.percentile(self.image, (5.0, 100.0))
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
            dc.SetBrush(wx.Brush(wx.Colour(0, 255, 0, 31)))
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
        wd, hd = self.ImageDisplay.Size
        if self.image is not None:
            wi, hi = self.image.shape
        else:
            wi, hi = wd, hd
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
        self.ToggleGuidingButton.Enable()
        self.Log('Guide box centred at ({:d},{:d})'.format(
            self.guide_box_position.x, self.guide_box_position.y))
        
    def OnImageReady(self, event):
        self.image = event.image
        if self.guiding_on:
            self.Guide()
        self.UpdateImageDisplay()

    def Guide(self):
        dx, dy = self.CentroidBox()
        dx = dx if (abs(dx) > self.min_guide_correction) else 0.0
        dy = dy if (abs(dy) > self.min_guide_correction) else 0.0
        self.AOcorrections.put((dx, dy))
            
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
