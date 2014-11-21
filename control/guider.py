#!/usr/bin/python
# -*- coding: utf-8 -*-

# guider.py

#mode = 'live'
#mode = 'sim'
mode = None

import numpy as np
import scipy.stats
import time

import wx
import threading

# Create a new event to signal that a new image is ready
myEVT_IMAGEREADY = wx.NewEventType()
EVT_IMAGEREADY = wx.PyEventBinder(myEVT_IMAGEREADY, 1)
class ImageReadyEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_IMAGEREADY, eid=wx.ID_ANY, image=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.image = image

# Create a new event to signal a new log entry is pending
myEVT_LOG = wx.NewEventType()
EVT_LOG = wx.PyEventBinder(myEVT_LOG, 1)
class LogEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_LOG, eid=wx.ID_ANY, text=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.text = text

# Class to obtain an image in a threaded manner
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
        if mode == 'sim':
            self.cam = win32com.client.Dispatch("ASCOM.Simulator.Camera")
        elif mode == 'live':
            self.cam = win32com.client.Dispatch("ASCOM.SXGuider0.Camera")
        
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
            time.sleep(0.5)
            size = 23
            shape = (600, 400)
            flux = 10000.0
            g = scipy.stats.norm.pdf(np.arange(size), (size-1)/2.0, 4.0)
            star = np.dot(g[:, None], g[None, :])
            image = np.zeros(shape)
            x = shape[0]//2 - size//2
            y = shape[1]//2 - size//2
            image[x:x+size,y:y+size] += star * flux
            image = np.random.poisson(image)
            image += np.random.normal(800, 20, size=shape)
        wx.PostEvent(self.parent, ImageReadyEvent(image=image))


class Guider(wx.Frame):
    
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, title='Guider',
                          size=(400, 400), **kwargs)
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
        if self.parent is None:
            self.Destroy()
        else:
            self.parent.panel.ToggleGuider(e)

        
class GuiderPanel(wx.Panel):

    def __init__(self, *args, **kwargs):
        wx.Panel.__init__(self, *args, **kwargs)
        self.main = args[0]
        # config start
        self.comport = 4
        self.timeout = 0.1
        self.dark = None
        self.default_exptime = 1.0
        self.image_shape = (600, 400)
        # config end
        self.camera_on = False
        self.guiding_on = False
        self.InitPanel()

    def InitPanel(self):
        MainBox = wx.BoxSizer(wx.VERTICAL)        
        sb = wx.StaticBox(self)
        ImageBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.ImageDisplay = wx.StaticBitmap(self,
                                bitmap=wx.EmptyBitmap(*self.image_shape))
        ImageBox.Add(self.ImageDisplay, 1,
                     flag=wx.ALIGN_CENTER | wx.ALL | wx.SHAPED,
                     border=10)
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
        self.logger = wx.TextCtrl(panel, size=(300,50),
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
            self.ToggleGuidingButton.Enable()

    def ToggleGuiding(self, e):
        if self.guiding_on:
            self.guiding_on = False
            self.ToggleGuidingButton.SetLabel('Start Guiding')
            self.ToggleCameraButton.Enable()
        else:
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
        
    def UpdateImageDisplay(self, image):
        wd, hd = self.ImageDisplay.Size
        wi, hi = image.shape
        # scale image from 5th to 100th percentile
        imin, imax = np.percentile(image, (5.0, 100.0))
        image = ((image-imin)/(imax-imin) * 255).clip(0, 255)
        # convert to RGB (but still greyscale)
        image = image.T.astype('uint8')
        image = np.dstack((image, image, image))
        wxImg = wx.EmptyImage(wi, hi)
        wxImg.SetData( image.tostring() )
        # scale the image, preserving the aspect ratio
        ad = float(hd)/wd
        ai = float(hi)/wi
        if ad > ai:
            hi_new = hd
            wi_new = hd / ai
        else:
            wi_new = wd
            hi_new = wd * ai
        wxImg = wxImg.Scale(wi_new, hi_new)
        self.ImageDisplay.SetBitmap(wxImg.ConvertToBitmap())

    def OnImageReady(self, event):
        self.UpdateImageDisplay(event.image)
        if self.guiding_on:
            pass

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
