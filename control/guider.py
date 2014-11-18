#!/usr/bin/python
# -*- coding: utf-8 -*-

# guider.py

#mode = 'live'
#mode = 'sim'
mode = None

import wx

class Guider(wx.Frame):
    
    def __init__(self, parent, *args, **kwargs):
        super(Guider, self).__init__(parent, *args, **kwargs)
        self.main = parent
        # config start
        self.comport = 4
        self.timeout = 0.1
        self.dark = None
        self.default_exptime = 1.0
        # config end
        self.camera = False
        self.cam = None
        self.InitFrame()

    def InitFrame(self):
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.InitPanel()
        self.SetSize((600, 600))
        self.SetTitle('Guider')
        self.Centre()
        self.Show(True)

    def OnQuit(self, e):
        self.main.ToggleGuider(e)
        
    def InitPanel(self):
        panel = wx.Panel(self)
        MainBox = wx.BoxSizer(wx.VERTICAL)        
        sb = wx.StaticBox(panel)
        ImageBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        MainBox.Add(ImageBox, 2, flag=wx.EXPAND)
        sb = wx.StaticBox(panel, label='Camera')
        CameraButtonBox = wx.StaticBoxSizer(sb, wx.HORIZONTAL)
        self.InitCameraButtons(panel, CameraButtonBox)
        MainBox.Add(CameraButtonBox, 0, flag=wx.EXPAND)
        GuidingButtonBox = wx.StaticBoxSizer(sb, wx.HORIZONTAL)
        self.InitGuidingButtons(panel, GuidingButtonBox)
        MainBox.Add(GuidingButtonBox, 0, flag=wx.EXPAND)
        panel.SetSizer(MainBox)

    def InitCameraButtons(self, panel, box):
        self.ToggleCameraButton = wx.Button(panel, label='Start Camera')
        self.ToggleCameraButton.Bind(wx.EVT_BUTTON, self.ToggleCamera)
        self.ToggleCameraButton.SetToolTip(wx.ToolTip(
            'Start/stop taking guider images'))
        box.Add(self.ToggleCameraButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticText(panel, label='Exp.Time'),
                       flag=wx.RIGHT, border=5)        
        self.ExpTimeCtrl = wx.TextCtrl(panel, size=(50,-1),)
        self.ExpTimeCtrl.ChangeValue('{:.3f}'.format(self.default_exptime))
        self.ExpTimeCtrl.SetToolTip(wx.ToolTip(
            'Exposure time for guider camera'))
        box.Add(self.ExpTimeCtrl)
        box.Add(wx.StaticText(panel, label='sec'),
                       flag=wx.LEFT, border=5)

    def InitGuidingButtons(self, panel, box):
        self.ToggleGuidingButton = wx.Button(panel, label='Start Guiding')
        self.ToggleGuidingButton.Bind(wx.EVT_BUTTON, self.ToggleGuiding)
        self.ToggleGuidingButton.SetToolTip(wx.ToolTip(
            'Start/stop guiding images'))
        box.Add(self.ToggleGuidingButton, flag=wx.EXPAND|wx.ALL, border=10)

        self.logger = wx.TextCtrl(panel, size=(300,50),
                        style=wx.TE_MULTILINE | wx.TE_READONLY)
        box.Add(self.logger, 1, flag=wx.EXPAND)

    def InitCamera(self):
        if mode == 'sim':
            self.cam = win32com.client.Dispatch("ASCOM.Simulator.Camera")
        elif mode == 'live':
            self.cam = win32com.client.Dispatch("ASCOM.SXGuider0.Camera")
        if not self.cam.Connected:
            self.cam.Connected = True
        if self.cam.Connected:
            self.cam.StartExposure(0, True) # discard first image
            # wait for camera to cool?
            self.Log("Connected to camera")
        else:
            self.Log("Unable to connect to camera")

    def ToggleCamera(self, e):
        if self.camera:
            self.StopCamera()
            self.ToggleCameraButton.SetLabel('Start Camera') 
        else:
            self.StartCamera()
            self.ToggleCameraButton.SetLabel('Stop Camera') 

    def ToggleGuiding(self, e):
        if self.guiding:
            self.guiding = True
            self.ToggleGuidingButton.SetLabel('Start Guiding') 
        else:
            self.guiding = False
            self.ToggleGuidingButton.SetLabel('Stop Guiding') 
            
    def StartCamera(self):
        if self.cam is not None:
            self.main.Log('Guider camera not found')
        exptime = self.GetExpTime()
        self.camera = True
        self.TakeImages(exptime)

    def UpdateImageDisplay(self, image_read):
        pass
            
    def StopCamera(self):
        self.camera = False

    def TakeImages(self, exptime, image_read, image_write):
        image_read.close()
        while self.camera:
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
                flux = 500.0
                g = scipy.stats.norm.pdf(np.arange(size), (size-1)/2.0, 4.0)
                star = np.dot(g[:, None], g[None, :])
                image = np.zeros(shape)
                x = self.image.shape[0]//2 - size//2
                y = self.image.shape[1]//2 - size//2
                image[x:x+size//2,y:y+size//2] += star * flux
                image = np.random.poisson(self.image)
                image += np.random.normal(800, 20, size=shape)
            if self.guiding:
                pass
            image_write.send(image)

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
        self.logger.AppendText(text)
        time.sleep(0.01)
        self.logger.Refresh()
