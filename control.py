#!/usr/bin/python
# -*- coding: utf-8 -*-

# control.py

mode = 'live'

import wx
from datetime import datetime, timedelta
import time
import os.path
import numpy as np
import pyfits
import urlparse
from astropy.vo.samp import SAMPIntegratedClient
if mode is not None:
    # http://www.ascom-standards.org/Help/Developer/html/N_ASCOM_DeviceInterface.htm
    import win32com.client

class Control(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(Control, self).__init__(*args, **kwargs)
        self.tel = None
        self.cam = None
        self.bias = None
        self.flat = None
        self.default_exptime = 1.0
        self.default_numexp = 1
        self.min_nbias = 3
        self.min_nflat = 3
        self.max_ncontinuous = 100
        self.flat_offset = (10.0, 10.0)
        self.readout_time = 3.0
        self.InitUI()
        self.InitSAMP()
        self.SetupDS9()
        if mode is not None:
            self.InitTelescope()
            self.InitCamera()
        self.images_root_path = "C:/Users/LabUser/Pictures/Telescope/"

    def InitTelescope(self):
        if mode == 'sim':
            self.tel = win32com.client.Dispatch("ASCOM.Simulator.Telescope")
        elif mode == 'live':
            self.tel = win32com.client.Dispatch("ASCOM.Celestron.Telescope")
        if not self.tel.Connected:
            self.tel.Connected = True
        if self.tel.Connected:
            self.Log("Connected to telescope")
        else:
            self.Log("Unable to connect to telescope")
            self.tel = None
            
        if self.tel is not None:
            self.Log("Telescope time is {}".format(self.tel.UTCDate))
            if not self.tel.Tracking:
                self.tel.Tracking = True
            if self.tel.Tracking:
                self.Log("Telescope tracking")
            else:
                self.Log("Unable to start telescope tracking")

    def InitCamera(self):
        if mode == 'sim':
            self.cam = win32com.client.Dispatch("ASCOM.Simulator.Camera")
        elif mode == 'live':
            self.cam = win32com.client.Dispatch("ASCOM.SXMain0.Camera")
        if not self.cam.Connected:
            self.cam.Connected = True
        if self.cam.Connected:
            self.Log("Connected to camera")
        else:
            self.Log("Unable to connect to camera")
        
    def InitSAMP(self):
        self.Log('Attempting to connect to SAMP hub')
        try:
            self.samp_client = SAMPIntegratedClient()
            self.samp_client.connect()
        except Exception as detail:
            self.samp_client = None
            self.Log('Connection to SAMP hub failed:\n{}'.format(detail))
        else:
            self.Log('Connected to SAMP hub')

    def SetupDS9(self):
        if self.samp_client is not None:
            self.DS9Command('frame delete all')
            self.DS9Command('tile')
            self.DS9Command('frame new')
            self.DS9Command('frame new rgb')
        else:
            self.Log('No connection to DS9')

    def InitUI(self):
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.InitMenuBar()
        self.InitPanel()
        self.SetSize((600, 600))
        self.SetTitle('Control')
        self.Centre()
        self.Show(True)

    def InitMenuBar(self):
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        fitem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
        menubar.Append(fileMenu, '&File')
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.OnQuit, fitem)

    def InitPanel(self):
        panel = wx.Panel(self)
        MainBox = wx.BoxSizer(wx.HORIZONTAL)        
        sb = wx.StaticBox(panel)
        ButtonBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitButtons(panel, ButtonBox)
        MainBox.Add(ButtonBox, 0, flag=wx.EXPAND)
        sb = wx.StaticBox(panel)
        feedbackbox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        #sb = wx.StaticBox(panel, label="Image")
        #ImageBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        #feedbackbox.Add(ImageBox, 2, flag=wx.EXPAND)
        sb = wx.StaticBox(panel, label="Log")
        LogBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitLog(panel, LogBox)
        feedbackbox.Add(LogBox, 1, flag=wx.EXPAND)
        MainBox.Add(feedbackbox, 1, flag=wx.EXPAND)
        panel.SetSizer(MainBox)

    def InitButtons(self, panel, box):
        # flag to indicate if an image is being taken
        self.working = False
        # flag to indicate if we need to abort
        self.need_abort = False

        # maintain a list of all work buttons
        self.WorkButtons = []

        BiasButton = wx.Button(panel, label='Take bias images')
        BiasButton.Bind(wx.EVT_BUTTON, self.TakeBias, BiasButton)
        self.WorkButtons.append(BiasButton)
        BiasButton.SetToolTip(wx.ToolTip(
            'Take a set of bias images and store a master bias'))
        box.Add(BiasButton, flag=wx.EXPAND|wx.ALL, border=10)

        FlatButton = wx.Button(panel, label='Take flat images')
        FlatButton.Bind(wx.EVT_BUTTON, self.TakeFlat, FlatButton)
        self.WorkButtons.append(FlatButton)
        FlatButton.SetToolTip(wx.ToolTip(
            'Take test images to determine optimum exposure time, then '
            'take a set of flat images and store a master flat'))
        box.Add(FlatButton, flag=wx.EXPAND|wx.ALL, border=10)

        AcquisitionButton = wx.Button(panel, label='Take acquisition image')
        AcquisitionButton.Bind(wx.EVT_BUTTON, self.TakeAcquisition, AcquisitionButton)
        self.WorkButtons.append(AcquisitionButton)
        AcquisitionButton.SetToolTip(wx.ToolTip(
            'Take single image of specified exposure time'))
        box.Add(AcquisitionButton, flag=wx.EXPAND|wx.ALL, border=10)
        
        ScienceButton = wx.Button(panel, label='Take science images')
        ScienceButton.Bind(wx.EVT_BUTTON, self.TakeScience, ScienceButton)
        self.WorkButtons.append(ScienceButton)
        ScienceButton.SetToolTip(wx.ToolTip(
            'Take science images of specified exposure time and number'))
        box.Add(ScienceButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL, border=10)
        
        ContinuousButton = wx.Button(panel, label='Continuous images')
        ContinuousButton.Bind(wx.EVT_BUTTON, self.TakeContinuous, ContinuousButton)
        self.WorkButtons.append(ContinuousButton)
        ContinuousButton.SetToolTip(wx.ToolTip(
            'Take continuous images of specified exposure time'))
        box.Add(ContinuousButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL, border=10)

        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label='Exp.Time'),
                       flag=wx.RIGHT, border=5)        
        self.ExpTimeCtrl = wx.TextCtrl(panel, size=(50,-1),)
        self.ExpTimeCtrl.ChangeValue('{}'.format(self.default_exptime))
        self.ExpTimeCtrl.SetToolTip(wx.ToolTip(
            'Exposure time for science image, or initial exposure '
            'time to try for flats'))
        subBox.Add(self.ExpTimeCtrl)
        subBox.Add(wx.StaticText(panel, label='sec'),
                       flag=wx.LEFT, border=5)
        box.Add(subBox, flag=wx.EXPAND|wx.ALL, border=10)

        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label='Num.Exp.'),
                       flag=wx.RIGHT, border=5)        
        self.NumExpCtrl = wx.TextCtrl(panel, size=(50,-1),)
        self.NumExpCtrl.ChangeValue('{}'.format(self.default_numexp))
        self.NumExpCtrl.SetToolTip(wx.ToolTip(
            'Number of exposures (subject to minimum of ' +
            '{} for biases and '.format(self.min_nbias) +
            '{} for flats)'.format(self.min_nflat)))
        subBox.Add(self.NumExpCtrl)
        box.Add(subBox, flag=wx.EXPAND|wx.ALL, border=10)
        
        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL,
                border=10)
        
        self.AbortButton = wx.Button(panel, label='Abort')
        self.AbortButton.Bind(wx.EVT_BUTTON, self.Abort)
        self.AbortButton.SetToolTip(wx.ToolTip(
            'Abort the current operation as soon as possible'))
        self.AbortButton.Disable()
        box.Add(self.AbortButton, flag=wx.wx.EXPAND|wx.ALL,
                border=10)
        
    def InitLog(self, panel, box):
        self.logger = wx.TextCtrl(panel, size=(400,100),
                        style=wx.TE_MULTILINE | wx.TE_READONLY)
        box.Add(self.logger, 1, flag=wx.EXPAND)
        now = datetime.utcnow()
        timeStamp = now.strftime('%a %d %b %Y %H:%M:%S UT')
        self.logger.AppendText("Log started {}\n".format(timeStamp))

    def Log(self, text):
        # Work out if we're at the end of the file
        currentCaretPosition = self.logger.GetInsertionPoint()
        (currentSelectionStart, currentSelectionEnd) = self.logger.GetSelection()
        self.holdingBack = (currentSelectionEnd - currentSelectionStart) > 0
        # If some text is selected, then hold back advancing the log
        if self.holdingBack:
            self.logger.Freeze()
        now = datetime.utcnow()
        timeStamp = now.strftime('%H:%M:%S UT')
        self.logger.AppendText("{} : {}\n".format(timeStamp, text))
        if self.holdingBack:
            self.logger.SetInsertionPoint(currentCaretPosition)
            self.logger.SetSelection(currentSelectionStart, currentSelectionEnd)
            self.logger.Thaw()
        time.sleep(0.01)
        self.logger.Refresh()

    def EnableWorkButtons(self):
        for button in self.WorkButtons:
            button.Enable()

    def DisableWorkButtons(self):
        for button in self.WorkButtons:
            button.Disable()
        
    def CheckForAbort(self):
        self.logger.Refresh()
        wx.Yield()
        if self.need_abort:
            raise ControlAbortError()
        
    def StartWorking(self):
        if self.working:
            return False
        else:
            self.working = True
            self.AbortButton.Enable()
            self.DisableWorkButtons()
            return True

    def StopWorking(self):
        if self.working:
            self.working = False
            self.AbortButton.Disable()
            self.EnableWorkButtons()
            wx.Bell()
            return True
        else:
            return False

    def Abort(self, e):
        if self.working:
            self.AbortButton.Disable()
            self.Log('Trying to abort...')
            self.need_abort = True
            return True
        else:
            return False
        
    def OnQuit(self, e):
        if self.samp_client is not None:
            self.Log('Disconnecting from SAMP hub')
            try:
                self.samp_client.disconnect()
                self.tel.Connected = False
                self.cam.Connected = False
            except:
                pass
        self.Destroy()

    def TakeBias(self, e):
        nbias = self.GetNumExp()
        if nbias is None or nbias < self.min_nbias:
            nbias = self.min_nbias
        if self.StartWorking():
            self.Log('### Taking {} bias images...'.format(nbias))
            try:
                for i in range(nbias):
                    self.Log('Starting bias {}'.format(i+1))
                    self.CheckForAbort()
                    self.TakeImage(exptime=0)
                    self.Log('Taken bias {}'.format(i+1))
                    self.CheckForAbort()                    
            except ControlAbortError:
                self.need_abort = False
                self.Log('Bias images aborted')
            except Exception as detail:
                self.Log('Bias images error:\n{}'.format(detail))
            else:
                self.Log('Bias images done')
            self.StopWorking()

    def TakeFlat(self, e):
        nflat = self.GetNumExp()
        if nflat is None or nflat < self.min_nflat:
            nflat = self.min_nflat
        if self.StartWorking():
            self.Log('### Taking {} flat images...'.format(nflat))
            try:
                exptime = self.GetExpTime()
                exptime = self.GetFlatExpTime(exptime)
                if exptime is None:
                    self.Log('Flat images not obtained')
                else:
                    self.Log('Using exptime of {} sec'.format(exptime))
                    for i in range(nflat):
                        self.Log('Starting flat {}'.format(i+1))
                        self.CheckForAbort()
                        self.TakeImage(exptime)
                        self.OffsetTelescope(self.flat_offset)
                        self.Log('Taken flat {}'.format(i+1))
                        self.CheckForAbort()                    
            except ControlAbortError:
                self.need_abort = False
                self.Log('Flat images aborted')
            except Exception as detail:
                self.Log('Flat images error:\n{}'.format(detail))
            else:
                self.Log('Flat images done')
            self.StopWorking()

    def TakeScience(self, e):
        nexp = self.GetNumExp()
        exptime = self.GetExpTime()
        if nexp is None or exptime is None:
            self.Log('Science images not obtained')
        elif self.StartWorking():
            self.Log('### Taking {} science images...'.format(nexp))
            try:
                self.Log('Using exptime of {} sec'.format(exptime))
                for i in range(nexp):
                    self.Log('Starting exposure {}'.format(i+1))
                    self.CheckForAbort()
                    self.TakeImage(exptime)
                    self.Log('Taken exposure {}'.format(i+1))
                    self.CheckForAbort()                    
            except ControlAbortError:
                self.need_abort = False
                self.Log('Science images aborted')
            except Exception as detail:
                self.Log('Science images error:\n{}'.format(detail))
            else:
                self.Log('Science images done')
            self.StopWorking()

    def TakeContinuous(self, e):
        exptime = self.GetExpTime()
        if self.StartWorking():
            self.Log('### Taking continuous images...')
            try:
                self.Log('Using exptime of {} sec'.format(exptime))
                for i in range(self.max_ncontinuous):
                    self.CheckForAbort()
                    self.TakeImage(exptime)
                    # Should not save all these images
                    # Need to to display in ds9
                    # if reusing filename remember neeed to clobber
            except ControlAbortError:
                self.need_abort = False
                self.Log('Continuous done')
            except Exception as detail:
                self.Log('Continuous images error:\n{}'.format(detail))
            else:
                self.Log('Continuous timed out')
            self.StopWorking()

    def TakeAcquisition(self, e):
        exptime = self.GetExpTime()
        if self.StartWorking():
            self.Log('### Taking single acquisition image...')
            try:
                self.Log('Using exptime of {} sec'.format(exptime))
                self.CheckForAbort()
                self.TakeImage(exptime)
            except ControlAbortError:
                self.need_abort = False
                self.Log('Acquisition image aborted')
            except Exception as detail:
                self.Log('Acquisition image error:\n{}'.format(detail))
            else:
                self.Log('Acquisition image done')
            self.StopWorking()

    def GetExpTime(self):
        try:
            exptime = float(self.ExpTimeCtrl.GetValue())
        except ValueError:
            exptime = None
            self.Log('Exposure time invalid, setting to '
                    '{} sec'.format(self.default_exptime))
            self.ExpTimeCtrl.ChangeValue('{}'.format(self.default_exptime))
        return exptime

    def GetNumExp(self):
        try:
            numexp = int(self.NumExpCtrl.GetValue())
        except ValueError:
            numexp = None
            self.Log('Number of exposures invalid, '
                     'setting to {}'.format(self.default_numexp))
            self.NumExpCtrl.ChangeValue('{}'.format(self.default_numexp))
        return numexp
            
    def BiasSubtract(self):
        if self.bias is not None:
            self.image -= self.bias
            return True
        else:
            return False

    def Flatfield(self):
        if self.flat is not None:
            self.image /= self.flat
            return True
        else:
            return False

    def OffsetTelescope(self, offset_arcsec):
        dra, ddec = offset_arcsec
        if self.tel is not None:
            ra = self.tel.RightAscension + dra / (60*60*24)
            dec = self.tel.Declination + ddec / (60*60*360)
            self.tel.TargetRightAscension = ra
            self.tel.TargetDeclination = dec
            self.tel.SlewToTarget()
        else:
            self.Log('NOT offsetting telescope {}" RA, {}" Dec'%format(dra, ddec))
            
    def GetFlatExpTime(self, start_exptime=None,
                        min_exptime=0.001, max_exptime=60.0,
                        min_counts=25000.0, max_counts=35000.0):
        target_counts = (min_counts + max_counts)/2.0
        if start_exptime is None:
            start_exptime = self.default_exptime
        exptime = start_exptime
        while True:
            self.Log('Taking test flat of exptime '
                     '{} sec'.format(exptime))
            self.CheckForAbort()
            self.TakeImage(exptime)
            self.CheckForAbort()
            self.BiasSubtract()
            med_counts = np.median(self.image)
            self.Log('Median counts = {}'.format(med_counts))
            self.CheckForAbort()
            if med_counts > min_counts and med_counts < max_counts:
                break
            else:
                exptime *= target_counts/med_counts
            if exptime > max_exptime:
                self.Log('Required exposure time '
                         'longer than {} sec'.format(max_exptime))
                exptime = None
                break
            if exptime < min_exptime:
                self.Log('Required exposure time '
                         'shorter than {} sec'.format(min_exptime))
                exptime = None
                break
        return exptime

    def TakeImage(self, exptime):
        self.image_time = datetime.utcnow()
        if self.cam is not None:
            self.cam.StartExposure(exptime, True)
            time.sleep(exptime)
            time.sleep(self.readout_time)
            while not self.cam.ImageReady:
                self.CheckForAbort()
                time.sleep(1)
            self.image = np.array(self.cam.ImageArray)
        else:
            self.Log('NOT taking exposure of {} sec'.format(exptime))
            time.sleep(3)
            self.image = np.random.poisson(10000 * exptime, (100,100))
        self.SaveImage()
        self.DisplayImage()
        self.DeBayer()
        self.SaveRGBImages()
        self.DisplayRGBImage()

    def DisplayImage(self):
        if self.samp_client is not None:
            self.DS9LoadImage(self.images_path, self.filename, frame=1)
        
    def DisplayRGBImage(self):
        if self.samp_client is not None:
            self.DS9SelectFrame(2)
            for f in ('red', 'green', 'blue'):
                self.DS9Command('rgb {}'.format(f))
                self.DS9LoadImage(self.images_path, self.filters_filename[f[0]])
            self.DS9Command('rgb close')

    def SaveRGBImages(self, type='raw', name=None):
        self.SaveImage(type, name, filters=True)

    def SaveImage(self, type='raw', name=None, filters=False):
        if name is None:
            name = self.image_time.strftime('%Y-%m-%d_%H-%M-%S')
        night = self.image_time - timedelta(hours=12)
        path = self.image_time.strftime('%Y-%m-%d')
        self.images_path = os.path.join(self.images_root_path, path)
        if not os.path.exists(self.images_path):
            os.makedirs(self.images_path)
        header = None
        if filters is False:
            filename = name+'.fits'
            self.filename = filename
            fullfilename = os.path.join(self.images_path, filename)
            pyfits.writeto(fullfilename, self.image, header)
            self.Log('Saved {}'.format(filename))
        else:
            self.filters_filename = {}
            for f in self.filters:
                filename = name+'_'+f+'.fits'
                self.filters_filename[f] = filename
                fullfilename = os.path.join(self.images_path, filename)
                pyfits.writeto(fullfilename, self.filters[f], header)
                self.Log('Saved {}'.format(filename))

    def DeBayer(self):
        filters = []
        for i in (0, 1):
            for j in (0, 1):
                d = self.image[i::2,j::2]
                f = np.zeros(self.image.shape, self.image.dtype)
                for p in (0, 1):
                    for q in (0, 1):
                        f[p::2,q::2] = d
                filters.append(f)
        r, g1, g2, b = filters
        g = (g1+g2)/2.0
        self.filters = {'r': r, 'g': g, 'b': b}

    def DS9Command(self, cmd, url=None):
        params = {'cmd': cmd}
        if url is not None:
            params['url'] = url
        message = {'samp.mtype': 'ds9.set', 'samp.params': params}
        self.samp_client.notify_all(message)

    def DS9SelectFrame(self, frame):
        self.DS9Command('frame {}'.format(frame))

    def DS9LoadImage(self, path, filename, frame=None):
        if frame is not None:
            self.DS9SelectFrame(frame)
        url = urlparse.urljoin('file:', os.path.abspath(os.path.join(path, filename)))
        url = 'file:///'+os.path.abspath(os.path.join(path, filename)).replace('\\', '/')
        self.DS9Command('fits', url)
        
class ControlError(Exception):
    pass

class ControlAbortError(ControlError):
    def __init__(self, expr=None, msg=None):
        self.expr = expr
        self.msg = msg

            
def main():
    ex = wx.App()
    Control(None)
    ex.MainLoop()    

        
if __name__ == '__main__':
    main()
