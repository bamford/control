#!/usr/bin/python
# -*- coding: utf-8 -*-

# control.py

# simulate obtaining images for testing
simulate = True
debug = True

import wx
from datetime import datetime, timedelta
import time
import os.path
from glob import glob
import numpy as np
import scipy.stats
import pyfits
import urlparse
from astropy.vo.samp import SAMPIntegratedClient
if debug:
    import traceback
if not simulate:
    # http://www.ascom-standards.org/Help/Developer/html/N_ASCOM_DeviceInterface.htm
    import win32com.client

from guider import Guider

class Control(wx.Frame):

    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, title='Control',
                          size=(800, 600), **kwargs)
        self.__DoLayout()
        self.Log = self.panel.Log
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.Show(True)
        self.guider = Guider(self)
        self.guider.Hide()

    def __DoLayout(self):
        self.InitMenuBar()
        self.panel = ControlPanel(self)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(self.panel, 1, wx.EXPAND)
        self.SetSizer(sizer)

    def InitMenuBar(self):
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        fitem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
        menubar.Append(fileMenu, '&File')
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.OnQuit, fitem)

    def OnQuit(self, e):
        self.panel.OnQuit(None)
        self.guider.Destroy()
        self.Destroy()


class ControlPanel(wx.Panel):

    def __init__(self, *args, **kwargs):
        wx.Panel.__init__(self, *args, **kwargs)
        self.main = args[0]
        # configuration:
        self.default_exptime = 1.0
        self.default_numexp = 1
        self.min_nbias = 3
        self.min_nflat = 3
        self.max_ncontinuous = 100
        self.flat_offset = (10.0, 10.0)
        self.readout_time = 3.0
        self.images_root_path = "C:/Users/LabUser/Pictures/Telescope/"
        # special objects:
        self.tel = None
        self.cam = None
        self.bias = None
        self.flat = None
        # initialisations
        self.InitPanel()
        self.InitSAMP()
        self.InitDS9()
        self.InitTelescope()
        self.InitCamera()
        self.InitPaths()
        self.LoadCalibrations()

    def InitTelescope(self):
        if not simulate:
            self.tel = win32com.client.Dispatch("ASCOM.Celestron.Telescope")
        else:
            self.tel = None
        if self.tel is not None:
            if not self.tel.Connected:
                self.tel.Connected = True
            if self.tel.Connected:
                self.Log("Connected to telescope")
            else:
                self.Log("Unable to connect to telescope")
                self.tel = None
            self.Log("Telescope time is {}".format(self.tel.UTCDate))
            if not self.tel.Tracking:
                self.tel.Tracking = True
            if self.tel.Tracking:
                self.Log("Telescope tracking")
            else:
                self.Log("Unable to start telescope tracking")

    def InitCamera(self):
        if not simulate:
            self.cam = win32com.client.Dispatch("ASCOM.SXMain0.Camera")
        else:
            self.cam = None
        if not self.cam.Connected:
            self.cam.Connected = True
        if self.cam.Connected:
            self.cam.StartExposure(0, True) # discard first image
            # wait for camera to cool?
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

    def InitDS9(self):
        if self.samp_client is not None:
            self.Log('Attempting to set up DS9')
            self.DS9Command('frame delete all')
            self.DS9Command('tile')
            self.DS9Command('frame new')
            self.DS9Command('frame new rgb')
            self.DS9Command('rgb close')
        else:
            self.Log('No connection to DS9')

    def InitPaths(self):
        night = datetime.utcnow() - timedelta(hours=12)
        self.night = night.strftime('%Y-%m-%d')
        self.images_path = os.path.abspath(os.path.join(self.images_root_path,
                                                        self.night))
        if not os.path.exists(self.images_path):
            os.makedirs(self.images_path)
        self.Log('Storing images in {}'.format(self.images_path))
        
    def LoadCalibrations(self):
        # look for existing masterbias and masterflat images
        mb = glob(os.path.join(self.images_path, '*masterbias.fits'))
        if len(mb) > 0:
            mb.sort()
            mb = mb[-1]
            self.bias = np.asarray(pyfits.getdata(mb))
            self.Log('Loaded masterbias: {}'.format(os.path.basename(mb)))
        mf = glob(os.path.join(self.images_path, '*masterflat.fits'))
        if len(mf) > 0:
            mf.sort()
            mf = mf[-1]
            self.flat = np.asarray(pyfits.getdata(mf))
            self.Log('Loaded masterflat: {}'.format(os.path.basename(mf)))
            
    def InitPanel(self):
        MainBox = wx.BoxSizer(wx.HORIZONTAL)        
        sb = wx.StaticBox(self)
        ButtonBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitButtons(self, ButtonBox)
        MainBox.Add(ButtonBox, 0, flag=wx.EXPAND)
        sb = wx.StaticBox(self)
        feedbackbox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        #sb = wx.StaticBox(self, label="Image")
        #ImageBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        #feedbackbox.Add(ImageBox, 2, flag=wx.EXPAND)
        sb = wx.StaticBox(self, label="Log")
        LogBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitLog(self, LogBox)
        feedbackbox.Add(LogBox, 1, flag=wx.EXPAND)
        MainBox.Add(feedbackbox, 1, flag=wx.EXPAND)
        self.SetSizer(MainBox)

    def InitButtons(self, panel, box):
        # flag to indicate if an image is being taken
        self.working = False
        # flag to indicate if we need to abort
        self.need_abort = False

        # maintain a list of all work buttons
        self.WorkButtons = []

        BiasButton = wx.Button(panel, label='Take bias images')
        BiasButton.Bind(wx.EVT_BUTTON, self.TakeBias)
        self.WorkButtons.append(BiasButton)
        BiasButton.SetToolTip(wx.ToolTip(
            'Take a set of bias images and store a master bias'))
        box.Add(BiasButton, flag=wx.EXPAND|wx.ALL, border=10)

        FlatButton = wx.Button(panel, label='Take flat images')
        FlatButton.Bind(wx.EVT_BUTTON, self.TakeFlat)
        self.WorkButtons.append(FlatButton)
        FlatButton.SetToolTip(wx.ToolTip(
            'Take test images to determine optimum exposure time, then '
            'take a set of flat images and store a master flat'))
        box.Add(FlatButton, flag=wx.EXPAND|wx.ALL, border=10)

        AcquisitionButton = wx.Button(panel, label='Take acquisition image')
        AcquisitionButton.Bind(wx.EVT_BUTTON, self.TakeAcquisition)
        self.WorkButtons.append(AcquisitionButton)
        AcquisitionButton.SetToolTip(wx.ToolTip(
            'Take single image of specified exposure time'))
        box.Add(AcquisitionButton, flag=wx.EXPAND|wx.ALL, border=10)
        
        ScienceButton = wx.Button(panel, label='Take science images')
        ScienceButton.Bind(wx.EVT_BUTTON, self.TakeScience)
        self.WorkButtons.append(ScienceButton)
        ScienceButton.SetToolTip(wx.ToolTip(
            'Take science images of specified exposure time and number'))
        box.Add(ScienceButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL, border=10)
        
        ContinuousButton = wx.Button(panel, label='Continuous images')
        ContinuousButton.Bind(wx.EVT_BUTTON, self.TakeContinuous)
        self.WorkButtons.append(ContinuousButton)
        ContinuousButton.SetToolTip(wx.ToolTip(
            'Take continuous images of specified exposure time'))
        box.Add(ContinuousButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL, border=10)

        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label='Exp.Time'),
                       flag=wx.RIGHT, border=5)        
        self.ExpTimeCtrl = wx.TextCtrl(panel, size=(50,-1),)
        self.ExpTimeCtrl.ChangeValue('{:.3f}'.format(self.default_exptime))
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
        self.NumExpCtrl.ChangeValue('{:d}'.format(self.default_numexp))
        self.NumExpCtrl.SetToolTip(wx.ToolTip(
            'Number of exposures (subject to minimum of ' +
            '{:d} for biases and '.format(self.min_nbias) +
            '{:d} for flats)'.format(self.min_nflat)))
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
        
        box.Add(wx.StaticLine(panel), flag=wx.wx.EXPAND|wx.ALL,
                border=10)
        
        self.GuiderButton = wx.Button(panel, label='Show Guider')
        self.GuiderButton.Bind(wx.EVT_BUTTON, self.ToggleGuider)
        self.GuiderButton.SetToolTip(wx.ToolTip('Toggle guider window'))
        box.Add(self.GuiderButton, flag=wx.EXPAND|wx.ALL, border=10)

    def InitLog(self, panel, box):
        self.logger = wx.TextCtrl(panel, size=(600,100),
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

    def OnQuit(self, e):
        if self.samp_client is not None:
            self.Log('Disconnecting from SAMP hub')
            try:
                self.samp_client.disconnect()
                self.tel.Connected = False
                self.cam.Connected = False
            except:
                pass
        
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

    def TakeBias(self, e):
        # Popup to check cover on?
        nbias = self.GetNumExp()
        if nbias is None or nbias < self.min_nbias:
            nbias = self.min_nbias
        if self.StartWorking():
            self.Log('### Taking {:d} bias images...'.format(nbias))
            try:
                for i in range(nbias):
                    self.Log('Starting bias {:d}'.format(i+1))
                    self.CheckForAbort()
                    self.TakeImage(exptime=0)
                    self.Log('Taken bias {:d}'.format(i+1))
                    self.CheckForAbort()
                    self.SaveImage('bias')
                    self.SaveRGBImages('bias')
                    self.CheckForAbort()
                    if i==0:
                        bias_stack = np.zeros((nbias,)+self.image.shape,
                                              self.image.dtype)
                    bias_stack[i] = self.image
                    self.CheckForAbort()
                self.ProcessBias(bias_stack)
                self.CheckForAbort()
                self.bias = self.image
                self.SaveImage('masterbias')
                self.SaveRGBImages('masterbias')
            except ControlAbortError:
                self.need_abort = False
                self.Log('Bias images aborted')
            except Exception as detail:
                self.Log('Bias images error:\n{}'.format(detail))
            else:
                self.Log('Bias images done')
            self.StopWorking()

    def TakeFlat(self, e):
        # Popup to check ready?
        nflat = self.GetNumExp()
        if nflat is None or nflat < self.min_nflat:
            nflat = self.min_nflat
        if self.StartWorking():
            self.Log('### Taking {:d} flat images...'.format(nflat))
            try:
                exptime = self.GetExpTime()
                exptime = self.GetFlatExpTime(exptime)
                if exptime is None:
                    self.Log('Flat images not obtained')
                else:
                    self.Log('Using exptime of {:.3f} sec'.format(exptime))
                    for i in range(nflat):
                        self.Log('Starting flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        self.TakeImage(exptime)
                        self.Log('Taken flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        self.SaveImage('flat')
                        self.SaveRGBImages('flat')
                        self.Log('Taken flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        self.OffsetTelescope(self.flat_offset)
                        self.CheckForAbort()
                        if i==0:
                            flat_stack = np.zeros((nflat,)+self.image.shape,
                                                  np.float)
                        ok = self.BiasSubtract()
                        if not ok:
                            raise ControlError('Cannot create flat without bias')
                        flat_stack[i] = self.image
                        self.CheckForAbort()
                self.ProcessFlat(flat_stack)
                self.CheckForAbort()
                self.flat = self.image
                self.SaveImage('masterflat')
                self.SaveRGBImages('masterflat')
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
            self.Log('### Taking {:d} science images...'.format(nexp))
            try:
                self.Log('Using exptime of {:.3f} sec'.format(exptime))
                for i in range(nexp):
                    self.Log('Starting exposure {:d}'.format(i+1))
                    self.CheckForAbort()
                    self.TakeImage(exptime)
                    self.Log('Taken exposure {:d}'.format(i+1))
                    self.SaveImage()
                    self.Reduce()
                    self.SaveRGBImages()
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
                self.Log('Using exptime of {:.3f} sec'.format(exptime))
                for i in range(self.max_ncontinuous):
                    self.CheckForAbort()
                    self.TakeImage(exptime)
                    self.SaveImage(name='continuous')
                    self.Reduce()
                    self.SaveRGBImages(name='continuous')
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
                self.Log('Using exptime of {:.3f} sec'.format(exptime))
                self.CheckForAbort()
                self.TakeImage(exptime)
                self.Log('Acquisition exposure taken')
                self.SaveImage('acq')
                self.Reduce()
                self.GetAstrometry()
                self.SaveRGBImages('acq')
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
                    '{:.3f} sec'.format(self.default_exptime))
            self.ExpTimeCtrl.ChangeValue('{:.3f}'.format(self.default_exptime))
        return exptime

    def GetNumExp(self):
        try:
            numexp = int(self.NumExpCtrl.GetValue())
        except ValueError:
            numexp = None
            self.Log('Number of exposures invalid, '
                     'setting to {:d}'.format(self.default_numexp))
            self.NumExpCtrl.ChangeValue('{:d}'.format(self.default_numexp))
        return numexp

    def Reduce(self):
        self.BiasSubtract()
        self.Flatfield()

    def ProcessBias(self, stack):
        self.Log("Creating master bias")
        # Take the median through the stack to produce masterbias
        self.image = np.median(stack, axis=0)

    def ProcessFlat(self, stack):
        self.Log("Creating master flat")
        # Calculate image medians on a subsample to save time
        s = np.random.choice(np.product(stack[0].shape), 100000, replace=False)
        # Normalise each image in the stack
        for im in stack:
            im /= np.median(im.ravel()[s])
        # Take the median through the stack to produce masterflat
        self.image = np.median(stack, axis=0)
        
    def BiasSubtract(self):
        if self.bias is not None:
            self.image -= self.bias
            self.Log("Subtracting bias")
            return True
        else:
            self.Log("No bias correction")
            return False

    def Flatfield(self):
        if self.flat is not None:
            self.image /= self.flat
            self.Log("Flatfielding")
            return True
        else:
            self.Log("No flatfield correction")
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
            self.Log('NOT offsetting telescope {:.1f}" RA, {:.1f}" Dec'.format(dra, ddec))
            
    def GetFlatExpTime(self, start_exptime=None,
                        min_exptime=0.001, max_exptime=60.0,
                        min_counts=25000.0, max_counts=35000.0):
        target_counts = (min_counts + max_counts)/2.0
        if start_exptime is None:
            start_exptime = self.default_exptime
        exptime = start_exptime
        while True:
            self.Log('Taking test flat of exptime '
                     '{:.3f} sec'.format(exptime))
            self.CheckForAbort()
            self.TakeImage(exptime)
            self.CheckForAbort()
            self.BiasSubtract()
            med_counts = np.median(self.image)
            self.Log('Median counts = {:.1f}'.format(med_counts))
            self.CheckForAbort()
            if med_counts > min_counts and med_counts < max_counts:
                break
            else:
                exptime *= target_counts/med_counts
            if exptime > max_exptime:
                self.Log('Required exposure time '
                         'longer than {:.3f} sec'.format(max_exptime))
                exptime = None
                break
            if exptime < min_exptime:
                self.Log('Required exposure time '
                         'shorter than {:.3f} sec'.format(min_exptime))
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
            self.Log('NOT taking exposure of {:.3f} sec'.format(exptime))
            time.sleep(0.1)
            shape = (3040, 2024)
            if self.flat is None:
                self.image = np.random.poisson(10000 * exptime, size=shape)
                #self.image *= np.arange(shape[0])/(2.0*shape[0]) + 0.75
            else:
                size = 23
                g = scipy.stats.norm.pdf(np.arange(size), (size-1)/2.0, 4.0)
                star = np.dot(g[:, None], g[None, :])
                self.image = np.zeros(shape)
                for i in range(100):
                    x = np.random.choice(self.image.shape[0]-size)
                    y = np.random.choice(self.image.shape[1]-size)
                    flux = np.random.poisson(100) * 500
                    self.image[x:x+size,y:y+size] += star * flux
                #self.image *= np.arange(shape[0])/(2.0*shape[0]) + 0.75 
                self.image = np.random.poisson(self.image)
            self.image += np.random.normal(800, 20, size=shape)
        self.filters = None  # do not use filters until debayered

    def DisplayImage(self):
        if self.samp_client is not None:
            self.DS9LoadImage(self.images_path, self.filename, frame=1)
        
    def DisplayRGBImage(self):
        if self.samp_client is not None:
            self.DS9SelectFrame(2)
            for f in ('red', 'green', 'blue'):
                # Could this be all done in one SAMP command?
                self.DS9Command('rgb {}'.format(f))
                self.DS9LoadImage(self.images_path, self.filters_filename[f[0]])
            self.DS9Command('rgb close')

    def SaveRGBImages(self, imtype=None, name=None):
        self.DeBayer()
        self.SaveImage(imtype, name, filters=True)
        self.DisplayRGBImage()

    def SaveImage(self, imtype=None, name=None, filters=False):
        clobber = name is not None
        if name is None:
            name = self.image_time.strftime('%Y-%m-%d_%H-%M-%S')
        if imtype is not None:
            name += '_{}'.format(imtype)
        header = None
        if filters is False:
            filename = name+'.fits'
            self.filename = filename
            fullfilename = os.path.join(self.images_path, filename)
            pyfits.writeto(fullfilename, self.image, header,
                           clobber=clobber)
            self.Log('Saved {}'.format(filename))
        else:
            self.filters_filename = {}
            for f in self.filters:
                filename = name+'_'+f+'.fits'
                self.filters_filename[f] = filename
                fullfilename = os.path.join(self.images_path, filename)
                pyfits.writeto(fullfilename, self.filters[f], header,
                               clobber=clobber)
                self.Log('Saved {}'.format(filename))
        self.DisplayImage()

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
        self.astrometry = False

    def GetAstrometry(self):
        # placeholder method to be implemented
        # obtain astrometry via astrometry.net
        # (local or web-based), then set
        # self.astrometry = True
        pass

    def ToggleGuider(self, e):
        if self.main.guider.IsShown():
            self.main.guider.Hide()
            self.GuiderButton.SetLabel('Show Guider')
        else:
            self.main.guider.Show()
            self.GuiderButton.SetLabel('Hide Guider')

    def DS9Command(self, cmd, url=None, params=None):
        if params is None:
            params = {'cmd': cmd}
        else:
            params['cmd'] = cmd
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
        self.DS9Command('fits', params={'url': url, 'name': filename})
        
class ControlError(Exception):
    def __init__(self, expr=None, msg=None):
        if debug:
            print traceback.format_exc()

class ControlAbortError(ControlError):
    def __init__(self, expr=None, msg=None):
        self.expr = expr
        self.msg = msg


def main():
    app = wx.App(False)
    Control(None)
    app.MainLoop()

        
if __name__ == '__main__':
    main()
