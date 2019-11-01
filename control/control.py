
#!/usr/bin/python
# -*- coding: utf-8 -*-

# control.py

from __future__ import print_function
import wx
import threading
from datetime import datetime, timedelta
import time
from Queue import Queue
import os.path
from glob import glob
import numpy as np
from scipy.ndimage.interpolation import shift
import astropy.coordinates as coord
import astropy.units as u
import astropy.io.fits as pyfits
from astropy.samp import SAMPIntegratedClient
import urlparse
import sys
import traceback
import win32api
import ntsecuritycon, win32security
from RGBImage import RGBImage

# simulate obtaining images for testing
simulate = False
test_image = False
debug = False
enable_guider = False
enable_windowing = False

if not simulate:
    # http://www.ascom-standards.org/Help/Developer/html/N_ASCOM_DeviceInterface.htm
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print('Windows COM modules not found.  Falling back to simulate mode.')
        simulate = True

from guider import Guider
from camera import TakeMainImageThread, EVT_IMAGEREADY_MAIN
from solver import SolverThread, EVT_SOLUTIONREADY
from logevent import EVT_LOG

class Control(wx.Frame):

    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, title='Control',
                          size=(800, 800), **kwargs)
        self.__DoLayout()
        self.Log = self.panel.Log
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.Bind(EVT_LOG, self.panel.OnLog)
        self.Bind(EVT_IMAGEREADY_MAIN, self.panel.OnImageReady)
        self.Bind(EVT_SOLUTIONREADY, self.panel.OnSolutionReady)
        self.Show(True)
        if enable_guider:
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
        if enable_guider:
            self.guider.OnExit(None)
        time.sleep(1)
        self.panel.OnQuit(None)
        wx.CallLater(1000, self.Destroy)
        e.Skip()

class ControlPanel(wx.Panel):

    def __init__(self, *args, **kwargs):
        wx.Panel.__init__(self, *args, **kwargs)
        self.main = args[0]
        # configuration:
        self.default_exptime = 1.0
        self.default_numexp = 1
        self.min_nbias = 5
        self.min_nflat = 3
        self.min_ndark = 5
        self.min_darktime = 5.0
        self.max_ncontinuous = 10000
        # do not subtract more dark than this
        # (to avoid oversubtracting saturated hot pixels):
        self.maxdark = 22500
        self.flat_offset = (10.0, 10.0)
        self.readout_time = 3.0

        self.images_root_path = "C:/Users/Labuser/The University of Nottingham/Physics Observatory - control/"
        # special objects:
        self.tel = None
        self.bias = None
        self.dark = None
        self.flat = None
        self.samp_client = None
        self.ast_position = None
        self.tel_position = None
        self.wcs = None
        self.image_time = None
        self.image_exptime = None
        self.image_tel_position = None
        self.last_telescope_move = datetime.utcnow()
        # initialisations
        self.InitPaths()
        self.InitPanel()
        wx.GetApp().Yield()
        wx.CallAfter(self.InitSAMP)
        wx.CallAfter(self.InitDS9)
        wx.CallAfter(self.InitTelescope)
        wx.CallAfter(self.InitCamera)
        wx.CallAfter(self.InitSolver)
        wx.CallAfter(self.LoadCalibrations)
        self.UpdateInfoTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.UpdateInfo, self.UpdateInfoTimer)
        self.UpdateInfoTimer.Start(1000) # 1 second interval

    def InitCamera(self):
        self.stop_camera = threading.Event()
        self.take_image = threading.Event()
        self.ImageTaker = TakeMainImageThread(self, self.stop_camera,
                                              self.take_image, 0.0)

    def StopCamera(self):
        self.stop_camera.set()
        self.take_image.clear()
        time.sleep(1)

    def InitTelescope(self):
        if not simulate:
            # Only in a thread:
            # win32com.client.pythoncom.CoInitialize()
            self.tel = win32com.client.Dispatch("ASCOM.Celestron.Telescope")
        else:
            self.tel = None
        if self.tel is not None:
            if not self.tel.Connected:
                try:
                    self.tel.Connected = True
                except pythoncom.com_error as error:
                    pass
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
        if self.tel is not None:
            now = self.tel.UTCDate
            now = datetime(now.year, now.month, now.day,
                           now.hour, now.minute, now.second,
                           now.msec * 1000)
            time_offset = abs(now - datetime.utcnow())
            if time_offset > timedelta(seconds=1):
                self.Log("Warning: PC and telescope times do not agree!")

    def InitSAMP(self):
        try:
            if self.samp_client is None:
                self.Log('Attempting to connect to SAMP hub')
                self.samp_client = SAMPIntegratedClient()
                self.samp_client.connect()
                self.Log('Connected to SAMP hub')
            else:
                self.samp_client.ping()
        except Exception as detail:
            self.samp_client = None
            self.Log('Connection to SAMP hub failed:\n{}'.format(detail))
            self.Log('Are TOPCAT and DS9 open? Is DS9 connected to SAMP?')

    def InitDS9(self, e=None):
        self.InitSAMP()
        if self.samp_client is not None:
            self.Log('Attempting to set up DS9')
            self.DS9Command('frame delete all')
            self.DS9Command('tile')
            self.DS9Command('frame new')
            self.DS9SelectFrame(1)
            self.DS9Command('zoom to 0.25')
            self.DS9Command('frame new rgb')
            self.DS9SelectFrame(2)
            self.DS9Command('zoom to 0.5')
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

    def InitSolver(self):
        self.wcs = None
        self.solver = Queue()
        self.SolverThread = SolverThread(self, self.solver,
                                directory='C:/Users/Labuser/solve')

    def StopSolver(self):
        self.solver.set(None)
        time.sleep(0.1)

    def LoadCalibrations(self, fallback=False):
        # look for existing masterbias, masterdark and masterflat images
        if fallback:
            path = os.path.join(self.images_root_path, '*')
        else:
            path = self.images_path
        if self.bias is None:
            mb = glob(os.path.join(path, '*masterbias.fits'))
            if len(mb) > 0:
                mb.sort()
                mb = mb[-1]
                self.bias = np.asarray(pyfits.getdata(mb))
                if fallback:
                    self.Log('Loaded OLD masterbias: {}'.format(os.path.basename(mb)))
                else:
                    self.Log('Loaded masterbias: {}'.format(os.path.basename(mb)))
        if self.dark is None:
            md = glob(os.path.join(path, '*masterdark.fits'))
            if len(md) > 0:
                md.sort()
                md = md[-1]
                self.dark = np.asarray(pyfits.getdata(md))
                if fallback:
                    self.Log('Loaded OLD masterdark: {}'.format(os.path.basename(md)))
                else:
                    self.Log('Loaded masterdark: {}'.format(os.path.basename(md)))
        if self.flat is None:
            mf = glob(os.path.join(path, '*masterflat.fits'))
            if len(mf) > 0:
                mf.sort()
                mf = mf[-1]
                self.flat = np.asarray(pyfits.getdata(mf))
                if fallback:
                    self.Log('Loaded OLD masterflat: {}'.format(os.path.basename(mf)))
                else:
                    self.Log('Loaded masterflat: {}'.format(os.path.basename(mf)))
        if not fallback:
            self.LoadCalibrations(fallback=True)

    def InitPanel(self):
        MainBox = wx.BoxSizer(wx.HORIZONTAL)
        sb = wx.StaticBox(self)
        ButtonBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitButtons(self, ButtonBox)
        MainBox.Add(ButtonBox, 0, flag=wx.EXPAND)
        sb = wx.StaticBox(self)
        feedbackbox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        sb = wx.StaticBox(self)
        InfoBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitInfo(self, InfoBox)
        feedbackbox.Add(InfoBox, 1, flag=wx.EXPAND)
        #sb = wx.StaticBox(self, label="Image")
        #ImageBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        #feedbackbox.Add(ImageBox, 2, flag=wx.EXPAND)
        sb = wx.StaticBox(self)
        LogBox = wx.StaticBoxSizer(sb, wx.VERTICAL)
        self.InitLog(self, LogBox)
        feedbackbox.Add(LogBox, 2, flag=wx.EXPAND)
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

        DarkButton = wx.Button(panel, label='Take dark images')
        DarkButton.Bind(wx.EVT_BUTTON, self.TakeDark)
        self.WorkButtons.append(DarkButton)
        DarkButton.SetToolTip(wx.ToolTip(
            'Take a set of dark images and store a master dark'))
        box.Add(DarkButton, flag=wx.EXPAND|wx.ALL, border=10)

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

        box.Add(wx.StaticLine(panel), flag=wx.EXPAND|wx.ALL, border=10)

        ContinuousButton = wx.Button(panel, label='Continuous images')
        ContinuousButton.Bind(wx.EVT_BUTTON, self.TakeContinuous)
        self.WorkButtons.append(ContinuousButton)
        ContinuousButton.SetToolTip(wx.ToolTip(
            'Take continuous images of specified exposure time'))
        box.Add(ContinuousButton, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.EXPAND|wx.ALL, border=10)

        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label='Exp.Time'),
                       flag=wx.RIGHT, border=5)
        self.ExpTimeCtrl = wx.TextCtrl(panel, size=(50,-1))
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
        self.NumExpCtrl = wx.TextCtrl(panel, size=(50,-1))
        self.NumExpCtrl.ChangeValue('{:d}'.format(self.default_numexp))
        self.NumExpCtrl.SetToolTip(wx.ToolTip(
            'Number of exposures (subject to minima for calibrations)'))
        subBox.Add(self.NumExpCtrl)
        box.Add(subBox, flag=wx.EXPAND|wx.ALL, border=10)

        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label='Delay'),
                       flag=wx.RIGHT, border=5)
        self.DelayTimeCtrl = wx.TextCtrl(panel, size=(50,-1))
        self.DelayTimeCtrl.ChangeValue('{:.0f}'.format(0))
        self.DelayTimeCtrl.SetToolTip(wx.ToolTip(
            'Optional delay between exposures'))
        subBox.Add(self.DelayTimeCtrl)
        subBox.Add(wx.StaticText(panel, label='sec'),
                       flag=wx.LEFT, border=5)
        box.Add(subBox, flag=wx.EXPAND|wx.ALL, border=10)

        if enable_windowing:
            subBox = wx.BoxSizer(wx.HORIZONTAL)
            subBox.Add(wx.StaticText(panel, label='Windowing'),
                           flag=wx.RIGHT, border=5)
            self.WindowCtrl = wx.CheckBox(panel)
            self.WindowCtrl.SetValue(False)
            self.WindowCtrl.Bind(wx.EVT_CHECKBOX, self.UpdateWindowing)
            self.WindowCtrl.SetToolTip(wx.ToolTip(
                'Limit to central region (for focussing, etc.)'))
            subBox.Add(self.WindowCtrl)
            box.Add(subBox, flag=wx.EXPAND|wx.ALL, border=10)

        box.Add(wx.StaticLine(panel), flag=wx.EXPAND|wx.ALL,
                border=10)

        self.AbortButton = wx.Button(panel, label='Abort')
        self.AbortButton.Bind(wx.EVT_BUTTON, self.Abort)
        self.AbortButton.SetToolTip(wx.ToolTip(
            'Abort the current operation as soon as possible'))
        self.AbortButton.Disable()
        box.Add(self.AbortButton, flag=wx.EXPAND|wx.ALL,
                border=10)

        box.Add(wx.StaticLine(panel), flag=wx.EXPAND|wx.ALL,
                border=10)

        if enable_guider:
            self.GuiderButton = wx.Button(panel, label='Show Guider')
            self.GuiderButton.Bind(wx.EVT_BUTTON, self.ToggleGuider)
            self.GuiderButton.SetToolTip(wx.ToolTip('Toggle guider window'))
            box.Add(self.GuiderButton, flag=wx.EXPAND|wx.ALL, border=10)
            box.Add(wx.StaticLine(panel), flag=wx.EXPAND|wx.ALL,
                    border=10)

        self.ResetDS9Button = wx.Button(panel, label='Reset DS9')
        self.ResetDS9Button.Bind(wx.EVT_BUTTON, self.InitDS9)
        self.ResetDS9Button.SetToolTip(wx.ToolTip('Reset image '
                                                  'display software'))
        box.Add(self.ResetDS9Button, flag=wx.EXPAND|wx.ALL, border=10)

    def InitInfo(self, panel, box):
        # Times
        subBox = wx.BoxSizer(wx.HORIZONTAL)
        self.pc_time = wx.StaticText(panel, size=(150,-1))
        subBox.Add(self.pc_time)
        subBox.Add((20, -1))
        self.tel_time = wx.StaticText(panel)
        subBox.Add(self.tel_time)
        box.Add(subBox, 0)
        box.Add((-1, 10))
        # Telescope position
        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label="Telescope:", size=(100,-1)))
        subBox.Add((20, -1))
        self.tel_ra = wx.StaticText(panel, size=(150,-1))
        subBox.Add(self.tel_ra)
        subBox.Add((20, -1))
        self.tel_dec = wx.StaticText(panel, size=(150,-1))
        subBox.Add(self.tel_dec)
        box.Add(subBox, 0)
        box.Add((-1, 10))
        # Astrometry position
        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label="Astrometry:", size=(100,-1)))
        subBox.Add((20, -1))
        self.ast_ra = wx.StaticText(panel, size=(150,-1))
        subBox.Add(self.ast_ra)
        subBox.Add((20, -1))
        self.ast_dec = wx.StaticText(panel, size=(150,-1))
        subBox.Add(self.ast_dec)
        subBox.Add((20, -1))
        #self.SyncButton = wx.Button(panel, label='Sync and Offset')
        #self.SyncButton.Bind(wx.EVT_BUTTON,
        #                     self.SyncToAstrometryAndOffsetTelescope)
        #self.SyncButton.SetToolTip(wx.ToolTip(
        #    'Sync telescope position to astrometry and '
        #    'offset to original target position'))
        #self.SyncButton.Disable()
        #subBox.Add(self.SyncButton, flag=wx.wx.EXPAND|wx.ALL,
        #           border=0)
        box.Add(subBox, 0)
        box.Add((-1, 10))
        # Target entry
        subBox = wx.BoxSizer(wx.HORIZONTAL)
        subBox.Add(wx.StaticText(panel, label="Target:", size=(100,-1)))
        subBox.Add((20, -1))
        self.TargetRACtrl = wx.TextCtrl(panel, size=(150,-1))
        self.TargetRACtrl.ChangeValue('00h00m00s')
        self.TargetRACtrl.SetToolTip(wx.ToolTip(
            'Target RA in format 00h00m00s'))
        subBox.Add(self.TargetRACtrl)
        subBox.Add((20, -1))
        self.TargetDecCtrl = wx.TextCtrl(panel, size=(150,-1))
        self.TargetDecCtrl.ChangeValue('+90d00m00s')
        self.TargetDecCtrl.SetToolTip(wx.ToolTip(
            'Target Dec in format +00d00m00s'))
        subBox.Add(self.TargetDecCtrl)
        subBox.Add((20, -1))
        self.SlewButton = wx.Button(panel, label='Slew to Target')
        self.SlewButton.Bind(wx.EVT_BUTTON,
                             self.SlewTelescope)
        self.SlewButton.SetToolTip(wx.ToolTip(
            'Slew telescope to given target position'))
        self.SlewButton.Enable()
        self.WorkButtons.append(self.SlewButton)
        subBox.Add(self.SlewButton, flag=wx.EXPAND|wx.ALL,
                   border=0)
        box.Add(subBox, 0)

    def UpdateInfo(self, event):
        self.UpdateTime()
        self.UpdatePosition()
        self.UpdateAstrometry()

    def UpdateTime(self):
        now = datetime.utcnow()
        timeStamp = now.strftime('%H:%M:%S UT')
        self.pc_time.SetLabel('PC time:  {}'.format(timeStamp))
        if self.tel is not None:
            try:
                now = self.tel.UTCDate
                self.tel_time.SetLabel('Tel. time:  {}'.format(now))
            except:
                self.tel = None
                self.Log('Telescope disconnected')
        else:
            self.tel_time.SetLabel('Tel. time:  not available')

    def UpdatePosition(self):
        # TODO: check self.tel.EquatorialSystem
        if self.tel is not None:
            c = coord.SkyCoord(ra=self.tel.RightAscension,
                               dec=self.tel.Declination,
                               unit=(u.hour, u.degree), frame='icrs')
            if self.tel_position is not None:
                if c.separation(self.tel_position).arcsecond > 15:
                    self.last_telescope_move = datetime.utcnow()
            self.tel_position = c
            ra = c.ra.to_string(u.hour, precision=1, pad=True)
            dec = c.dec.to_string(u.degree, precision=1, pad=True, alwayssign=True)
            self.tel_ra.SetLabel('RA:  ' + ra)
            self.tel_dec.SetLabel('Dec:  '+ dec)
            self.SlewButton.Enable()
        else:
            self.tel_position = None
            self.tel_ra.SetLabel('RA:  not available')
            self.tel_dec.SetLabel('Dec:  not available')
            self.SlewButton.Disable()

    def UpdateAstrometry(self):
        if self.image_time is None or self.last_telescope_move > self.image_time:
            self.ast_position = None
            self.wcs = None
        if self.ast_position is not None:
            ra = self.ast_position.ra.to_string(u.hour, precision=1, pad=True)
            dec = self.ast_position.dec.to_string(u.degree, precision=1, pad=True,
                                                  alwayssign=True)
            self.ast_ra.SetLabel('RA:  ' + ra)
            self.ast_dec.SetLabel('Dec:  ' + dec)
            #self.SyncButton.Enable()
        else:
            self.ast_ra.SetLabel('None')
            self.ast_dec.SetLabel('')
            #self.SyncButton.Disable()


    def UpdateWindowing(self, e):
        if self.ImageTaker is not None:
            self.ImageTaker.SetWindowing(self.WindowCtrl.GetValue())

    def InitLog(self, panel, box):
        self.logger = wx.TextCtrl(panel, size=(600,100),
                        style=wx.TE_MULTILINE | wx.TE_READONLY)
        box.Add(self.logger, 1, flag=wx.EXPAND)
        now = datetime.utcnow()
        timeStamp = now.strftime('%a %d %b %Y %H:%M:%S UT')
        text = "Log started {}\n".format(timeStamp)
        text += 'Storing images in {}\n'.format(self.images_path)
        self.logger.AppendText(text)
        self.logfilename = os.path.abspath(os.path.join(self.images_path,
                                                    'log_' + self.night))
        self.logfile = file(self.logfilename, 'a')
        self.logfile.write(text)

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
        text = "{} : {}\n".format(timeStamp, text)
        self.logger.AppendText(text)
        try:
            self.logfile.write(text)
            self.logfile.flush()
        except ValueError:
            # no logfile to write to
            pass
        if self.holdingBack:
            self.logger.SetInsertionPoint(currentCaretPosition)
            self.logger.SetSelection(currentSelectionStart, currentSelectionEnd)
            self.logger.Thaw()
        time.sleep(0.01)
        self.logger.Refresh()

    def OnLog(self, event):
        self.Log(event.text)

    def OnQuit(self, e):
        try:
            self.StopCamera()
        except:
            pass
        try:
            self.StopSolver()
        except:
            pass
        if self.samp_client is not None:
            self.Log('Disconnecting from SAMP hub')
            self.samp_client.disconnect()
        try:
            self.tel.Connected = False
            # Only in a thread:
            # win32com.client.pythoncom.CoUninitialize() # tel
        except:
            pass
        self.UpdateInfoTimer.Stop()
        self.logfile.close()

    def EnableWorkButtons(self):
        for button in self.WorkButtons:
            button.Enable()
        #if self.ast_position is not None:
        #    self.SyncButton.Enable()
        #else:
        #    self.SyncButton.Disable()


    def DisableWorkButtons(self):
        for button in self.WorkButtons:
            button.Disable()
        #self.SyncButton.Disable()

    def CheckForAbort(self):
        self.logger.Refresh()
        #wx.GetApp().Yield()
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
            self.worker = None
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
            self.take_image.clear()  # stop current exposure
            try:
                self.worker.next()
            except StopIteration:
                pass
            return True
        else:
            return False

    def OnImageReady(self, event):
        # The way images are obtained is a bit clever/complicated...
        # Clicking a "Take XXXX" button runs the corresponding TakeXXXX
        # method, which (via TakeWorker) runs TakeXXXXWorker to create a
        # generator, which is assigned to an instance variable, self.worker.
        # TakeWorker then calls next() on this generator, which starts an
        # exposure via the ImageTaker thread, then yields. TakeWorker then
        # completes, returning control to the WX panel.
        # When the exposure is done and the new image is ready, an
        # ImageReadyEvent is posted, running OnImageReady.
        # This transfers the image and its time to instance variables, then
        # calls next() on self.worker to continue from where it left off.
        # If an abort is issued, then the current exposure is stopped,
        # self.worker.next() is called and the worker handles the abort.
        if self.worker is not None:
            self.image = event.image
            self.image_time = event.image_time
            self.image_exptime = event.image_exptime
            self.image_tel_position = self.tel_position
            try:
                self.worker.next()
            except StopIteration:
                pass

    def TakeWorker(self, worker):
        if not self.working:
            self.worker = worker()
            self.worker.next()

    def TakeBias(self, e):
        if self.CheckReadyForBias():
            self.TakeWorker(self.TakeBiasWorker)

    def TakeBiasWorker(self):
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
                    yield
                    self.CheckForAbort()
                    self.Log('Taken bias {:d}'.format(i+1))
                    self.SaveImage('bias')
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
            except ControlAbortError:
                self.need_abort = False
                self.Log('Bias images aborted')
            except Exception as detail:
                self.Log('Bias images error:\n{}'.format(detail))
                traceback.print_exc()
            else:
                self.Log('Bias images done')
            self.StopWorking()

    def TakeDark(self, e):
        if self.CheckReadyForBias():
            self.TakeWorker(self.TakeDarkWorker)

    def TakeDarkWorker(self):
        ndark = self.GetNumExp()
        darktime = self.GetExpTime()
        if (ndark is None) or ((ndark < self.min_ndark) and (darktime < self.min_darktime)):
            ndark = self.min_ndark
        if darktime is None or darktime < self.min_darktime:
            darktime = self.min_darktime
        if self.StartWorking():
            self.Log('### Taking {:d} dark images of {:.3f} sec...'.format(ndark, darktime))
            try:
                for i in range(ndark):
                    self.Log('Starting dark {:d}'.format(i+1))
                    self.CheckForAbort()
                    self.TakeImage(darktime)
                    yield
                    self.CheckForAbort()
                    self.Log('Taken dark {:d}'.format(i+1))
                    self.SaveImage('dark')
                    self.CheckForAbort()
                    if i==0:
                        dark_stack = np.zeros((ndark,)+self.image.shape,
                                              self.image.dtype)
                    ok = self.BiasSubtract()
                    if not ok:
                        raise ControlError('Cannot create dark without bias')
                    dark_stack[i] = self.image
                    self.CheckForAbort()
                self.ProcessDark(dark_stack, darktime)
                self.dark = self.image
                self.SaveImage('masterdark')
            except ControlAbortError:
                self.need_abort = False
                self.Log('Dark images aborted')
            except Exception as detail:
                self.Log('Dark images error:\n{}'.format(detail))
                traceback.print_exc()
            else:
                self.Log('Dark images done')
            self.StopWorking()

    def TakeFlat(self, e):
        if self.CheckReadyForFlat():
            self.TakeWorker(self.TakeFlatWorker)

    def TakeFlatWorker(self):
        nflat = self.GetNumExp()
        if nflat is None or nflat < self.min_nflat:
            nflat = self.min_nflat
        if self.StartWorking():
            self.Log('### Taking {:d} flat images...'.format(nflat))
            try:
                startexptime = self.GetExpTime()
                GetFlatExpTime = self.GetFlatExpTime(startexptime)
                exptime = GetFlatExpTime.next()
                while exptime is None:
                    yield
                    exptime = GetFlatExpTime.next()
                if exptime < 0:
                    self.Log('Flat images not obtained')
                else:
                    self.Log('Using exptime of {:.3f} sec'.format(exptime))
                    for i in range(nflat):
                        self.Log('Starting flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        self.TakeImage(exptime)
                        yield
                        self.Log('Taken flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        self.SaveImage('flat')
                        self.Log('Taken flat {:d}'.format(i+1))
                        self.CheckForAbort()
                        #self.OffsetTelescope(self.flat_offset)
                        self.CheckForAbort()
                        if i==0:
                            flat_stack = np.zeros((nflat,)+self.image.shape,
                                                  np.float)
                        ok = self.BiasSubtract()
                        if not ok:
                            raise ControlError('Cannot create flat without bias')
                        self.DarkSubtract(exptime)
                        self.SaveRGBImages('flat')
                        self.DisplayRGBImage()
                        flat_stack[i] = self.image
                        self.CheckForAbort()
                    self.ProcessFlat(flat_stack)
                    self.CheckForAbort()
                    self.flat = self.image
                    self.SaveImage('masterflat')
                    self.SaveRGBImages('masterflat')
                    self.DisplayRGBImage()
            except ControlAbortError:
                self.need_abort = False
                self.Log('Flat images aborted')
            except Exception as detail:
                self.Log('Flat images error:\n{}'.format(detail))
                traceback.print_exc()
            else:
                self.Log('Flat images done')
            self.StopWorking()

    def CheckAdjustTime(self):
        dial = wx.MessageDialog(None,
                                'Adjust system time to telescope time?\n',
                                'System and telescope times do not match',
                                wx.OK | wx.CANCEL | wx.ICON_QUESTION)
        response = dial.ShowModal()
        return response == wx.ID_OK

    def CheckReadyForFlat(self):
        dial = wx.MessageDialog(None,
                                'Telescope pointing at twilight sky '
                                '/ illuminated dome?\n'
                                'Cover off?\n',
                                'Are you ready?',
                                wx.OK | wx.CANCEL | wx.ICON_QUESTION)
        response = dial.ShowModal()
        return response == wx.ID_OK

    def CheckReadyForBias(self):
        dial = wx.MessageDialog(None,
                                'Cover on?\n'
                                'Lights low?\n',
                                'Are you ready?',
                                wx.OK | wx.CANCEL | wx.ICON_QUESTION)
        response = dial.ShowModal()
        return response == wx.ID_OK

    def TakeScience(self, e):
        self.TakeWorker(self.TakeScienceWorker)

    def TakeScienceWorker(self):
        nexp = self.GetNumExp()
        exptime = self.GetExpTime()
        delaytime = self.GetDelayTime()
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
                    yield
                    self.CheckForAbort()
                    self.Log('Taken exposure {:d}'.format(i+1))
                    if test_image:
                        self.Log('Using test image')
                        fullfilename = os.path.join('C:/Users/Labuser/solve',
                                                    'test.fits')
                        self.image = pyfits.getdata(fullfilename)
                    self.SaveImage('sci')
                    self.Log('Reducing images')
                    self.Reduce(exptime)
                    try:
                        wx.Yield()
                        self.SaveRGBImages('sci', jpeg=True)
                        wx.Yield()
                        self.GetAstrometry()
                        wx.Yield()
                        self.DisplayRGBImage()
                    except Exception as detail:
                        tb = traceback.format_exc()
                        self.Log('Science images error:\n{}'.format(tb))
                        self.Log('Attempting to continue')
                    self.CheckForAbort()
                    if i < nexp - 1:
                        self.Delay(delaytime)
                    self.CheckForAbort()
            except ControlAbortError:
                self.need_abort = False
                self.Log('Science images aborted')
            except Exception as detail:
                self.Log('Science images error:\n{}'.format(detail))
                traceback.print_exc()
            else:
                self.Log('Science images done')
            self.StopWorking()

    def Delay(self, delaytime):
        if delaytime > 0:
            self.Log('### Waiting {:.1f} sec'.format(delaytime))
            delay = delaytime
            while delay > 0:
                self.CheckForAbort()
                wx.GetApp().Yield()
                time.sleep(0.1)
                delay -= 0.1

    def TakeContinuous(self, e):
        self.TakeWorker(self.TakeContinuousWorker)

    def TakeContinuousWorker(self):
        exptime = self.GetExpTime()
        if self.StartWorking():
            self.Log('### Taking continuous images...')
            try:
                self.Log('Using exptime of {:.3f} sec'.format(exptime))
                for i in range(self.max_ncontinuous):
                    self.CheckForAbort()
                    self.TakeImage(exptime)
                    yield
                    self.CheckForAbort()
                    self.SaveImage(name='continuous')
                    #self.Reduce(exptime)
                    #self.SaveRGBImages(name='continuous')
                    #self.DisplayRGBImage()
            except ControlAbortError:
                self.need_abort = False
                self.Log('Continuous done')
            except Exception as detail:
                self.Log('Continuous images error:\n{}'.format(detail))
                traceback.print_exc()
            else:
                self.Log('Continuous timed out')
            self.StopWorking()

    def TakeAcquisition(self, e):
        self.TakeWorker(self.TakeAcquisitionWorker)

    def TakeAcquisitionWorker(self):
        exptime = self.GetExpTime()
        if self.StartWorking():
            self.Log('### Taking single acquisition image...')
            try:
                self.Log('Using exptime of {:.3f} sec'.format(exptime))
                self.CheckForAbort()
                self.TakeImage(exptime)
                yield
                self.CheckForAbort()
                self.Log('Acquisition exposure taken')
                # TESTING!!!
                #fullfilename = os.path.join(self.images_root_path, 'solve',
                #                            'test.fits')
                #self.image = pyfits.getdata(fullfilename)
                self.SaveImage('acq')
                self.Log('Reducing images')
                self.Reduce(exptime)
                self.SaveRGBImages('acq')
                self.GetAstrometry()
                self.DisplayRGBImage()
            except ControlAbortError:
                self.need_abort = False
                self.Log('Acquisition image aborted')
            except Exception as detail:
                self.Log('Acquisition image error:\n{}'.format(detail))
                traceback.print_exc()
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

    def GetDelayTime(self):
        try:
            delaytime = float(self.DelayTimeCtrl.GetValue())
            if delaytime < 0:
                raise ValueError
        except ValueError:
            delaytime = 0
            self.Log('Delay time invalid, setting to zero')
            self.DelayTimeCtrl.ChangeValue('{:.0f}'.format(delaytime))
        return delaytime

    def GetNumExp(self):
        try:
            numexp = int(self.NumExpCtrl.GetValue())
        except ValueError:
            numexp = None
            self.Log('Number of exposures invalid, '
                     'setting to {:d}'.format(self.default_numexp))
            self.NumExpCtrl.ChangeValue('{:d}'.format(self.default_numexp))
        return numexp

    def Reduce(self, exptime):
        self.BiasSubtract()
        self.DarkSubtract(exptime)
        self.Flatfield()

    def ProcessBias(self, stack):
        self.Log("Creating master bias")
        # Take the median through the stack to produce masterbias
        self.image = np.median(stack, axis=0)

    def ProcessDark(self, stack, darktime):
        self.Log("Creating master dark")
        # Take the median through the  stack and divide by
        #  exposure time to produce dark
        # (counts per second assuming constant linear response)
        dark_base = np.median(stack, axis=0)
        dark_base /= darktime
        self.image = dark_base

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
            self.image = self.image - self.bias
            self.Log("Subtracting bias")
            return True
        else:
            self.Log("No bias correction")
            return False

    def DarkSubtract(self, exptime):
        if self.dark is not None:
            dark = self.dark * exptime
            dark[dark > self.maxdark] = self.maxdark
            self.image = self.image - dark
            self.Log("Subtracting dark")
            return True
        else:
            self.Log("No dark correction")
            return False

    def Flatfield(self):
        if self.flat is not None:
            self.image = self.image / self.flat
            self.Log("Flatfielding")
            return True
        else:
            self.Log("No flatfield correction")
            return False

    def SlewTelescope(self, event):
        try:
            ra_str = self.TargetRACtrl.GetValue()
            dec_str = self.TargetDecCtrl.GetValue()
            target = coord.SkyCoord(ra_str, dec_str)
        except:
            target = None
            self.Log('Target coordinates not recognised')
            traceback.print_exc()
        if target is not None:
            ra_str = target.ra.to_string(u.hour, precision=1, pad=True)
            dec_str = target.dec.to_string(u.deg, precision=1, pad=True,
                                           alwayssign=True)
            self.TargetRACtrl.ChangeValue(ra_str)
            self.TargetDecCtrl.ChangeValue(dec_str)
            self.Log('Slewing to {} {}'.format(ra_str, dec_str))
            self.tel.TargetRightAscension = target.ra.hour
            self.tel.TargetDeclination = target.dec.deg
            self.ast_position = None
            try:
                self.tel.SlewToTarget()
            except:
                self.Log('Slew failed')
                traceback.print_exc()
            else:
                self.Log('Slew complete')

    def SyncToAstrometryAndOffsetTelescope(self, event):
        if self.tel is not None and self.ast_position is not None:
            ra = self.tel.RightAscension
            dec = self.tel.Declination
            sep = self.tel_position.separation(self.ast_position)
            if (sep.degree < 5 or self.CheckSync()):
                dra, ddec = self.tel_position.spherical_offsets_to(self.ast_position)
                self.Log('Offsetting telescope to astrometry')
                self.OffsetTelescope((dra.arcsec, ddec.arcsec))
                self.Log('Telescope offset to astrometry')
        else:
            self.Log('NOT syncing telescope to astrometry')

    def CheckSync(self):
        dial = wx.MessageDialog(None,
                                'Requested offset > 5 deg.\n'
                                'Telescope may need aligning.\n'
                                'Are you sure you want to offset?',
                                wx.OK | wx.CANCEL | wx.ICON_QUESTION)
        response = dial.ShowModal()
        return response == wx.ID_OK

    def OffsetTelescope(self, offset_arcsec):
        dra, ddec = offset_arcsec
        if self.tel is not None:
            # why did pulse guiding not work?
            # this is severely flawed - currently disabled
            # could do this more correctly with astropy
            ra = self.tel.RightAscension + dra / (60*60*24)
            if ra > 24:
                ra -= 24
            elif ra < 0:
                ra += 24
            dec = self.tel.Declination + ddec / (60*60*360)
            if dec > 90:
                dec = 90.0
            elif dec < -90:
                dec = -90.0
            self.tel.SlewToCoordinates(ra, dec)
        else:
            self.Log('NOT offsetting telescope {:.1f}" RA, {:.1f}" Dec'.format(dra, ddec))

    def GetFlatExpTime(self, start_exptime=None,
                       min_exptime=0.1, max_exptime=120.0,
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
            yield
            self.CheckForAbort()
            self.BiasSubtract()
            self.SaveImage(name='flattest')
            self.DisplayImage()
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
                exptime = -1
            if exptime < min_exptime:
                self.Log('Required exposure time '
                         'shorter than {:.3f} sec'.format(min_exptime))
                exptime = -1
            break  # only try one test image
        yield exptime

    def TakeImage(self, exptime):
        self.image = None
        self.filters = None
        self.filters_interp = None
        if not self.ImageTaker.isAlive():
            self.Log("Restarting camera")
            self.StopCamera()
            self.ImageTaker = TakeMainImageThread(self, self.stop_camera,
                                                  self.take_image, 0.0)
        self.ImageTaker.SetExpTime(exptime)
        self.take_image.set()

    def DisplayImage(self):
        self.InitSAMP()
        if self.samp_client is not None:
            self.DS9LoadImage(self.images_path, self.filename, frame=1)
        # when in continuous mode place a region in the centre to help
        # alignment, same size as optional windowing
        if 'continuous' in self.filename:
            self.DS9SelectFrame(1)
            nx = ny = 100
            cx = self.image.shape[1] // 2
            cy = self.image.shape[0] // 2
            self.DS9Command('regions command "box({},{},{},{},0)"'.format(cx, cy, nx, ny))

    def DisplayRGBImage(self):
        self.InitSAMP()
        if self.samp_client is not None:
            self.DS9SelectFrame(2)
            for f in ('red', 'green', 'blue'):
                # Could this be all done in one SAMP command?
                self.DS9Command('rgb {}'.format(f))
                self.DS9LoadImage(self.images_path, self.filters_filename[f[0]])
            self.DS9Command('rgb close')
            #self.DS9LoadRGBImage(self.images_path, self.rgb_filename, frame=2)

    def SaveRGBImages(self, imtype=None, name=None, jpeg=False):
        self.DeBayer()
        name = self.SaveImage(imtype, name, filters=True)
        if jpeg:
            self.SaveJpeg(imtype, name)

    def SaveImage(self, imtype=None, name=None, filters=False, filtersum=False):
        clobber = name is not None
        if name is None:
            name = self.image_time.strftime('%Y-%m-%d_%H-%M-%S')
        if imtype is not None:
            name += '_{}'.format(imtype)
        # Could use the current (approximate) wcs, but can lead to confusion
        #header = self.wcs if self.wcs is not None else None
        header = pyfits.Header()
        header['DATE-OBS'] = self.image_time.strftime('%Y-%m-%d')
        header['TIME-OBS'] = self.image_time.strftime('%H:%M:%S.%f')
        header['EXPTIME'] = (self.image_exptime, 'seconds')
        if ((self.image_tel_position is not None) and
            imtype not in ('bias', 'dark', 'flat')):
            header['RA'] = self.image_tel_position.ra.to_string(u.hour, sep=':', precision=1, pad=True)
            header['DEC'] = self.image_tel_position.dec.to_string(u.degree, sep=':', precision=1,
                                                                  pad=True, alwayssign=True)
        if imtype is not None:
            header['OBJECT'] = imtype
        if (filters or filtersum) is False:
            filename = name+'.fits'
            self.filename = filename
            fullfilename = os.path.join(self.images_path, filename)
            pyfits.writeto(fullfilename, self.image, header,
                           overwrite=clobber)
            self.Log('Saved {}'.format(filename))
        elif filtersum:
            filename = name+'.fits'
            fullfilename = os.path.join(self.images_path, filename)
            pyfits.writeto(fullfilename, np.sum(self.filters, 0), header,
                           overwrite=clobber)
        else:
            self.filters_filename = {}
            for i, f in enumerate('rgb'):
                filename = name+'_'+f+'.fits'
                self.filters_filename[f] = filename
                fullfilename = os.path.join(self.images_path, filename)
                pyfits.writeto(fullfilename, self.filters[i], header,
                               overwrite=clobber)
                self.Log('Saved {}'.format(filename))
            #self.rgb_filename = name+'_rgb.fits'
            #fullfilename = os.path.join(self.images_path, self.rgb_filename)
            #pyfits.writeto(fullfilename, self.filters[0], header,
            #               clobber=clobber)
            #pyfits.append(fullfilename, self.filters[1], header)
            #pyfits.append(fullfilename, self.filters[2], header)
            #self.Log('Saved {}'.format(self.rgb_filename))
        self.DisplayImage()
        return name

    def SaveJpeg(self, imtype=None, name=None):
        im = RGBImage(self.filters_interp[0],
                      self.filters_interp[1],
                      self.filters_interp[2],
                      process=True, desaturate=True,
                      scales=[1.0, 0.8, 1.0])
        filename = name+'.jpg'
        fullfilename = os.path.join(self.images_path, filename)
        im.save_as(fullfilename)

    def DeBayer(self):
        filters = []
        filters_interp = []
        for i in (0, 1):
            for j in (0, 1):
                f = self.image[i::2,j::2]
                fi = np.zeros(self.image.shape, self.image.dtype)
                for p in (0, 1):
                    for q in (0, 1):
                        fi[p::2,q::2] = shift(f,
                                              (0.5 * (i - p),
                                               0.5 * (j - q)),
                                              order=1)
                filters.append(f)
                filters_interp.append(fi)
        r, g1, g2, b = filters
        g = (g1+g2)/2.0
        self.filters = np.array([r, g, b])
        r, g1, g2, b = filters_interp
        g = (g1+g2)/2.0
        self.filters_interp = np.array([r, g, b])

    def GetAstrometry(self):
        self.Log('Attempting to determine astrometry')
        path = 'C:/Users/Labuser/solve'
        if not os.path.exists(path):
            os.makedirs(path)
        solvefilename = os.path.join(path, 'solve.fits')
        self.solver.put((self.filters, solvefilename,
                         self.image_time,
                         self.filters_filename.values(),
                         self.image_tel_position))

    def OnSolutionReady(self, event):
        if event.solution is not None:
            message = 'Astrometry for image taken {}:\n{}'
            message = message.format(event.image_time,
                                     event.solution)
            c = coord.SkyCoord(ra=event.solution.center.RA,
                               dec=event.solution.center.dec,
                               unit=(u.degree, u.degree), frame='icrs')
            self.ast_position = c
            wcs = pyfits.getheader(os.path.join('C:/Users/Labuser/solve', 'solve.wcs'))
            if self.last_telescope_move <= event.image_time:
                self.wcs = wcs
            else:
                self.wcs = None
        else:
            message = 'Astrometry for image taken {}: failed'
            message = message.format(event.image_time)
            wcs = None
        self.Log(message)
        self.UpdateFileWCS(event.filenames, wcs)
        if wcs is not None and self.image_time == event.image_time:
            # no other image taken in meantime
            self.DisplayRGBImage()

    def UpdateFileWCS(self, filenames, wcs):
        if wcs is not None:
            filenames = [os.path.join(self.images_path, f) for f in filenames]
            for fn in filenames:
                # in principle could tweak WCS for each filter here
                for attempt in range(3):
                    # try several times as might be being accessed by DS9
                    try:
                        with pyfits.open(fn, mode='update') as f:
                            f[0].header.update(wcs)
                    except WindowsError:
                        time.sleep(3)
                    else:
                        self.Log('Updated WCS of {}'.format(os.path.basename(fn)))
                        break

    def ToggleGuider(self, e):
        if self.main.guider.IsShown():
            self.main.guider.Hide()
            self.GuiderButton.SetLabel('Show Guider')
        else:
            self.main.guider.Show()
            self.GuiderButton.SetLabel('Hide Guider')

    def DS9Command(self, cmd, url=None, params=None):
        wx.GetApp().Yield()
        if params is None:
            params = {'cmd': cmd}
        else:
            params['cmd'] = cmd
        if url is not None:
            params['url'] = url
        message = {'samp.mtype': 'ds9.set', 'samp.params': params}
        if self.samp_client is not None:
            try:
                self.samp_client.notify_all(message)
            except:
                self.samp_client = None
        time.sleep(0.1)

    def DS9SelectFrame(self, frame):
        self.DS9Command('frame {}'.format(frame))

    def DS9LoadImage(self, path, filename, frame=None):
        if frame is not None:
            self.DS9SelectFrame(frame)
        url = urlparse.urljoin('file:', os.path.abspath(os.path.join(path, filename)))
        url = 'file:///'+os.path.abspath(os.path.join(path, filename)).replace('\\', '/')
        self.DS9Command('fits', params={'url': url, 'name': filename})

    def DS9LoadRGBImage(self, path, filename, frame=None):
        if frame is not None:
            self.DS9SelectFrame(frame)
        url = urlparse.urljoin('file:', os.path.abspath(os.path.join(path, filename)))
        url = 'file:///'+os.path.abspath(os.path.join(path, filename)).replace('\\', '/')
        self.DS9Command('rgbimage', params={'url': url, 'name': filename})


def AdjustPrivilege( priv ):
    flags = ntsecuritycon.TOKEN_ADJUST_PRIVILEGES | ntsecuritycon.TOKEN_QUERY
    htoken =  win32security.OpenProcessToken(win32api.GetCurrentProcess(), flags)
    id = win32security.LookupPrivilegeValue(None, priv)
    newPrivileges = [(id, ntsecuritycon.SE_PRIVILEGE_ENABLED)]
    win32security.AdjustTokenPrivileges(htoken, 0, newPrivileges)

class ControlError(Exception):
    def __init__(self, expr=None, msg=None):
        if debug:
            print(traceback.format_exc())

class ControlAbortError(ControlError):
    def __init__(self, expr=None, msg=None):
        self.expr = expr
        self.msg = msg

def excepthook(type, value, tb):
    message = 'Uncaught exception:\n'
    message += ''.join(traceback.format_exception(type, value, tb))
    message += '\nSorry, something has gone wrong.\n'
    message += 'If problems continue it is probably best to restart Control.'
    print(message)
    wx.MessageDialog(None, message)

def main():
    sys.excepthook = excepthook
    app = wx.App(False)
    Control(None)
    app.MainLoop()


if __name__ == '__main__':
    main()
