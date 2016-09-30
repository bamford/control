import wx
import threading
import time
from datetime import datetime
import numpy as np
from scipy.stats import norm

from logevent import *

# simulate obtaining images for testing
simulate = False

if not simulate:
    # http://www.ascom-standards.org/Help/Developer/html/N_ASCOM_DeviceInterface.htm
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        simulate = True

# Avoid multiple processes connecting to camera at same time
COMlock = threading.Lock()
        
# ------------------------------------------------------------------------------
# Event to signal that a new image is ready for use
myEVT_IMAGEREADY_MAIN = wx.NewEventType()
EVT_IMAGEREADY_MAIN = wx.PyEventBinder(myEVT_IMAGEREADY_MAIN, 1)
myEVT_IMAGEREADY_GUIDER = wx.NewEventType()
EVT_IMAGEREADY_GUIDER = wx.PyEventBinder(myEVT_IMAGEREADY_GUIDER, 1)

class ImageReadyEventMain(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_IMAGEREADY_MAIN, eid=wx.ID_ANY, image=None,
                 image_time=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.image = image
        self.image_time = image_time

class ImageReadyEventGuider(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_IMAGEREADY_GUIDER, eid=wx.ID_ANY, image=None,
                 image_time=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.image = image
        self.image_time = image_time

# ------------------------------------------------------------------------------
# Class to obtain images on a separate thread.
# When run, this connects to the camera, waits for events requesting images,
# or a continuous stream of images, and posts events when each image is ready.
# The camera is disconnected before ending.
class TakeImageThread(threading.Thread):
    def __init__(self, parent, stopevent, onevent, exptime):
        threading.Thread.__init__(self)
        self.daemon = True
        self.parent = parent
        self.stopevent = stopevent
        self.onevent = onevent
        self.continuous = False
        self.camera_id = "ASCOM.SXMain0.Camera"
        self.imshape = (2024, 3040)
        self.ImageReadyEvent = ImageReadyEventMain
        self.check_period = 1.0
        self.cam = None
        self.exptime_lock = threading.Lock()
        self.SetExpTime(exptime)

    def run(self):
        with COMlock:
            self.InitCamera()
        try:
            while not self.stopevent.is_set():
                # only take images when camera is "on" and
                # check for stopevent every second
                if self.onevent.wait(1.0):
                    if self.cam.CameraState > 4:
                        self.Log('Camera error')
                        break
                    if self.cam.CameraState > 0:
                        time.sleep(5)
                        if self.cam.CameraState > 0:
                            self.Log('Aborting current exposure')
                            self.cam.AbortExposure()
                    exptime = self.GetExpTime()
                    self.TakeImage(exptime)
                    if not self.continuous:
                        self.onevent.clear()
        finally:
            self.Log('Disconnecting camera')
            self.Disconnect()

    def InitCamera(self):
        if not simulate:
            try:
                win32com.client.pythoncom.CoInitialize()
                self.cam = win32com.client.Dispatch(self.camera_id)
                self.Log('Starting camera {}'.format(self.camera_id))
                self.Connect()
            except pythoncom.com_error:
                self.Log('Error starting camera - is it connected and on?')
                self.Log("Falling back to simulating camera")
                self.cam = None
        else:
            self.Log("Simulating camera")
            self.cam = None

    def Connect(self):
        if self.cam is not None:
            for i in range(3):
                try:
                    self.Log('Connecting...')
                    #self.cam.Connected = False
                    #time.sleep(1)
                    if not self.cam.Connected:
                        self.cam.Connected = True
                except:
                    self.Log("Problem connecting to camera")
                    self.Log("Trying again in 20 sec")
                    time.sleep(20)
                else:
                    self.cam.StartExposure(0, True) # discard first image
                    # wait for camera to cool?
                    self.Log("Connected to camera {} {}".format(
                             self.camera_id, self.cam.Description))
                    break
            if not self.cam.Connected:
                self.Log("Unable to connect to camera")

    def Disconnect(self):
        if self.cam is not None:
            self.cam.Connected = False
            if not self.cam.Connected:
                self.Log("Disconnected from camera")
            else:
                self.Log("Unable to disconnect from camera")
            self.cam = None
            win32com.client.pythoncom.CoUninitialize()

    def SetExpTime(self, exptime):
        with self.exptime_lock:
            self.exptime = exptime

    def GetExpTime(self):
        with self.exptime_lock:
            return self.exptime

    def Log(self, text):
        wx.PostEvent(self.parent, LogEvent(text=text))

    def TakeImage(self, exptime):
        image = None
        image_time = datetime.utcnow()
        if self.cam is not None:
            self.Log('Taking exposure with {}'.format(self.cam.Description))
            self.cam.StartExposure(exptime, True)
            while (not self.cam.ImageReady) and self.onevent.is_set():
                time.sleep(self.check_period)
            if self.cam.ImageReady and self.onevent.is_set():
                image = np.array(self.cam.ImageArray)
                #self.Log("Check image size: {}x{}, {}x{}".format(
                #         self.cam.CameraXSize,
                #         self.cam.CameraYSize,
                #         image.shape[0],
                #         image.shape[1]))
            else:
                #self.Log('Stopping current exposure early')
                self.cam.StopExposure()
        else:
            image = self.SimulateImage(exptime)
        #self.filters = None  # do not use filters until debayered
        if image is not None:
            wx.PostEvent(self.parent,
                         self.ImageReadyEvent(image=image,
                                              image_time=image_time))

    def SimulateImage(self, exptime):
        # simulate an image
        time.sleep(self.check_period)
        image = np.zeros(self.imshape)
        # add one star per 10000 pixels
        sigma = 4.0
        size = 23
        g = norm.pdf(np.arange(size), (size-1)/2.0, sigma)
        star = np.dot(g[:, None], g[None, :])
        for i in range(np.product(image.shape)//10000):
            x = np.random.choice(image.shape[0]-size)
            y = np.random.choice(image.shape[1]-size)
            flux = np.random.poisson(10000 * exptime)
            image[x:x+size,y:y+size] += star * flux
        # add bright sky background
        image += 1000 * exptime
        # sloping response / vignetting
        image *= np.arange(image.shape[1])/(2.0*image.shape[0]) + 0.75
        # poisson sample
        image = np.random.poisson(image)
        # add bias with read noise
        image += np.random.normal(800, 20, size=image.shape)
        return image

# ------------------------------------------------------------------------------
# Subclass to obtain images from main camera on a separate thread.
class TakeMainImageThread(TakeImageThread):
    def __init__(self, parent, stopevent, onevent, exptime):
        TakeImageThread.__init__(self, parent, stopevent, onevent, exptime)
        self.start()

# ------------------------------------------------------------------------------
# Subclass to obtain images from guide camera on a separate thread.
class TakeGuiderImageThread(TakeImageThread):
    def __init__(self, parent, stopevent, onevent, exptime):
        TakeImageThread.__init__(self, parent, stopevent, onevent, exptime)
        self.continuous = True
        # Something bizarre is happening!
        # Somehow, the camera connection for the guider thread is being used
        # for the main thread.  Don't know how!
        # Next thing to try is using different ImageReadyEvent classes.
        self.camera_id = "ASCOM.SXGuide0.Camera"
        self.imshape = (600, 400)
        self.check_period = 0.1
        self.ImageReadyEvent = ImageReadyEventGuider
        self.start()
