import wx
import threading

from logevent import *

from scipy.ndimage.filters import median_filter, gaussian_filter
import astropy.io.fits as pyfits

import astrotortilla.solver.AstrometryNetSolver as AstSolve
from astrotortilla.units import Coordinate

# ------------------------------------------------------------------------------
# Event to signal that a new solution is ready for use
myEVT_SOLUTIONREADY = wx.NewEventType()
EVT_SOLUTIONREADY = wx.PyEventBinder(myEVT_SOLUTIONREADY, 1)
class SolutionReadyEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_SOLUTIONREADY, eid=wx.ID_ANY, solution=None,
                 image_time=None, filenames=[]):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.solution = solution
        self.image_time = image_time
        self.filenames = filenames

# ------------------------------------------------------------------------------
# Class to obtain plate solution on a separate thread.
# When run, this creates a solver, waits filenames in a Queue,
# and posts events when a solution is ready.
# Stops when a None is added to the Queue.
class SolverThread(threading.Thread):
    def __init__(self, parent, incoming, directory=None, timeout=60):
        threading.Thread.__init__(self)
        self.daemon = True
        self.parent = parent
        self.incoming = incoming
        self.timeout = timeout
        self.dir = directory
        self.lastlog = ''
        self.start()

    def run(self):
        self.solver = AstSolve.AstrometryNetSolver(workDirectory=self.dir)
        self.solver.timeout = self.timeout
        self.solver.setProperty('downscale', 2)
        self.solver.setProperty('xtra', '--depth 20 --no-plots '
                                        '-N none')
        try:
            while True:
                incoming = self.incoming.get()
                if incoming is None:
                    break
                filters, fn, filenames, image_time, position = incoming
                self.CreateSolveImage(filters, fn)
                kwargs = {'minFov': 0.25, 'maxFov': 0.5,
                          'targetRadius': 5}
                if position is not None:
                    target = Coordinate(position.ra.deg, position.dec.deg)                    
                    kwargs['target'] = target
                solution = self.solver.solve(fn, callback=self.Log,
                                             **kwargs)
                if solution is None:
                    del kwargs['target']
                    solution = self.solver.solve(fn, callback=self.Log,
                                                 **kwargs)
                wx.PostEvent(self.parent,
                             SolutionReadyEvent(solution=solution,
                                            image_time=image_time,
                                            filenames=filenames))
        except Exception as detail:
            self.Log('Error in solver:\n{}'.format(detail))
            raise

    def Log(self, text):
        try:
            if text is not None:
                text = text.strip()
            if ((len(text) > 0) and (text != self.lastlog)
                and ('did not' not in text)):
                    wx.PostEvent(self.parent, LogEvent(text=text))
                    self.lastlog = text
        except:
            pass
            #self.Log('logging error')
            
    def CreateSolveImage(self, filters, filename):
        self.Log('Filtering image for astrometry')
        image = filters.sum(0)
        #background = median_filter(image, (25, 25))
        #image -= background
        image = median_filter(image, (3,3))
        image = gaussian_filter(image, 2)
        pyfits.writeto(filename, image, clobber=True)
