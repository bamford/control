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
                 image_time=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.solution = solution
        self.image_time = image_time

# ------------------------------------------------------------------------------
# Class to obtain plate solution on a separate thread.
# When run, this creates a solver, waits filenames in a Queue,
# and posts events when a solution is ready.
# Stops when a None is added to the Queue.
class SolverThread(threading.Thread):
    def __init__(self, parent, incoming, timeout=300):
        threading.Thread.__init__(self)
        self.parent = parent
        self.incoming = incoming
        self.timeout = timeout
        self.start()

    def run(self):
        self.solver = AstSolve.AstrometryNetSolver()
        self.solver.timeout = self.timeout
        try:
            while True:
                incoming = self.incoming.get()
                if incoming is None:
                    break
                filters, fn, image_time, position = incoming
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
                                            image_time=image_time))
        except Exception as detail:
            self.Log('Error in solver:\n{}'.format(detail))
            raise

    def Log(self, text):
        if (text is not None) and (len(text.strip()) > 0):
            wx.PostEvent(self.parent, LogEvent(text=text.strip()))
            
    def CreateSolveImage(self, filters, filename):
        self.Log('Filtering image for astrometry')
        image = filters.sum(0)
        background = median_filter(image, (25, 25))
        image -= background
        image = gaussian_filter(image, 3)
        pyfits.writeto(filename, image, clobber=True)
