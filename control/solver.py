import wx
import threading
import time
from datetime import datetime
import numpy as np

from logevent import *

from astrotortilla.solver import AstrometryNetSolver, AstrometryNetWebSolver


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
        self.solver = AstrometryNetWebSolver(timout=self.timeout,
                                             callback=self.Log)
        try:
            while True:
                incoming = self.incoming.get()
                if incoming is not None:
                    fn, kwargs = incoming
                    solution = self.solver.solve(fn, **kwargs)
                    wx.PostEvent(self.parent,
                                 SolutionReadyEvent(solution=solution,
                                                    image_time=image_time))
        except Exception as detail:
            self.Log('Error in solver:\n{}'.format(detail))
            raise

    def Log(self, text):
        wx.PostEvent(self.parent, LogEvent(text=text))
