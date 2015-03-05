import wx

# ------------------------------------------------------------------------------        
# Event to signal a new log entry is pending
myEVT_LOG = wx.NewEventType()
EVT_LOG = wx.PyEventBinder(myEVT_LOG, 1)
class LogEvent(wx.PyCommandEvent):
    def __init__(self, etype=myEVT_LOG, eid=wx.ID_ANY, text=None):
        wx.PyCommandEvent.__init__(self, etype, eid)
        self.text = text
