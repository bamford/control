# -*- coding: utf-8 -*-

# sxvao.py

import serial


class SXVAO():
    def __init__(self, parent, comport, timeout=10):
        self.parent = parent
        self.comport = comport
        self.timeout = timeout
        # configuration attributes
        self.steps_per_pixel = 6.0
        self.switch_xy = False
        self.reverse_x = False
        self.reverse_y = False
        self.mount_steps_per_pixel = 6.0
        self.mount_switch_xy = False
        self.mount_reverse_x = False
        self.mount_reverse_y = False
        self.max_steps = 10
        self.steps_limit = 1000
        # internal variables
        self.ao = None
        self.count_steps_N = 0
        self.count_steps_W = 0
    
    def Connect(self):
        if self.ao is None:
            self.ao = serial.Serial(self.comport, timeout=10)
            self.ao.write('X')
            response = self.ao.read(1)
            return response == 'Y'
        else:
            return False
        
    def Disconnect(self):
        if self.ao is not None:
            self.ao.close()
            self.ao = None

    def MakeCorrection(self, dx, dy):
        if self.reverse_x:
            dx = -dx
        if self.reverse_y:
            dy = -dy
        if self.switch_xy:
            dx, dy = dy, dx
        nx = self.DeltaToSteps(dx)
        dir = 'T' if dx > 0 else 'W'
        while nx > 0:
            n = nx%self.max_steps
            MakeSteps(dir, n)
            nx -= n
        ny = self.DeltaToSteps(dy)
        dir = 'N' if dy > 0 else 'S'
        while ny > 0:
            n = ny%self.max_steps
            MakeSteps(dir, n)
            ny -= n

    def MakeMountCorrection(self, dx, dy):
        if self.mount_reverse_x:
            dx = -dx
        if self.mount_reverse_y:
            dy = -dy
        if self.mount_switch_xy:
            dx, dy = dy, dx
        nx = self.DeltaToMountSteps(dx)
        dir = 'T' if dx > 0 else 'W'
        MakeMountSteps(dir, nx)
        ny = self.DeltaToMountSteps(dy)
        dir = 'N' if dy > 0 else 'S'
        MakeMountSteps(dir, ny)

    def DeltaToSteps(self, d):
        return int(round(abs(dx)/self.steps_per_pixel))

    def MakeSteps(self, dir, n=1):
        # dir must be one of [N, S, T, W]
        if dir == 'N':
            self.count_steps_N += n
        elif dir == 'S':
            self.count_steps_N -= n
        if dir == 'W':
            self.count_steps_W += n
        elif dir == 'T':
            self.count_steps_W -= n
        command = 'G{:1s}{:05d}'.format(dir, n)
        self.ao.write(command)
        response = self.ao.read(1)
        if response == 'G':
            self.parent.Log('AO unit took {:d} steps {:s}'.format(n, dir))
        elif response == 'L':
            self.parent.Log('AO unit hit limit')
        else:
            self.parent.Log('AO unit stepping failed')
        RecentreMountIfNeeded(force=(response=='L'))

    def MakeMountSteps(self, dir, n=1):
        # dir must be one of [N, S, T, W]
        command = 'M{:1s}{:05d}'.format(dir, n)
        self.ao.write(command)
        response = self.ao.read(1)
        if response == 'M':
            self.parent.Log('Mount took {:d} steps {:s}'.format(n, dir))
        else:
            self.parent.Log('Mount stepping failed')
        
    def RecentreMountIfNeeded(self):
        # Move to approximately recentre AO with
        # opposing move for scope to keep image stationary
        # North-South
        mount_steps_N = (self.count_steps_N * self.mount_steps_per_pixel
                         / self.steps_per_pixel)
        if self.count_steps_N > self.steps_limit:
            MakeSteps('S', abs(self.count_steps_N))
            MakeMountSteps('N', abs(mount_steps_N))
            self.count_steps_N = 0
        elif self.count_steps_N < -self.steps_limit:
            MakeSteps('N', abs(self.count_steps_N))
            MakeMountSteps('S', abs(mount_steps_N))
            self.count_steps_N = 0
        # West-East
        mount_steps_W = (self.count_steps_W * self.mount_steps_per_pixel
                         / self.steps_per_pixel)
        if self.count_steps_W > self.steps_limit:
            MakeSteps('T', abs(self.count_steps_W))
            MakeMountSteps('W', abs(mount_steps_W))
            self.count_steps_W = 0
        elif self.count_steps_W < -self.steps_limit:
            MakeSteps('W', abs(self.count_steps_W))
            MakeMountSteps('T', abs(mount_steps_W))
            self.count_steps_W = 0

    def Centre(self):
        self.ao.write('K')
        response = self.ao.read(1)
        return response == 'K'

    
