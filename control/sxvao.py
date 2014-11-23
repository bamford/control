# -*- coding: utf-8 -*-

# sxvao.py

import serial


class SXVAO():
    def __init__(self, parent, comport, timeout=10,
                 steps_per_pixel=6.0, mount_steps_per_pixel=6.0,
                 switch_xy=False, reverse_x=False, reverse_y=False,
                 max_steps=10, steps_limit=1000):
        self.parent = parent
        self.comport = comport
        self.timeout = timeout
        self.steps_per_pixel = steps_per_pixel
        self.mount_steps_per_pixel = mount_steps_per_pixel
        self.switch_xy = switch_xy
        self.reverse_x = reverse_x
        self.reverse_y = reverse_y
        self.max_steps = max_steps
        self.steps_limit = steps_limit
        self.ao = None
        self.count_steps_N = 0
        self.count_steps_W = 0
    
    def Connect():
        if self.ao is None:
            self.ao = serial.Serial(self.comport, timeout=10)
            self.ao.write('X')
            response = self.ao.read(1)
            return response == 'Y'
        else:
            return False
        
    def Disconnect():
        if self.ao is not None:
            self.ao.close()
            self.ao = None

    def MakeCorrection(dx, dy):
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

    def DeltaToSteps(d):
        return int(round(abs(dx)/self.steps_per_pixel))

    def MakeSteps(dir, n=1):
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

    def MakeMountSteps(dir, n=1):
        # dir must be one of [N, S, T, W]
        command = 'M{:1s}{:05d}'.format(dir, n)
        self.ao.write(command)
        response = self.ao.read(1)
        if response == 'M':
            self.parent.Log('Mount took {:d} steps {:s}'.format(n, dir))
        else:
            self.parent.Log('Mount stepping failed')
        
    def RecentreMountIfNeeded():
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

    def Centre():
        self.ao.write('K')
        response = self.ao.read(1)
        return response == 'K'

    
