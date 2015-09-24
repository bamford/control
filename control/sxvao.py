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
            try:
                self.ao = serial.Serial(self.comport, timeout=10)
            except OSError:
                return False
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
            self.MakeSteps(dir, n)
            nx -= n
        ny = self.DeltaToSteps(dy)
        dir = 'N' if dy > 0 else 'S'
        while ny > 0:
            n = ny%self.max_steps
            self.MakeSteps(dir, n)
            ny -= n
        return True

    def MakeMountCorrection(self, dx, dy):
        if self.mount_reverse_x:
            dx = -dx
        if self.mount_reverse_y:
            dy = -dy
        if self.mount_switch_xy:
            dx, dy = dy, dx
        nx = self.DeltaToMountSteps(dx)
        dir = 'T' if dx > 0 else 'W'
        self.MakeMountSteps(dir, nx)
        ny = self.DeltaToMountSteps(dy)
        dir = 'N' if dy > 0 else 'S'
        self.MakeMountSteps(dir, ny)
        return True

    def DeltaToSteps(self, d):
        return int(round(abs(d)/self.steps_per_pixel))

    def DeltaToMountSteps(self, d):
        return int(round(abs(d)/self.mount_steps_per_pixel))

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
        self.RecentreMountIfNeeded(force=(response=='L'))

    def MakeMountSteps(self, dir, n=1):
        # dir must be one of [N, S, T, W]
        command = 'M{:1s}{:05d}'.format(dir, n)
        self.ao.write(command)
        response = self.ao.read(1)
        if response == 'M':
            self.parent.Log('Mount took {:d} steps {:s}'.format(n, dir))
        else:
            self.parent.Log('Mount stepping failed')

    def RecentreMountIfNeeded(self, force=False):
        # Move to approximately recentre AO with
        # opposing move for scope to keep image stationary
        # North-South
        mount_steps_N = (self.count_steps_N * self.mount_steps_per_pixel
                         / self.steps_per_pixel)
        if self.count_steps_N > self.steps_limit or force:
            self.MakeSteps('S', abs(self.count_steps_N))
            self.MakeMountSteps('N', abs(mount_steps_N))
            self.count_steps_N = 0
        elif self.count_steps_N < -self.steps_limit:
            self.MakeSteps('N', abs(self.count_steps_N))
            self.MakeMountSteps('S', abs(mount_steps_N))
            self.count_steps_N = 0
        # West-East
        mount_steps_W = (self.count_steps_W * self.mount_steps_per_pixel
                         / self.steps_per_pixel)
        if self.count_steps_W > self.steps_limit or force:
            self.MakeSteps('T', abs(self.count_steps_W))
            self.MakeMountSteps('W', abs(mount_steps_W))
            self.count_steps_W = 0
        elif self.count_steps_W < -self.steps_limit:
            self.MakeSteps('W', abs(self.count_steps_W))
            self.MakeMountSteps('T', abs(mount_steps_W))
            self.count_steps_W = 0

    def Centre(self):
        self.ao.write('K')
        response = self.ao.read(1)
        return response == 'K'
