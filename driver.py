import sys
import time

import numpy as np
import serial
import win32api
import win32con


def kalmanFilter(data,preData,var,Q):
    preVar = var + Q
    K = preVar / (preVar + Q)
    outputData = preData + K * (data - preData)
    var = (1- K) * preVar
    preData = outputData
    return (outputData,var)

def leftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

def rightClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)

if __name__ == '__main__':
    s = serial.Serial(sys.argv[1])
    # command line input should be as the following format
    # python driver.py [serial port number] [cursor move speed] [threshhold for moving] [damp rate] [Q] [R]
    # python driver.py COM3 20 0.4 0.8 0.1 0.25
    if len(sys.argv) >= 3:
        v = int(sys.argv[2])
    else:
        v = 20
    if len(sys.argv) >= 4:
        Thresh = float(sys.argv[3])
    else:
        Thresh = 0.4
    if len(sys.argv) >= 5:
        dampRate = float(sys.argv[4])
    else:
        dampRate = 0.8

    # initialization for kalman filter
    preData = 0
    var = 0
    if len(sys.argv) >= 6:
        Q = float(sys.argv[5])
    else:
        Q = 0.1

    xPreData = 0
    yPreData = 0
    xVar = 0
    yVar = 0

    print('Begin to calibrate, please put the mobile phone on the surface and do not move')
    t1 = time.process_time()
    # List to store the acceleration data for calibration
    caliAcc = []
    # The number of calibrated data
    numCali = 0
    while numCali < 100:
        byte = s.read(1)
        while(byte != bytes([22])):
            byte = s.read(1)
        x0b = s.read(1)
        y0b = s.read(1)
        x0int = int.from_bytes(x0b,byteorder = 'big')-128
        y0int = int.from_bytes(y0b,byteorder = 'big')-128
        ax = x0int / 32.0 * 9.8
        ay = y0int / 32.0 * 9.8
        caliAcc.append([ax,ay])
        numCali += 1
    print('Calibration Completed. If you change the working environment, please restart the program')
    arrayCaliAcc = np.array(caliAcc)
    axave = arrayCaliAcc[:,0].mean()
    ayave = arrayCaliAcc[:,1].mean()

    print('Calibration result:',axave,ayave)
    t2 = time.process_time()
    print('Time:',t2 - t1,'seconds')
    print('It is now available to use your mobile phone to control the cursor')
    # set the initial speed of the cursor to zero
    vx = 0
    vy = 0
    t0 = 1
    t1 = 0
    while(True):
        byte = s.read(1)
        if byte == bytes([22]):
            x0b = s.read(1)
            y0b = s.read(1)
            x0int = int.from_bytes(x0b,byteorder = 'big')-128
            y0int = int.from_bytes(y0b,byteorder = 'big')-128
            # decode the acceleration data and calibrate it using the data in the calibration step
            ax = x0int / 32.0 * 9.8 - axave
            ay = y0int / 32.0 * 9.8 - ayave
            ax, xVar = kalmanFilter(ax,xPreData,xVar,Q)
            ay, yVar = kalmanFilter(ay,yPreData,yVar,Q)
            # damp the speed
            vx = vx * dampRate
            vy = vy * dampRate
            # change the speed if the acceleration exceed the threshhold
            if abs(ax) > Thresh:
                dvx = np.exp(-abs(vx)) * ax * v
                vx = vx + dvx
            if abs(ay) > Thresh:
                dvy = np.exp(-abs(vy)) * ay * v
                vy = vy + dvy
            # update the position of the cursor
            pos = win32api.GetCursorPos()
            pos = (int(round(pos[0] + vx)),int(round(pos[1] - vy)))
            win32api.SetCursorPos(pos)

        # deal with click
        elif byte == bytes([17]):
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
        elif byte == bytes([18]):
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
        elif byte == bytes([20]):
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
        elif byte == bytes([21]):
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)