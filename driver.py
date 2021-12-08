import sys
import time

import numpy as np
import win32api
import win32con
from serial import Serial


class MouseData:
    def __init__(self, x, y, z, left_clicked, right_clicked):
        self.x = x
        self.y = y
        self.z = z
        self.left_clicked = left_clicked
        self.right_clicked = right_clicked

    def __str__(self):
        return f"MouseData(x={self.x}, y={self.y}, z={self.z}, left_clicked={self.left_clicked}, right_clicked={self.right_clicked})"


def kalman_filter(data,preData,var,Q):
    preVar = var + Q
    K = preVar / (preVar + Q)
    outputData = preData + K * (data - preData)
    var = (1- K) * preVar
    preData = outputData
    return (outputData,var)

def get_data(serial_connection: Serial) -> MouseData:
    line = serial_connection.readline().decode('utf-8')
    if not line:
        return None
    line = line[:-1].split(',')
    x = float(line[0])
    y = float(line[1])
    z = float(line[2])
    left_clicked = line[3] == '1'
    right_clicked = line[4] == '1'
    result = MouseData(x, y, z, left_clicked, right_clicked)
    print(result)
    return result

if __name__ == '__main__':
    serial_connection = Serial(sys.argv[1] if len(sys.argv) >= 3 else 'COM3')
    try:
        cursor_speed =  int(sys.argv[2]) if len(sys.argv) >= 3 else 30
        moving_threshold = float(sys.argv[3]) if len(sys.argv) >= 4 else 0.4
        damp_rate = float(sys.argv[4]) if len(sys.argv) >= 5 else 0.8
        Q = float(sys.argv[5]) if len(sys.argv) >= 6 else 0.1

        left_clicked, right_clicked = False, False

        # initialization for kalman filter
        preData, var, xPreData, yPreData, xVar, yVar = 0, 0, 0, 0, 0, 0

        print('Fase de calibracion, por favor coloque el telefono sobre una superficie y no lo mueva.')
        t1 = time.process_time()
        # List to store the acceleration data for calibration
        caliAcc = []
        # The number of calibrated data
        numCali = 0
        while numCali < 100:
            data = get_data(serial_connection)
            if data is None:
                continue
            caliAcc.append([data.x / 32.0 * 9.8, data.y / 32.0 * 9.8])
            numCali += 1
        print('Calibracion completada. Si los parametros cambian, por favor reinicie el programa.')
        arrayCaliAcc = np.array(caliAcc)
        axave = arrayCaliAcc[:,0].mean()
        ayave = arrayCaliAcc[:,1].mean()

        print('Resultado de calibracion:',axave,ayave)
        t2 = time.process_time()
        print('Tiempo:',t2 - t1,'segundos')
        print('Puede en este momento utilizar su telefono para controlar el mouse.')
        # set the initial speed of the cursor to zero
        vx, vy, t0, t1 = 0, 0, 1, 0
        while(True):
            data = get_data(serial_connection)
            if data is None:
                continue
            ax, xVar = kalman_filter(data.x / 32.0 * 9.8 - axave, xPreData,xVar,Q)
            ay, yVar = kalman_filter(data.y / 32.0 * 9.8 - ayave,yPreData,yVar,Q)
            # controlar velocidad
            vx = vx * damp_rate
            vy = vy * damp_rate
            if abs(ax) > moving_threshold:
                dvx = np.exp(-abs(vx)) * ax * cursor_speed
                vx = vx + dvx
            if abs(ay) > moving_threshold:
                dvy = np.exp(-abs(vy)) * ay * cursor_speed
                vy = vy + dvy
            # actualizar posicion de cursor
            pos = win32api.GetCursorPos()
            pos = (int(round(pos[0] + vx)),int(round(pos[1] - vy)))
            win32api.SetCursorPos(pos)

            # Clicks
            if data.left_clicked != left_clicked:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN if data.left_clicked else win32con.MOUSEEVENTF_LEFTUP,0,0)
            
            if data.right_clicked != right_clicked:
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN if data.right_clicked else win32con.MOUSEEVENTF_RIGHTUP,0,0)

            left_clicked = data.left_clicked
            right_clicked = data.right_clicked

    finally:
        serial_connection.close()
