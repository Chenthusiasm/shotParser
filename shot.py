# === IMPORTS ==================================================================

import enum
import math
import operator
import os
import string
import typing


# === GLOBAL CONSTANTS =========================================================

LSB_TO_G_DIVISOR = 4096

TYPE_IMU_GRYO = 1
TYPE_IMU_ACCEL = 2
TYPE_HI_G_ACCEL = 3
TYPE_CALIBRATION = 4
TYPE_SETTINGS = 5
TYPE_HI_G_ACCEL_COMP = 6
MAX_ATTRIBUTE = 'magnitude'

# === ENUM =====================================================================

class LineIndex(enum.Enum):
    Type = 0
    X = 1
    Y = 2
    Z = 3
NUM_LINE_INDICES = len(LineIndex)


class Handedness(enum.Enum):
    Right = 0
    Left = 1
NUM_HANDEDNESS = len(Handedness)


# === HELPER FUNCTIONS =========================================================

def findThreeAxisMagnitude(x : float, y : float, z : float) -> float:
    return math.sqrt((x * x) + (y * y) + (z * z))

def convertLsbToG(a : int) -> float:
    return float(a / LSB_TO_G_DIVISOR)

def convertLsbToG(a : float) -> float:
    return float(a / LSB_TO_G_DIVISOR)


# === CLASSES ==================================================================

class vector:
    def __init__(self, x : float, y : float, z: float):
        self.x: float = x
        self.y: float = y
        self.z: float = z
        self.magnitude: float = findThreeAxisMagnitude(self.x, self.y, self.z)
        self.xG: float = convertLsbToG(self.x)
        self.yG: float = convertLsbToG(self.y)
        self.zG: float = convertLsbToG(self.z)
        self.magnitudeG: float = convertLsbToG(self.magnitude)
        
    def getVectorSum(self) -> int:
        return abs(self.x) + abs(self.y) + abs(self.z) 
    
    
class data:
    def __init__(self, fileName: str = ''):
        self.fileName: str = fileName
        self.name: str = ''
        self.numIMU: int = 0
        self.gyro: typing.List[vector] = []
        self.accel: typing.List[vector] = []
        self.calibration: vector
        self.handedness: Handedness = Handedness.Right
        self.maxGyro: vector
        self.maxAccel: vector
        
    def process(self):
        if self.fileName:
            self.name = self.fileName.replace('.csv', '')
            with open(os.path.join(os.getcwd(), self.fileName), 'r') as file:
                readLines = file.readLines()
            file.close()
            self.numIMU = 0
            for line in readLines:
                line = line.strip()
                entries = line.split(',')
                type = int(entries[LineIndex.Type.value])
                v = vector(float(entries[LineIndex.X.value]),
                           float(entries[LineIndex.Y.value]),
                           float(entries[LineIndex.Z.value]))
                
                if (type == TYPE_IMU_GRYO):
                    self.gyro.append(v)
                elif (type == TYPE_IMU_ACCEL):
                    self.accel.append(v)
                elif (type == TYPE_CALIBRATION):
                    self.calibration = v
                elif (type == TYPE_SETTINGS):
                    n = int(entries[LineIndex.X.value])
                    self.handedness = Handedness.Right
                    if (n == Handedness.Left.value):
                        self.handedness = Handedness.Left
        
    def processFile(self, fileName : str):
        self.fileName = fileName
        self.process()
        
    def analyze(self):
        self.maxGyro = max(self.gyro, key = operator.attrgetter(MAX_ATTRIBUTE))
        self.maxAccel = max(self.gyro, key = operator.attrgetter(MAX_ATTRIBUTE))
