# === IMPORTS ==================================================================

import enum
import math
import operator
import os
import string
import typing


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

class ShotConfidence(enum.Enum):
    Reset = 0
    NoShot = 1
    VeryLow = 2
    Low = 3
    Medium = 4
    High = 5
    VeryHigh = 6
NUM_SHOT_CONFIDENCE = len(ShotConfidence)


# === GLOBAL CONSTANTS =========================================================

LSB_TO_G_DIVISOR = 4096

TYPE_IMU_GRYO = 1
TYPE_IMU_ACCEL = 2
TYPE_HI_G_ACCEL = 3
TYPE_CALIBRATION = 4
TYPE_SETTINGS = 5
TYPE_HI_G_ACCEL_COMP = 7
MAX_ATTRIBUTE = 'magnitude'


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
        self.filePath: str = ''
        self.name: str = ''
        self.gyro: typing.List[vector] = []
        self.accel: typing.List[vector] = []
        self.hiG: typing.List[vector] = []
        self.calibration: vector
        self.handedness: Handedness = Handedness.Right
        self.maxGyro: vector
        self.maxAccel: vector
        self.maxHiG: vector
        self.maxGyroIndex: int = 0
        self.maxAccelIndex: int = 0
        self.maxHiGIndex: int = 0
        self.shot: vector
        self.shotIndex: int = 0
        self.shotConfidence: ShotConfidence = ShotConfidence.Reset
        
    def process(self):
        if self.fileName:
            self.name = self.fileName.replace('.csv', '')
            self.filePath = os.path.join(os.getcwd(), self.fileName)
            with open(self.filePath, 'r') as file:
                readLines = file.readlines()
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
                elif (type == TYPE_HI_G_ACCEL):
                    self.hiG.append(v)
                elif (type == TYPE_CALIBRATION):
                    self.calibration = v
                elif (type == TYPE_SETTINGS):
                    n = int(entries[LineIndex.X.value])
                    self.handedness = Handedness.Right
                    if (n == Handedness.Left.value):
                        self.handedness = Handedness.Left
                elif (type == TYPE_HI_G_ACCEL_COMP):
                    x : int = int(entries[LineIndex.X.value])
                    y : int = int(entries[LineIndex.Y.value])
                    z : int = int(entries[LineIndex.Z.value])
                    HI_SHIFT = 8
                    LO_SHIFT = 0
                    MASK = 0xff
                    v0 : vector = vector(
                        (x >> LO_SHIFT) & MASK,
                        (x >> HI_SHIFT) & MASK,
                        (y >> LO_SHIFT) & MASK)
                    v1 : vector = vector(
                        (y >> HI_SHIFT) & MASK,
                        (z >> LO_SHIFT) & MASK,
                        (z >> HI_SHIFT) & MASK)
                    self.hiG.append(v0)
                    self.hiG.append(v1)
        
    def processFile(self, fileName : str):
        self.fileName = fileName
        self.process()
        
    def analyze(self):
        self.maxGyro = max(self.gyro, key = operator.attrgetter(MAX_ATTRIBUTE))
        self.maxAccel = max(self.accel, key = operator.attrgetter(MAX_ATTRIBUTE))
        self.maxHiG = max(self.hiG, key = operator.attrgetter(MAX_ATTRIBUTE))
        self.maxGyroIndex = self.gyro.index(self.maxGyro)
        self.maxAccelIndex = self.accel.index(self.maxAccel)
        self.maxHiGIndex = self.hiG.index(self.maxHiG)
        
    def findAndScoreShot(self):
        self.shotConfidence = ShotConfidence.NoShot
        for i, v in enumerate(self.accel):
            if v.magnitude >= 40000 and self.shotConfidence.value < ShotConfidence.VeryHigh.value:
                self.shotConfidence = ShotConfidence.VeryHigh
                self.shotIndex = i
            elif v.magnitude >= 35000 and self.shotConfidence.value < ShotConfidence.High.value:
                self.shotConfidence = ShotConfidence.High
                self.shotIndex = i
            elif v.magnitude >= 30000 and self.shotConfidence.value < ShotConfidence.Medium.value:
                self.shotConfidence = ShotConfidence.Medium
                self.shotIndex = i
            elif v.magnitude >= 20000 and self.shotConfidence.value < ShotConfidence.Low.value:
                NEIGHBOR_THRESHOLD = 12000
                numSamples = len(self.accel)
                if (i > 0) and (self.accel[i - 1].magnitude >= NEIGHBOR_THRESHOLD):
                    self.shotConfidence = ShotConfidence.Low
                elif (i < (numSamples - 1)) and (self.accel[i + 1].magnitude >= NEIGHBOR_THRESHOLD):
                    self.shotConfidence = ShotConfidence.Low
                else:
                    self.shotConfidence = ShotConfidence.VeryLow
                self.shotIndex = i
        self.shot = self.accel[self.shotIndex]
