# === IMPORTS ==================================================================

import enum
import shot
import string
import typing
import xlsxwriter


# === ENUM =====================================================================


# === GLOBAL CONSTANTS =========================================================

DEFAULT_FILE_NAME = "shotParser"


# === CLASSES ==================================================================

class xlsx:
    class Column(enum.Enum):
        Field = 0,
        Value = 1,
        
    class Row(enum.Enum):
        Name = 0,
        MaxMagnitude = 1
        MaxMagnitudeIndex = 2,
        ShotConfidence = 3,
        ShotMagnitude = 4,
        ShotMagnitudeIndex = 5,
        ShotMagnitudeMinus4 = 6,
        ShotMagnitudeMinus3 = 7,
        ShotMagnitudeMinus2 = 8,
        ShotMagnitudeMinus1 = 9,
        ShotMagnitudeZero = 10,
        ShotMagnitudePlus1 = 11,
        ShotMagnitudePlus2 = 12,
        ShotMagnitudePlus3 = 13,
        ShotMagnitudePlus4 = 14,
        
    EXTENSION = '.xlsx'
    
    DEFAULT_SHEET_NAMES = (
        'Reset',
        'No Shot',
        'Low',
        'Medium',
        'High',
        'Very High',
    )
    
    FIELD_NAMES = [
        'Name',
        'Max Magnitude',
        '[Max Magnitude]',
        'Shot Confidence',
        'Shot Magnitude',
        '[Shot Magnitude]',
        'Shot[-4] Magnitude',
        'Shot[-3] Magnitude',
        'Shot[-2] Magnitude',
        'Shot[-1] Magnitude',
        'Shot[ 0] Magnitude',
        'Shot[+1] Magnitude',
        'Shot[+2] Magnitude',
        'Shot[+3] Magnitude',
        'Shot[+4] Magnitude',
    ]
        
    class sheet:
        def __init__(self, name: str, ws : xlsxwriter.Workbook.worksheet_class):
            self.name = name
            self.colIndex = 0
            self.ws: xlsxwriter.Workbook.worksheet_class = ws
        
    def __init__(self, fileName: str = DEFAULT_FILE_NAME, sheetNames: typing.List[str] = DEFAULT_SHEET_NAMES):
        self.fileName: str = str(fileName)
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.fileName + self.EXTENSION)
        self.sheets: typing.List[self.sheet] = []
        for name in sheetNames:
            s = self.sheet(name, self.wb.add_worksheet(name))
            for i, field in enumerate(self.FIELD_NAMES):
                s.ws.write(i, s.colIndex, field)
            s.colIndex += 1
            self.sheets.append(s)
            
    def writeShotData(self, data : shot.data):
        s : self.sheet = self.sheets[data.shotConfidence.value]
        col = s.colIndex
        ws : xlsxwriter.Workbook.worksheet_class = s.ws
        ws.write(self.Row.Name.value, col, data.name)
        ws.write(self.Row.MaxMagnitude.value, col, data.maxAccel)
        ws.write(self.Row.MaxMagnitudeIndex.value, col, data.maxAccelIndex)
        ws.write(self.Row.ShotConfidence.value, col, data.shotConfidence)
        ws.write(self.Row.ShotMagnitude.value, col, data.shotMagnitude)
        ws.write(self.Row.ShotMagnitudeIndex.value, col, data.shotIndex)
        for i, j in enumerate(range(-4, 4 + 1)):
            index = data.shotIndex + j
            if index >= 0 and index < data.numSamples:
                ws.write(self.Row.ShotMagnitudeMinus1.value + i, col, data.accel[index].magnitude)
        s.colIndex += 1
        
    def finalize(self):
        self.wb.close()