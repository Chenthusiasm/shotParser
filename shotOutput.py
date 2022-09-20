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
        Field = 0
        Value = 1
        
    class Row(enum.Enum):
        Name = 0
        MaxMagnitude = 1
        MaxMagnitudeIndex = 2
        MaxMagnitudeRange = 4
        ShotConfidence = 14
        ShotMagnitude = 15
        ShotMagnitudeIndex = 16
        ShotMagnitudeRange = 18
        
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
        'Max',
        'Max[]',
        '',
        'Max[-4]',
        'Max[-3]',
        'Max[-2]',
        'Max[-1]',
        'Max[ 0]',
        'Max[+1]',
        'Max[+2]',
        'Max[+3]',
        'Max[+4]',
        '',
        'Confidence',
        'Shot',
        'Shot[]',
        '',
        'Shot[-4]',
        'Shot[-3]',
        'Shot[-2]',
        'Shot[-1]',
        'Shot[ 0]',
        'Shot[+1]',
        'Shot[+2]',
        'Shot[+3]',
        'Shot[+4]',
    ]
        
    class sheet:
        def __init__(self, name: str, ws : xlsxwriter.Workbook.worksheet_class):
            self.name : str = name
            self.colIndex : int = 0
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
        ws.write(self.Row.MaxMagnitude.value, col, data.maxAccel.magnitude)
        ws.write(self.Row.MaxMagnitudeIndex.value, col, data.maxAccelIndex)
        for i, j in enumerate(range(-4, 4 + 1)):
            index = data.maxAccelIndex + j
            if index >= 0 and index < data.numSamples:
                ws.write(self.Row.MaxMagnitudeRange.value + i, col, data.accel[index].magnitude)
        ws.write(self.Row.ShotConfidence.value, col, data.shotConfidence.value)
        ws.write(self.Row.ShotMagnitude.value, col, data.shot.magnitude)
        ws.write(self.Row.ShotMagnitudeIndex.value, col, data.shotIndex)
        for i, j in enumerate(range(-4, 4 + 1)):
            index = data.shotIndex + j
            if index >= 0 and index < data.numSamples:
                ws.write(self.Row.ShotMagnitudeRange.value + i, col, data.accel[index].magnitude)
        s.colIndex += 1
        
    def finalize(self):
        self.wb.close()