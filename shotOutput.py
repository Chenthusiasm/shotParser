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
        MaxGyro = 1
        MaxGyroIndex = 2
        MaxHiG = 3
        MaxHiGIndex = 4
        MaxAccel = 5
        MaxAccelIndex = 6
        MaxAccelRange = 8
        ShotConfidence = 18
        ShotAccel = 19
        ShotGyro = 20
        ShotAccelIndex = 21
        ShotAccelRange = 23
        
    EXTENSION = '.xlsx'
    
    DEFAULT_SHEET_NAMES = (
        'Reset',
        'No Shot',
        'Very Low',
        'Low',
        'Medium',
        'High',
        'Very High',
    )
    
    FIELD_NAMES = [
        'Name',
        'MaxGyro',
        'MaxGyro[]',
        'MaxHiG',
        'MaxHiG[]',
        'MaxAccel',
        'MaxAccel[]',
        '',
        'MaxAccel[-4]',
        'MaxAccel[-3]',
        'MaxAccel[-2]',
        'MaxAccel[-1]',
        'MaxAccel[ 0]',
        'MaxAccel[+1]',
        'MaxAccel[+2]',
        'MaxAccel[+3]',
        'MaxAccel[+4]',
        '',
        'Confidence',
        'ShotAccel',
        'ShotGyro',
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
        ws.write(self.Row.MaxGyro.value, col, data.maxGyro.v.magnitude)
        ws.write(self.Row.MaxGyroIndex.value, col, data.maxGyro.index)
        ws.write(self.Row.MaxHiG.value, col, data.maxHiG.v.magnitude)
        ws.write(self.Row.MaxHiGIndex.value, col, data.maxHiG.index)
        ws.write(self.Row.MaxAccel.value, col, data.maxAccel.v.magnitude)
        ws.write(self.Row.MaxAccelIndex.value, col, data.maxAccel.index)
        numSamples = len(data.accel)
        for i, j in enumerate(range(-4, 4 + 1)):
            index = data.maxAccel.index + j
            if index >= 0 and index < numSamples:
                ws.write(self.Row.MaxAccelRange.value + i, col, data.accel[index].magnitude)
        ws.write(self.Row.ShotConfidence.value, col, data.shotConfidence.value)
        shotIndex = data.shot.index
        ws.write(self.Row.ShotAccel.value, col, data.shot.v.magnitude)
        if (shotIndex < len(data.gyro)):
            ws.write(self.Row.ShotGyro.value, col, data.gyro[shotIndex].magnitude)
        ws.write(self.Row.ShotAccelIndex.value, col, shotIndex)
        for i, j in enumerate(range(-4, 4 + 1)):
            index = data.shot.index + j
            if index >= 0 and index < numSamples:
                ws.write(self.Row.ShotAccelRange.value + i, col, data.accel[index].magnitude)
        s.colIndex += 1
        
    def finalize(self):
        self.wb.close()