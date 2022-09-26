# === IMPORTS ==================================================================

import enum
from re import X
import shot
import string
import typing
import xlsxwriter



# === GLOBAL CONSTANTS =========================================================

DEFAULT_FILE_NAME = "shotParser"

RANGE = [ -4, -3, -2, -1, 0, 1, 2, 3, 4 ]
RANGE_LENGTH = len(RANGE)


# === ENUM =====================================================================

class Row(enum.Enum):
    Header = 0
    Data = 1
    
class VectorOffset(enum.Enum):
    Magnitude = 0
    Index = 1
    X = 2
    Y = 3
    Z = 4
VECTOR_OFFSET_LENGTH = len(VectorOffset)
    
class Col(enum.Enum):
    Name = 0
    Samples = Name + 1
    V = Samples + 1
    VRange = V + VECTOR_OFFSET_LENGTH + 1
    X = VRange + RANGE_LENGTH + 1
    Y = X + VECTOR_OFFSET_LENGTH + 1
    Z = Y + VECTOR_OFFSET_LENGTH + 1
    Shot = Z + VECTOR_OFFSET_LENGTH + 1
    ShotConfidence = Shot + VECTOR_OFFSET_LENGTH
    ShotRange = 1 + ShotConfidence + 1
    AltShot = ShotRange + RANGE_LENGTH + 1
    AltShotConfidence = AltShot + VECTOR_OFFSET_LENGTH
    AltShotRange = 1 + AltShotConfidence + 1


# === CLASSES ==================================================================

class xlsx:
    __EXTENSION = '.xlsx'
    
    __ALL_SHEET = 'ALL'
    
    __DEFAULT_COLUMN_WIDTH = 20
    
    __DEFAULT_SHEET_NAMES = (
        'No Shot',
        'Very Low',
        'Low',
        'Medium',
        'High',
        'Very High',
    )
    
    __HEADERS = [
        'Name',
        'Samples',
        'V',
        'V[]',
        'V-x',
        'V-y',
        'V-z',
        '',
        'V[-4]',
        'V[-3]',
        'V[-2]',
        'V[-1]',
        'V[ 0]',
        'V[+1]',
        'V[+2]',
        'V[+3]',
        'V[+4]',
        '',
        '|X|',
        '|X|[]',
        '|X|-x',
        '|X|-y',
        '|X|-z',
        '',
        '|Y|',
        '|Y|[]',
        '|Y|-x',
        '|Y|-y',
        '|Y|-z',
        '',
        '|Z|',
        '|Z|[]',
        '|Z|-x',
        '|Z|-y',
        '|Z|-z',
        '',
        'Shot',
        'Shot[]',
        'Shot-x',
        'Shot-y',
        'Shot-z',
        'Confidence',
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
        '',
        'Alt',
        'Alt[]',
        'Alt-x',
        'Alt-y',
        'Alt-z',
        'Confidence',
        '',
        'Alt[-4]',
        'Alt[-3]',
        'Alt[-2]',
        'Alt[-1]',
        'Alt[ 0]',
        'Alt[+1]',
        'Alt[+2]',
        'Alt[+3]',
        'Alt[+4]',
        
    ]
    __HEADERS_LENGTH = len(__HEADERS)
        
    class sheet:
        def __init__(self, name: str, ws : xlsxwriter.Workbook.worksheet_class):
            self.name : str = name
            self.row : int = 0
            self.ws: xlsxwriter.Workbook.worksheet_class = ws
        
    def __init__(self, fileName: str = DEFAULT_FILE_NAME, sheetNames: typing.List[str] = __DEFAULT_SHEET_NAMES):
        self.fileName: str = str(fileName)
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.fileName + self.__EXTENSION)
        self.rankedSheets: typing.List[self.sheet] = []
        self.allSheet: self.sheet = self.sheet(self.__ALL_SHEET, self.wb.add_worksheet(self.__ALL_SHEET))
        self.__initSheet(self.allSheet)
        for name in sheetNames:
            s = self.sheet(name, self.wb.add_worksheet(name))
            self.__initSheet(s)
            self.rankedSheets.append(s)
            
    def __initSheet(self, s : sheet):
        for i, field in enumerate(self.__HEADERS):
            s.ws.write(s.row, i, field)
        # Set the column width.
        s.ws.set_column(Row.Header.value, Row.Header.value, self.__DEFAULT_COLUMN_WIDTH)
        # Freeze the header rows and columns.
        s.ws.freeze_panes(Row.Data.value, Col.Samples.value)
        s.row += 1
            
    def __writeVectorDatum(self, ws : xlsxwriter.Workbook.worksheet_class, row : int, col :int, datum : shot.vectorDatum, ) -> int:
        ws.write(row, col + VectorOffset.Magnitude.value, datum.v.magnitude)
        ws.write(row, col + VectorOffset.Index.value, datum.index)
        ws.write(row, col + VectorOffset.X.value, datum.v.x)
        ws.write(row, col + VectorOffset.Y.value, datum.v.y)
        ws.write(row, col + VectorOffset.Z.value, datum.v.z)
        return col + VECTOR_OFFSET_LENGTH
    
    def __writeRange(self, ws : xlsxwriter.Workbook.worksheet_class, row : int, col :int, accel : typing.List[shot.vector], index : int) -> int:
        samples = len(accel)
        for i, j in enumerate(RANGE):
            j += index
            if (j >= 0) and (j < samples):
                ws.write(row, col + i, accel[j].magnitude)
        return col + RANGE_LENGTH
    
    def __writeShotData(self, s : sheet, data : shot.data):
        row: int = s.row
        ws : xlsxwriter.Workbook.worksheet_class = s.ws
        ws.write(row, Col.Name.value, data.name)
        ws.write(row, Col.Samples.value, len(data.accel))
        self.__writeVectorDatum(ws, row, Col.V.value, data.maxAccel)
        self.__writeRange(ws, row, Col.VRange.value, data.accel, data.maxAccel.index)
        self.__writeVectorDatum(ws, row, Col.X.value, data.maxAccelX)
        self.__writeVectorDatum(ws, row, Col.Y.value, data.maxAccelY)
        self.__writeVectorDatum(ws, row, Col.Z.value, data.maxAccelZ)
        self.__writeVectorDatum(ws, row, Col.Shot.value, data.shot.datum)
        ws.write(row, Col.ShotConfidence.value, data.shot.confidence.value)
        self.__writeRange(ws, row, Col.ShotRange.value, data.accel, data.shot.datum.index)
        self.__writeVectorDatum(ws, row, Col.AltShot.value, data.altShot.datum)
        ws.write(row, Col.AltShotConfidence.value, data.altShot.confidence.value)
        self.__writeRange(ws, row, Col.AltShotRange.value, data.accel, data.altShot.datum.index)
        s.row += 1
    
    def writeShotData(self, data : shot.data):
        self.__writeShotData(self.rankedSheets[data.shot.confidence.value], data)
        self.__writeShotData(self.allSheet, data)
        
    def finalize(self):
        self.wb.close()