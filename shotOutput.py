# === IMPORTS ==================================================================

import enum
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
    V = Name + 1
    VRange = V + VECTOR_OFFSET_LENGTH + 1
    X = VRange + RANGE_LENGTH + 1
    Y = X + VECTOR_OFFSET_LENGTH + 1
    Z = Y + VECTOR_OFFSET_LENGTH + 1
    Shot = Z + VECTOR_OFFSET_LENGTH + 1
    ShotConfidence = Shot + VECTOR_OFFSET_LENGTH
    ShotRange = 1 + ShotConfidence + 1


# === CLASSES ==================================================================

class xlsx:
    __EXTENSION = '.xlsx'
    
    __DEFAULT_COLUMN_WIDTH = 20
    
    __DEFAULT_SHEET_NAMES = (
        'Reset',
        'No Shot',
        'Very Low',
        'Low',
        'Medium',
        'High',
        'Very High',
    )
    
    __HEADERS = [
        'NAME',
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
        self.sheets: typing.List[self.sheet] = []
        for name in sheetNames:
            s = self.sheet(name, self.wb.add_worksheet(name))
            for i, field in enumerate(self.__HEADERS):
                s.ws.write(s.row, i, field)
            # Set the column width.
            s.ws.set_column(Row.Header.value, Row.Header.value, self.__DEFAULT_COLUMN_WIDTH)
            # Freeze the header rows and columns.
            s.ws.freeze_panes(Row.Data.value, Col.V.value)
            s.row += 1
            self.sheets.append(s)
            
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
    
    def writeShotData(self, data : shot.data):
        s : self.sheet = self.sheets[data.shotConfidence.value]
        row = s.row
        ws : xlsxwriter.Workbook.worksheet_class = s.ws
        ws.write(row, Col.Name.value, data.name)
        self.__writeVectorDatum(ws, row, Col.V.value, data.maxAccel)
        self.__writeRange(ws, row, Col.VRange.value, data.accel, data.maxAccel.index)
        self.__writeVectorDatum(ws, row, Col.X.value, data.maxAccelX)
        self.__writeVectorDatum(ws, row, Col.Y.value, data.maxAccelY)
        self.__writeVectorDatum(ws, row, Col.Z.value, data.maxAccelZ)
        self.__writeVectorDatum(ws, row, Col.Shot.value, data.shot)
        ws.write(row, Col.ShotConfidence.value, data.shotConfidence.value)
        self.__writeRange(ws, row, Col.ShotRange.value, data.accel, data.shot.index)
        ws.write(row, Col.V.value + VectorOffset.Magnitude.value, data.maxAccel.v.magnitude)
        s.row += 1
        
    def finalize(self):
        self.wb.close()