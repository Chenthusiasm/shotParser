# === IMPORTS ==================================================================
import enum
import glob
import math
import operator
import os
import shot
import shotOutput
import shutil
import string
import typing
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================

MAG_SPAN_VALUES = (-4, -3, -2, -1, 0, 1, 2, 3, 4)
MAX_SPAN_SIZE = len(MAG_SPAN_VALUES)
MAX_COL = 1000
NUM_LETTERS = len(string.ascii_uppercase)

SHOT_CONFIDENCE = (
    '0_Reset',
    '1_NoShot',
    '2_VeryLow',
    '3_Low',
    '4_Medium',
    '5_High',
    '6_VeryHigh',
)

COPY_PATHS = (
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.Reset.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.NoShot.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.VeryLow.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.Low.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.Medium.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.High.value]),
    os.path.join(os.getcwd(), SHOT_CONFIDENCE[shot.ShotConfidence.VeryHigh.value]),
)


# === ENUM =====================================================================

class TableEntry(enum.Enum):
    Offset = 0
    Value = 1
NUM_TABLE_ENTRIES = len(TableEntry)

    
class MagColOffset(enum.Enum):
    Index = 0
    Data = 1
NUM_MAG_COL_OFFSETS = len(MagColOffset)


class MagRowOffset(enum.Enum):
    Header = 0
    Max = 1
    IndexOfMax = 2
    MaxSpan = 3
    HeaderRepeat = MaxSpan + MAX_SPAN_SIZE
    Data = HeaderRepeat + 1
NUM_MAG_ROW_OFFSETS = len(MagRowOffset)


class ColOffset(enum.Enum):
    Index = 0
    X_LSB = 1
    Y_LSB = 2
    Z_LSB = 3
    Mag_LSB = 4
    X_g = 6
    Y_g = 7
    Z_g = 8
    Mag_g = 9
NUM_COL_OFFSETS = len(ColOffset)


class RowOffset(enum.Enum):
    Header = 0
    Data = 1
NUM_ROW_OFFSETS = len(RowOffset)
    

# === DICTIONARIES =============================================================

DATA_HEADER_TABLE = (
    (ColOffset.Index.value, 'index'),
    (ColOffset.X_LSB.value, 'x (lsb)'),
    (ColOffset.Y_LSB.value, 'y (lsb)'),
    (ColOffset.Z_LSB.value, 'z (lsb)'),
    (ColOffset.Mag_LSB.value, 'mag (lsb)'),
    (ColOffset.X_g.value, 'x (g)'),
    (ColOffset.Y_g.value, 'y (g)'),
    (ColOffset.Z_g.value, 'z (g)'),
    (ColOffset.Mag_g.value, 'mag (g)'),
    )

MAG_FIRST_COL_TABLE = (
    (MagRowOffset.Header.value, 'NAME'),
    (MagRowOffset.Max.value, 'MAX'),
    (MagRowOffset.IndexOfMax.value, '[MAX]'),
    (MagRowOffset.HeaderRepeat.value, 'NAME'),
    )


# === FUNCTIONS ================================================================

def getXlsxColStr(col : int) -> str:
    pre : int = int(col / NUM_LETTERS)
    post : int = int(col % NUM_LETTERS)
    preChar : str = ''
    if (pre > NUM_LETTERS):
        pre = NUM_LETTERS
    if (pre > 0):
        preChar = string.ascii_uppercase[pre - 1]
    postChar : str = string.ascii_uppercase[post]
    return preChar + postChar
    

def writeMagFirstColHeader(ws : xlsxwriter.workbook.Worksheet):
    for entry in MAG_FIRST_COL_TABLE:
        ws.write(entry[TableEntry.Offset.value], MagColOffset.Index.value, entry[TableEntry.Value.value])
        
    for i, spanValue in enumerate(MAG_SPAN_VALUES):
        ws.write(MagRowOffset.MaxSpan.value + i, MagColOffset.Index.value, spanValue)
        
        
def writeMagHeader(ws : xlsxwriter.workbook.Worksheet, nFile : int, name : str):
    col : int = MagColOffset.Data.value + nFile
    colStr : str = getXlsxColStr(col)
    dataRow : int = MagRowOffset.Data.value
    firstColStr : str = getXlsxColStr(MagColOffset.Index.value)
    ws.write(MagRowOffset.Header.value, col, name)
    ws.write(MagRowOffset.HeaderRepeat.value, col, name)
    ws.write_formula(MagRowOffset.Max.value, col, '=MAX({0}${1}:{0}${2})'.format(colStr, dataRow + 1, MAX_COL))
    ws.write_formula(MagRowOffset.IndexOfMax.value, col, '=MATCH({0}${1},{0}${2}:{0}${3},0)'.format(colStr, MagRowOffset.Max.value + 1, dataRow + 1, MAX_COL))
    for i, spanValue in enumerate(MAG_SPAN_VALUES):
        row = MagRowOffset.MaxSpan.value + i
        ws.write_formula(row, col, '=INDEX({0}${1}:{0}${2},{0}${3}+${4}{5})'.format(colStr, dataRow + 1, MAX_COL, MagRowOffset.IndexOfMax.value + 1, firstColStr, row + 1))
    
    
def writeMagData(ws : xlsxwriter.workbook.Worksheet, nFile : int, n : int, mag : float):
    row : int = MagRowOffset.Data.value + n
    col : int = MagColOffset.Data.value + nFile
    ws.write(row, col, mag)


def writeDataHeader(ws : xlsxwriter.workbook.Worksheet):
    for entry in DATA_HEADER_TABLE:
        ws.write(RowOffset.Header.value, entry[TableEntry.Offset.value], entry[TableEntry.Value.value])
        
        
def writeData(ws : xlsxwriter.workbook.Worksheet, n : int, v : shot.vector):
    row : int = RowOffset.Data.value + n
    ws.write(row, DATA_HEADER_TABLE[0][TableEntry.Offset.value], n)
    ws.write(row, DATA_HEADER_TABLE[1][TableEntry.Offset.value], v.x)
    ws.write(row, DATA_HEADER_TABLE[2][TableEntry.Offset.value], v.y)
    ws.write(row, DATA_HEADER_TABLE[3][TableEntry.Offset.value], v.z)
    ws.write(row, DATA_HEADER_TABLE[4][TableEntry.Offset.value], v.magnitude)
    ws.write(row, DATA_HEADER_TABLE[5][TableEntry.Offset.value], v.xG)
    ws.write(row, DATA_HEADER_TABLE[6][TableEntry.Offset.value], v.yG)
    ws.write(row, DATA_HEADER_TABLE[7][TableEntry.Offset.value], v.zG)
    ws.write(row, DATA_HEADER_TABLE[8][TableEntry.Offset.value], v.magnitudeG)
    
    
def process1():
    wb = xlsxwriter.Workbook(XLSX_WORKBOOK)
    ws_mag = wb.add_worksheet("mag")
    writeMagFirstColHeader(ws_mag)
    ws_magg = wb.add_worksheet("mag (g)")
    writeMagFirstColHeader(ws_magg)
    nFile = 0
    for fileName in glob.glob('*.csv'):
        print("processing {0}.".format(fileName))
        name = fileName.replace(".csv", "")
        filePath = os.path.join(os.getcwd(), fileName)
        with open(filePath, 'r') as file:
            readLines = file.readlines()
        file.close()
        ws = wb.add_worksheet(name)
        writeDataHeader(ws)
        writeMagHeader(ws_mag, nFile, name)
        writeMagHeader(ws_magg, nFile, name)
        nLine = 0
        for line in readLines:
            line = line.strip()
            entries = line.split(',')
            type = int(entries[0])
            if (type == 2):
                v = shot.vector(int(entries[1]), int(entries[2]), int(entries[3]))
                writeData(ws, nLine, v)
                writeMagData(ws_mag, nFile, nLine, v.magnitude)
                writeMagData(ws_magg, nFile, nLine, v.magnitudeG)
                nLine += 1
        nFile += 1
    wb.close()
    
def process2():
    output = shotOutput.xlsx()
    for path in COPY_PATHS:
        try:
            os.makedirs(path)
        except OSError as error:
            print('{0} already exists'.format(path))
    data = []
    for fileName in glob.glob('*.csv'):
        print('processing {0}...'.format(fileName))
        datum = shot.data(fileName)
        confidenceIndex = datum.shotConfidence.value
        copyPath = COPY_PATHS[confidenceIndex]
        shutil.copy2(datum.filePath, copyPath)
        output.writeShotData(datum)
        data.append(datum)
    #data.sort(key=operator.attrgetter('maxAccel.magnitude'))
    #data.sort(key=operator.attrgetter('shotMagnitude'))
    output.finalize()
    

# === MAIN =====================================================================

if __name__ == "__main__":
    process2()
else:
    print("ERROR: shotParser needs to be the calling python module!")