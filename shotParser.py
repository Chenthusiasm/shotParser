# === IMPORTS ==================================================================
import enum
import glob
import math
import os
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================

MAG_SPAN_VALUES = (-4, -3, -2, -1, 0, 1, 2, 3, 4)
MAX_SPAN_SIZE = len(MAG_SPAN_VALUES)


# === ENUM =====================================================================

class TableEntry(enum.Enum):
    Offset = 0,
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
    Data = MaxSpan + MAX_SPAN_SIZE
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
    (ColOffset.Mag_g.value, 'mag (g)')
    )

MAG_FIRST_COL_TABLE = (
    (MagRowOffset.Max.value, 'MAX'),
    (MagRowOffset.IndexOfMax.value, 'MAX[idx]')
    )


# === FUNCTIONS ================================================================

def convertLsbToG(x : int) -> float:
    return float(x / 4096)


def writeMagFirstColHeader(ws : xlsxwriter.workbook.Worksheet):
    for entry in MAG_FIRST_COL_TABLE:
        ws.write(entry[TableEntry.Offset.value], MagColOffset.Index.value, entry[TableEntry.Value.value])
        
        
def writeMagHeader(ws : xlsxwriter.workbook.Worksheet, dataSet : int):
    print("stub")


def writeDataHeader(ws : xlsxwriter.workbook.Worksheet):
    for entry in DATA_HEADER_TABLE:
        ws.write(RowOffset.Header.value, entry[TableEntry.Offset.value], entry[TableEntry.Value.value])

    
def process():
    wb = xlsxwriter.Workbook('shotParser.xlsx')
    ws_mag = wb.add_worksheet("mag")
    ws_magg = wb.add_worksheet("mag (g)")
    mag_col = 0
    for filename in glob.glob('*.csv'):
        print("processing {0}.".format(filename))
        name = filename.replace(".csv", "")
        with open(os.path.join(os.getcwd(), filename), 'r') as file:
            readLines = file.readlines()
        file.close()
        ws = wb.add_worksheet(name)
        row = 0
        writeDataHeader(ws)
        ws_mag.write(row, mag_col, name)
        ws_magg.write(row, mag_col, name)
        for line in readLines:
            line = line.strip()
            entries = line.split(',')
            type = int(entries[0])
            if (type == 2):
                row += 1
                x = int(entries[1])
                y = int(entries[2])
                z = int(entries[3])
                magnitude = math.sqrt(x*x + y*y + z*z)
                magnitudeG = convertLsbToG(magnitude)
                ws.write(row, 0, row - 1)
                ws.write(row, 1, x)
                ws.write(row, 2, y)
                ws.write(row, 3, z)
                ws.write(row, 4, magnitude)
                ws.write(row, 6, convertLsbToG(x))
                ws.write(row, 7, convertLsbToG(y))
                ws.write(row, 8, convertLsbToG(z))
                ws.write(row, 9, magnitudeG)
                ws_mag.write(row, mag_col, magnitude)
                ws_magg.write(row, mag_col, magnitudeG)
        mag_col += 1
    wb.close()

# === MAIN =====================================================================

if __name__ == "__main__":
    process()
else:
    print("ERROR: shotParser needs to be the calling python module!")