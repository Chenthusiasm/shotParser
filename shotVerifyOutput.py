# === IMPORTS ==================================================================

import enum
from io import FileIO
import os
import shot
import string
import typing
import xlsxwriter


# === GLOBAL CONSTANTS =========================================================

DEFAULT_FILE_NAME = "shotVerify"


# === ENUM =====================================================================
    
class Col(enum.Enum):
    FileName = 0
    CountDebug = 1
    CountImuGyro = 2
    CountImuAccel = 3
    CountHiGAccel = 4
    CountCalibrationAccel = 5
    CountSettings = 6
    CountCalibrationGyro = 7
    CountHiGAccelCompressed = 8
    CountInvalid = 9


# === CLASSES ==================================================================

class shotVerifyXlsx:
    __EXTENSION = '.xlsx'
    
    __SHEET_NAME = 'counts'
    
    __DEFAULT_COLUMN_WIDTH = 20
    
    __HEADER_LABELS : typing.List[str] = (
        'FILENAME',
        'DEBUG [15]',
        'IMU GYRO [1]',
        'IMU ACCEL [2]',
        'HI-G ACCEL [3]',
        'CAL ACCEL [4]',
        'HANDINESS [5]',
        'CAL GYRO [6]',
        'HI-G ACCEL COMP [7]',
        'INVALID'
    )
    __HEADER_LABELS_LENGTH = len(__HEADER_LABELS)
    
    __DATA_ROW_START = 1
        
    def __init__(self, fileName: str = DEFAULT_FILE_NAME):
        self.fileName: str = str(fileName)
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.fileName + self.__EXTENSION)
        self.sheet: xlsxwriter.Workbook.worksheet_class = self.wb.add_worksheet(self.__SHEET_NAME)
        self.sheetRow: int = 0
        self.__initSheet()
        
    def __initSheet(self):
        for i, field in enumerate(self.__HEADER_LABELS):
            self.sheet.write(self.sheetRow, i, field)
        self.sheetRow += 1
    
    def writeShotData(self, data : shot.data):
        self.sheet.write(self.sheetRow, Col.FileName.value, data.fileName)
        self.sheet.write_number(self.sheetRow, Col.CountDebug.value, data.counts[0])
        self.sheet.write_number(self.sheetRow, Col.CountImuGyro.value, data.counts[1])
        self.sheet.write_number(self.sheetRow, Col.CountImuAccel.value, data.counts[2])
        self.sheet.write_number(self.sheetRow, Col.CountHiGAccel.value, data.counts[3])
        self.sheet.write_number(self.sheetRow, Col.CountCalibrationAccel.value, data.counts[4])
        self.sheet.write_number(self.sheetRow, Col.CountSettings.value, data.counts[5])
        self.sheet.write_number(self.sheetRow, Col.CountCalibrationGyro.value, data.counts[6])
        self.sheet.write_number(self.sheetRow, Col.CountHiGAccelCompressed.value, data.counts[7])
        self.sheet.write_number(self.sheetRow, Col.CountInvalid.value, data.counts[8])
        self.sheetRow += 1
        
    def finalize(self):
        self.wb.close()
    