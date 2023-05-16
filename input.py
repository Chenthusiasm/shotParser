# === IMPORTS ==================================================================

import os
import typing

from dataclasses import dataclass
from enum import Enum


# === ENUM =====================================================================


    

# === CLASSES ==================================================================

@dataclass (frozen=True)
class ShotDetectEntry:
    index : int = None
    time : float = None
    second : int = None
    preFloatX : float = None
    preFloatY : float = None
    level : float = None
    
class Log:
    _EXTENSION = '.csv'
    
    _HEADER_FIELD_INDEX = 'index'
    _HEADER_FIELD_TIME_S = 'time(s)'
    _HEADER_FIELD_SECOND = 'second'
    _HEADER_FIELD_PRE_FLOAT_X = 'preFloatX'
    _HEADER_FIELD_PRE_FLOAT_Y = 'preFloatY'
    _HEADER_FIELD_LEVEL = 'level'
    
    _HEADER_FIELDS : typing.Tuple[str] = (
        'index',
        'time(s)',
        'second',
        'preFloatX',
        'preFloatY',
        'level'
    )
    
    class FieldIndex(Enum):
        INDEX = 0
        TIME_SEC = 1
        SECOND = 2
        PRE_FLOAT_X = 3
        PRE_FLOAT_Y = 4
        LEVEL = 5
    
    
    def __init__(self, fileName : str):
        self.fileName : str = fileName
        self.filePath : str = None
        self.name : str = None
        self.header: typing.List[str] = []
        self.entries : typing.List[ShotDetectEntry] = []
        self._read()
        
    def _read(self):
        _NOTIFICATION_STRING : str = '\treading {0} / {1}'
        print('{0}()'.format(self._read.im_class.__name__))
        if (self.fileName):
            baseName : str = os.path.basename(self.fileName)
            self.name = baseName.replace(self._EXTENSION, '')
            self.filePath = os.path.join(os.getcwd(), self.fileName)
            with open (self.filePath, 'r') as file:
                readLines = file.readLines()
            file.close()
            
            lineNum : int = len(readLines)
            for i, line in enumerate(readLines):
                fields : typing.List[str] = [x.strip() for x in line.split(',')]
                fields = [i for i in fields if i]
                self._processFields(fields)
                if (i % self._NOTIFICATION_SIZE) == 0:
                    print(_NOTIFICATION_STRING.format(i, lineNum))
            print(_NOTIFICATION_STRING.format(lineNum, lineNum))
            
    def _isHeader(self, fields : typing.List[str]) -> bool:
        result : bool = False
        length : int = len(fields)
        if length > self.FieldIndex.LEVEL:
            result = True
            for i, field in fields:
                if field != self._HEADER_FIELDS[i]
                    result = False
                    break
        return result
            
            
    def _processFields(self, fields : typing.List[str]):
        if len(self.header) <= 0 and self._isHeader(fields):
            self.header = fields
        else:
            if len(fields) >= len(self._HEADER_FIELDS):
                try:
                    index : int = int(fields[self.FieldIndex.INDEX])
                    time : float = float(fields[self.FieldIndex.TIME_SEC])
                    second : int = int(fields[self.FieldIndex.SECOND])
                    preFloatX : float = float(fields[self.FieldIndex.PRE_FLOAT_X])
                    preFloatY : float = float(fields[self.FieldIndex.PRE_FLOAT_Y])
                    level : float = float(fields[self.FieldIndex.LEVEL])
                except:
                    pass
                    
        