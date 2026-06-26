from io import BytesIO
from openpyxl import load_workbook
from typing import Dict, Any

from config import (
    EXCEL_COMMON_FIELDS,
    RESOURCE_START_ROW,
    SPRINT_CYCLE_COLUMN,
    ALLOCATION_COLUMN,
    EXCEPTIONAL_RATE_COLUMN
)

class ExcelReader:

    def __init__(self, file):
        # Read raw bytes safely if a stream object is passed
        file_bytes = file.read() if hasattr(file, "read") else file
        
        # Convert explicit memoryview blocks into traditional bytes
        if isinstance(file_bytes, memoryview):
            file_bytes = file_bytes.tobytes()
            
        # Load cleanly via a BytesIO wrapper stream
        self.workbook = load_workbook(BytesIO(file_bytes))
        self.sheet = self.workbook.worksheets[0]

    def read_common_fields(self) -> dict[str, Any]:
        data = {}
        for field, cell in EXCEL_COMMON_FIELDS.items():
            data[field] = self.sheet[cell].value
        return data 

    def read_resource_fields(self) -> dict[str, Any]:
        return {
            "Sprint Cycle": self.sheet[f"{SPRINT_CYCLE_COLUMN}{RESOURCE_START_ROW}"].value,
            "Allocation": self.sheet[f"{ALLOCATION_COLUMN}{RESOURCE_START_ROW}"].value,
            "Exceptional Rate": self.sheet[f"{EXCEPTIONAL_RATE_COLUMN}{RESOURCE_START_ROW}"].value
        }

    def read(self) -> dict[str, Any]:
        data = {}
        data.update(self.read_common_fields())
        data.update(self.read_resource_fields())
        return data 

def read_excel(file) -> dict[str, Any]:
    reader = ExcelReader(file)
    return reader.read()
