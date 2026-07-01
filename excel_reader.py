from io import BytesIO
import openpyxl
from config import EXCEL_COMMON_FIELDS

def read_excel(file_bytes):
    if isinstance(file_bytes, memoryview):
        file_bytes = file_bytes.tobytes()
        
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    target_sheet = "Team Loading Sheet"
    if target_sheet in wb.sheetnames:
        sheet = wb[target_sheet]
    else:
        sheet = wb.worksheets[0]
        
    data = {}
    for field, cell_id in EXCEL_COMMON_FIELDS.items():
        data[field] = sheet[cell_id].value
        
    return data
