from io import BytesIO
from openpyxl import load_workbook

from config import (
    EXCEL_COMMON_FIELDS,
    TEAM_NAME_COLUMN,
    SPRINT_CYCLE_COLUMN,
    ALLOCATION_COLUMN,
    EXCEPTIONAL_RATE_COLUMN,
    RESOURCE_START_ROW
)

class ExcelUpdater:

    def __init__(self, file):
        file_bytes = file.read() if hasattr(file, "read") else file
        if isinstance(file_bytes, memoryview):
            file_bytes = file_bytes.tobytes()
        self.workbook = load_workbook(BytesIO(file_bytes))
        self.sheet = self.workbook.active

    def update_common_fields(self, common_fields):
        for field, cell in EXCEL_COMMON_FIELDS.items():
            if field in common_fields:
                self.sheet[cell] = common_fields[field]
                if field == "Total Cost":
                    self.sheet[cell].number_format = '€* #,##0.00'

    def update_resource_table(self, sprint_cycle, allocation, exceptional_rate):
        row = RESOURCE_START_ROW
        while True:
            team_name = self.sheet[f"{TEAM_NAME_COLUMN}{row}"].value
            if team_name is None or str(team_name).strip() == "":
                break
            self.sheet[f"{SPRINT_CYCLE_COLUMN}{row}"] = sprint_cycle
            self.sheet[f"{ALLOCATION_COLUMN}{row}"] = allocation
            self.sheet[f"{EXCEPTIONAL_RATE_COLUMN}{row}"] = exceptional_rate
            row += 1

    def update(self, common_fields, sprint_cycle, allocation, exceptional_rate):
        self.update_common_fields(common_fields)
        self.update_resource_table(sprint_cycle, allocation, exceptional_rate)

    def get_file(self):
        output = BytesIO()
        self.workbook.save(output)
        output.seek(0)
        return output

def update_excel(uploaded_file, common_fields, sprint_cycle, allocation, exceptional_rate):
    updater = ExcelUpdater(uploaded_file)
    updater.update(common_fields, sprint_cycle, allocation, exceptional_rate)
    return updater.get_file()
