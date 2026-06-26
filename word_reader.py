from io import BytesIO
from docx import Document
import pandas as pd

from config import (
    MAIN_TABLE_INDEX,
    WORD_FIELDS,
    DUPLICATE_COLUMNS,
    MILESTONE_HEADER_ROW,
    MILESTONE_PARENT_ROW,
    MILESTONE_COLUMNS,
    MILESTONE_PARENT_COLUMN,
    MILESTONE_TABLE_INDEX
)

class WordReader:

    def __init__(self, file):
        # Read raw bytes safely if a stream object is passed
        file_bytes = file.read() if hasattr(file, "read") else file
        
        # Convert explicit memoryview blocks into traditional bytes
        if isinstance(file_bytes, memoryview):
            file_bytes = file_bytes.tobytes()
            
        # Safely wrap with BytesIO before giving it to python-docx
        self.doc = Document(BytesIO(file_bytes))
        self.table = self.doc.tables[MAIN_TABLE_INDEX]

    def read_fields(self):
        data = {}
        for field, (row, col) in WORD_FIELDS.items():
            value = self.table.rows[row].cells[col].text.strip()
            data[field] = value
        return data

    def read_milestones(self):
        parent_cell = self.table.rows[MILESTONE_PARENT_ROW].cells[MILESTONE_PARENT_COLUMN]
        nested_table = parent_cell.tables[MILESTONE_TABLE_INDEX]
        milestones = []

        for row in nested_table.rows[MILESTONE_HEADER_ROW+1:]:
            values = [cell.text.strip() for cell in row.cells]
            if not any(values):
                continue
            
            milestone = {
                "name": values[MILESTONE_COLUMNS["name"]],
                "date": values[MILESTONE_COLUMNS["date"]],
                "monthly": values[MILESTONE_COLUMNS["monthly"]],
                "quality": values[MILESTONE_COLUMNS["quality"]],
                "invoice": values[MILESTONE_COLUMNS["invoice"]]
            }
            milestones.append(milestone)

        return pd.DataFrame(milestones)


def read_word(file):
    reader = WordReader(file)
    return {
        "fields": reader.read_fields(),
        "milestones": reader.read_milestones()
    }
