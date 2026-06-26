# Automated Field Updater Tool

A powerful cross-document data synchronization solution that streamlines the management and updating of enterprise documents (Word and Excel) with consistent field values and data.

## 📋 Overview

The **Automated Field Updater Tool** is a web-based application built with Streamlit that enables organizations to maintain consistency across multiple document types by:

- Extracting field values from Word (DOCX) and Excel (XLSX) template documents
- Allowing centralized editing of shared parameters across documents
- Automatically synchronizing updated values back to both documents
- Generating a packaged bundle of updated documents for deployment

This tool is particularly useful for enterprises managing project documentation, statements of work (SOWs), resource allocation sheets, and other documents that require synchronized field values.

## ✨ Features

### 📄 Multi-Format Support
- **Word Documents (.docx)**: Extract and update custom fields, dates, costs, and milestone tables
- **Excel Spreadsheets (.xlsx)**: Modify resource allocation, sprint cycles, rates, and project parameters

### 🔄 Intelligent Field Extraction
- Automatically identifies key fields from both document types
- Extracts project metadata (dates, costs, team information)
- Parses complex structures like milestone tables in Word documents
- Flexible date parsing supporting multiple formats (DD-MMM-YYYY, YYYY-MM-DD, DD/MM/YYYY)

### 🎯 Unified Parameter Editing
- Single interface to edit shared parameters across both documents
- Organized into logical sections:
  - **Global Core Parameters**: Project dates, costs, approvals
  - **Excel-Only Operational Settings**: Sprint cycles, allocation percentages, rates
  - **Payment Milestones**: Dynamic table editor for milestone management

### 📦 Batch Processing
- Generate updated versions of both documents simultaneously
- Export as a convenient ZIP bundle containing both synchronized files
- Maintains document structure and formatting integrity

### 🎨 Enterprise UI/UX
- Clean, professional interface with step-based workflow visualization
- Intuitive progress indicators (Template Ingestion → Parameter Alignment → Export Bundle)
- Real-time validation and error reporting
- Mobile-responsive design

## 🚀 Getting Started

### Prerequisites
- Python 3.8+
- pip (Python package manager)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/aniket-cog/automated-field-updater.git
   cd automated-field-updater
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Running the Application

Start the Streamlit server:

```bash
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`

## 📖 Usage Guide

### Step 1: Template Ingestion
1. Upload your base **Statement of Work Document (.docx)** file
2. Upload your base **Team Loading Sheet (.xlsx)** file
3. Click "Analyze and Extract Data Components" to parse both documents

### Step 2: Parameter Alignment
The application extracts and displays all key fields organized into sections:

- **Global Core Parameters**
  - Project Start Date
  - Project End Date
  - Total Cost (in EUR)
  - Release Start/End Dates
  - SOW Approval details

- **Excel-Only Operational Settings**
  - Number of Sprint Cycles
  - Allocation Percentage (%)
  - Exceptional Rate (in EUR)

- **Payment Milestones Table**
  - Edit milestone names, dates, and payment breakdowns
  - Add or remove milestones dynamically

### Step 3: Generate & Export
1. Review all edited parameters
2. Click "Generate and Harmonize Documents"
3. Download the ZIP bundle containing:
   - `_Updated_SOW.docx` - Updated Word document
   - `_Updated_TLS.xlsx` - Updated Excel spreadsheet

## 🏗️ Architecture

### Core Modules

| Module | Purpose |
|--------|---------|
| **app.py** | Main Streamlit application and UI orchestration |
| **config.py** | Configuration constants mapping document field locations |
| **word_reader.py** | Extracts fields and tables from Word documents |
| **word_updater.py** | Updates Word document fields and milestones |
| **excel_reader.py** | Reads and parses Excel spreadsheet data |
| **excel_updater.py** | Updates Excel cells with new values |

### Data Flow

```
User Upload
    ↓
[word_reader.py] → Extract DOCX fields & tables
[excel_reader.py] → Extract XLSX fields
    ↓
Streamlit UI ← Display in editable form
    ↓
User Edits Parameters
    ↓
[word_updater.py] → Apply changes to DOCX
[excel_updater.py] → Apply changes to XLSX
    ↓
ZIP Bundle Generation
    ↓
User Download
```

## ⚙️ Configuration

Field mappings and document structure are configured in `config.py`:

- **WORD_FIELDS**: Maps field names to (row, column) coordinates in the Word document
- **EXCEL_COMMON_FIELDS**: Maps field names to Excel cell references (e.g., "C10")
- **MILESTONE_PARENT_ROW/COLUMN**: Location of milestone table in Word document
- **RESOURCE_START_ROW**: Starting row for resource entries in Excel

To adapt this tool to your document structure, update these mappings to match your template files.

## 📦 Dependencies

- **streamlit** (≥1.46.0) - Web application framework
- **python-docx** (≥1.1.2) - Word document manipulation
- **openpyxl** (≥3.1.5) - Excel spreadsheet handling
- **pandas** (≥2.2.2) - Data manipulation and table editing

## 🔒 License

This project is licensed under the terms specified in the [LICENSE](LICENSE) file.

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 Notes

- Ensure your template documents follow the expected structure defined in `config.py`
- Dates can be in multiple formats; the tool attempts to parse them intelligently
- The generated documents preserve the original formatting and structure
- ZIP exports are cleaned up after download to maintain server efficiency

## 🐛 Troubleshooting

**Issue**: "Error parsing templates"
- Ensure both files are in the correct format (.docx for Word, .xlsx for Excel)
- Verify the document structure matches the mappings in `config.py`

**Issue**: Field values not updating
- Check that the row/column mappings in `config.py` are correct for your template
- Ensure cells are not locked or protected in the original documents

**Issue**: Application crashes during processing
- Check the browser console and terminal for detailed error messages
- Verify all files have appropriate read/write permissions

## 📧 Support

For issues, questions, or suggestions, please open an issue on the [GitHub repository](https://github.com/aniket-cog/automated-field-updater/issues).

---

**Built with ❤️ by Aniket Saha**
