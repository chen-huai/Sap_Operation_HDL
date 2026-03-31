# CLAUDE.md

请用中文回复，所有测试模块都放在test中
This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a SAP automation tool built with Python and PyQt5 that provides automated SAP order creation, data processing, and PDF file renaming capabilities. The application features a tabbed GUI interface for different operations.

## Key Dependencies

- **PyQt5**: GUI framework (5.15.11)
- **pandas**: Data processing (2.2.2)
- **win32com**: SAP GUI automation
- **PDF libraries**: pdfminer.six, pdfplumber, pypdfium2 for PDF processing
- **Excel libraries**: openpyxl for Excel file operations
- **PyInstaller**: For building executable files

## Main Application Structure

### Core Files
- `Sap_Operate.py`: Main application entry point with GUI logic and business operations
- `Sap_Operate_Ui.py`: Auto-generated PyQt UI code from Sap_Operate_Ui.ui
- `Sap_Operate_theme.py`: Themed version of the main application
- `build_with_pyinstaller.py`: Build script for creating executables

### Functional Modules
- `Get_Data.py`: Data processing from Excel/CSV files with field mapping
- `Sap_Function.py`: SAP GUI automation using win32com client
- `File_Operate.py`: File system operations and path management
- `PDF_Operate.py`: PDF processing and renaming functionality
- `Data_Table.py`: Table data handling and display
- `Excel_Field_Mapper.py`: Excel field mapping utilities
- `Logger.py`: Logging functionality

### UI Components
- `Table_Ui.py`: Additional table interface components
- `Sap_Operate_Ui.ui`: Qt Designer UI file
- `Table_Ui.ui`: Table-specific UI file

## Common Development Commands

### Running the Application
```bash
# Main application
python Sap_Operate.py

# Themed version
python Sap_Operate_theme.py
```

### Building Executable
```bash
# Using the build script
python build_with_pyinstaller.py

# Manual PyInstaller command
pyinstaller --onefile --windowed --clean --noconfirm --icon=Sap_Operate_Logo.ico Sap_Operate_theme.py
```


## Application Features

### Main Operations
- **SAP Order Creation**: Automated creation of orders in SAP system
- **Data Processing**: Split and merge data based on billing information
- **PDF Renaming**: Automatic PDF file renaming based on invoice data
- **Data Recovery**: Retrieve and ensure data integrity

### GUI Interface
- Tabbed interface with multiple operation sections
- File selection dialogs for data input
- Real-time logging and status display
- Configuration import/export functionality

## Configuration

The application generates a `config` folder on the desktop with `config_sap.csv` for user-specific settings. The configuration file contains parameters for SAP operations and data processing rules.

## Data Flow

1. **Data Input**: Excel/CSV files are loaded via `Get_Data.py`
2. **Data Processing**: Field mapping and transformation using `Excel_Field_Mapper.py`
3. **SAP Operations**: Automated GUI interactions via `Sap_Function.py`
4. **File Operations**: PDF processing and file management
5. **Output**: Results displayed in GUI and logged for audit

## SAP Integration

The application uses win32com to interact with SAP GUI Scripting engine. Key requirements:
- SAP GUI must be installed and running
- Scripting must be enabled in SAP GUI
- User must have appropriate SAP permissions

## File Naming Conventions

- Main executables: `Sap_Operate_theme.exe` (built version)
- UI files: `*_Ui.py` (generated), `*.ui` (Qt Designer source)
- Icon files: `*.ico` for application branding
- Build artifacts: `dist/`, `build/` directories

## Error Handling

The application includes comprehensive error handling with user feedback through the GUI. Key error scenarios include:
- SAP connection failures
- File access issues
- Data format problems
- GUI automation timeouts