# Well Log Data Tools

This python script offers two different conversions for well log data:
1) Conversion from a BoreholeLASLogData.xls model to a LAS 2.0 file. 
2) Conversion from a LAS 2.0 file to an entry in the WellLog Content Model.

### This repository contains the following files:
- BoreholeLASLogData.xls: Excel spreadsheet for inputting Well Log data
- LAS20_Standards.txt: LAS 2.0 Standards
- LAS_3_File_Structure.PDF: LAS 3.0 Standards
- SPWLA_LogTypeClass.xlsx: SPWLA Log Type Classes
- WellLogDataConverter.pyw: Python code for the tools
- WellLogsTemplate.xls: WellLogs Content Model Template

#### WellLogDataConverter Tool Requirements:
- [Python 2.6x] (http://www.python.org/)
- [xlrd] (http://pypi.python.org/pypi/xlrd)
- [xlutils] (http://pypi.python.org/pypi/xlutils)

#### To Run the WellLogDataConverter Tool:
- Double click WellLogDataConverter.pyw to run the script.
- Choose the type of conversion.
- Choose to convert either a single file or convert all files in a folder.
- Click Start Conversion!
- Output Locations:
-- The created LAS file(s) are written to the same folder as the original Excel file(s).
-- The created WellLogs Content Model Excel file (WellLogs.xls) is written to the same folder as the original LAS file(s).

#### Troubleshooting:
- Make sure the dependencies for the tool are installed properly.

After Python is installed it needs to be added to the system's Environment Variables (Win7 Directions):

1. Go to the Start Menu.
2. Right Click "Computer".
3. Select "Properties".
4. A dialog should pop up with a link on the left called "Advanced system settings". Click it.
5. In the System Properties dialog, click the button called "Environment Variables".
6. In the Environment Variables dialog look for "Path" under the System Variables window.
7. Add "C:\Python26\ArcGIS10.0\" (or wherever python.exe is) to the end of it. The semicolon is the path separator on windows.
8. Click Ok and close the dialogs.

Installing the xlrd and xlutils modules (Win7 Directions):

1. Download the modules from the links above.
2. Completely unpack the modules.
3. Hold down Shift and right click on the folder which setup.py is located.
4. Click ‘Open command window here’.
5. Type ‘python setup.py install’ in the command window.
6. Do this for each module.
