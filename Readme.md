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