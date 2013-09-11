# Well Log Data Tools

Tools for managing well log data.

### This repository contains the following:
- BoreholeLASLogData.xls: Excel spreadsheet for inputting Well Log data
- LAS20_Standards.txt: LAS 2.0 Standards
- LAS_3_File_Structure.PDF: LAS 3.0 Standards
- SPWLA_LogTypeClass.xlsx: SPWLA Log Type Classes
- WellLogDataConverter Tool: Tool for converting the data inputted into the BoreholeLASLog content model to a LAS 2.0 formatted file.

#### WellLogDataConverter Tool Requirements:
- [Python 2.6x] (http://www.python.org/)
- [xlrd] (https://github.com/python-excel)

#### To Run the WellLogDataConverter Tool:
- Double click WellLogDataConverter.pyw to run the script.
- Choose to convert either a single Excel file in the BoreholeLASLog content model or convert all the Excel files in a folder.
- Created LAS files are written to the same folder as the original Excel file.