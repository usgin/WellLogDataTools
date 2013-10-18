'''
Well Log Data Tools
Created on Jun 12, 2013

@author: Jessica Good Alisdairi, AZGS jalisdairi@azgs.az.gov

This python script offers two different conversions for well log data:
1) Conversion from a BoreholeLASLogData.xls model to a LAS 2.0 file. 
2) Conversion from a LAS 2.0 file to an entry in the WellLog Content Model.

The BoreholeLASLogData.xls model to a LAS 2.0 file converter 
makes the following assumptions about the inputted Excel workbook:

The workbook tabs are named such that the tab name for: 
    ~Version Information Section     starts with     ~V
    ~Well Information Section        starts with     ~W
    ~Curve Information Section       starts with     ~C
    ~Parameter Information Section   starts with     ~P
    ~Other  Information Section      starts with     ~O
    ~ASCII Log Data Section          starts with     ~A
And there are no other tabs in the workbook that start with ~V, ~W, ~C, ~P, ~O or ~A.

The following assumptions are made about the ~V Excel worksheet:
    - The data starts in the 2nd row of the 1st column.
    
The following assumptions are made about the ~W Excel worksheet:
    - The 1st column contains the header "LAS Mnemonic" or "Mnemonic" in any row on the sheet and the data follows.
    - The description for values in the 1st column are, in this exact order: "START DEPTH", "STOP DEPTH", "STEP", "NULL VALUE", 
    "COMPANY", "WELL", "FIELD", "LOCATION", "COUNTY", "STATE", "COUNTRY", "SERVICE COMPANY", "DATE", "UNIQUE WELL ID", 
    "API NUMBER", "LOG URI", "LOG TYPE", "WELL TYPE", "LAT DEGREE", "LONG DEGREE", "SRS", "LOCATION UNCERTAINTY", "TOTAL DEPTH", 
    "ELEVATION GL", "LINK", "SOURCE", "NOTE"

The following assumptions are made about the ~C Excel worksheet:
    - The 1st column contains the header "LAS Mnemonic" or "Mnemonic" in any row on the sheet and the data follows.
    - If the 1st column contains the word "value" the data for that row will not be written.
    - If the 2nd column contains the word "units" the data for that row will not be written.
    - If the 5th column contains the word "example" the data for that row will not be written.
    
The following assumptions are made about the ~P Excel worksheet:
    - The 1st column contains the header "LAS Mnemonic" or "Mnemonic" in any row on the sheet and the data follows.
    
The following assumptions are made about the ~O Excel worksheet:
    - The data starts in the 3rd row of the 1st column.

The following assumptions are made about the ~A Excel worksheet:
    - The data headers are in the 3rd row.
    - If a data header cotains the words value and units that column is ignored.
    - The data starts in the 4th row.
    
'''

import os, glob, re
import datetime
try:
    import xlrd
except:
    print "Import of XLRD module failed.\nThe XLRD module can be downloaded from: http://pypi.python.org/pypi/xlrd"
try:
    import xlutils
except:
    print "Import of XLWT module failed.\nThe XLWT module can be downloaded from: http://pypi.python.org/pypi/xlutils"    
from xlutils.copy import copy
 
import Tkinter, tkFileDialog
from Tkinter import *
import tkMessageBox

# Main function for the Well Log Data Converter Tool
# Set up the user interface
def main(argv=None):
    root = Tkinter.Tk()
    root.title("Well Log Data Converter")
    root.minsize(0, 0)

    # Create the frame for the conversion type radio buttons
    labelframe = LabelFrame(root, text = "Conversion Type")
    labelframe.pack(fill = "both", expand = "yes")
    global cType, cOpt
    cType = IntVar()
    cOpt = IntVar()
    R1 = Radiobutton(labelframe, text = "BoreholeLASLogData Model  --->  LAS 2.0 text file", variable = cType, value = 1)
    R1.select()
    R1.pack(anchor = W)
    R2 = Radiobutton(labelframe, text = "LAS 2.0 text file  --->  WellLog Content Model", variable = cType, value = 2)
    R2.pack(anchor = W)
    
    # Create the frame for the conversion option radio buttons
    labelframe = LabelFrame(root, text = "Conversion Options")
    labelframe.pack(fill = "both", expand = "yes")
    R3 = Radiobutton(labelframe, text = "Convert a Single File", variable = cOpt, value = 1)
    R3.select()
    R3.pack(anchor = W)
    R4 = Radiobutton(labelframe, text = "Convert All Files in a Folder", variable = cOpt, value = 2)
    R4.pack(anchor = W)
    
    # Create the Start Conversion button to start the conversion
    B = Tkinter.Button(root, text = "Start Conversion!", command = GetFiles)
    B.pack()

    # Create the frame for the messages list
    global textFrame
    msgFrame = LabelFrame(root, text="Messages")
    textFrame = Tkinter.Text(msgFrame, height = 30)
    yscrollbar = Scrollbar(msgFrame)
    yscrollbar.pack(side = RIGHT, fill = Y)
    textFrame.pack(side = LEFT, fill = Y)
    yscrollbar.config(command = textFrame.yview)
    textFrame.config(yscrollcommand = yscrollbar.set)
    msgFrame.pack(fill = X, expand = "yes")

    # Create the frame for the exit button
    b3frame = Frame(root, bd = 5)
    b3frame.pack()
    B3 = Tkinter.Button(b3frame, text = "Exit", command = ExitConvert)
    B3.pack()
    
    root.mainloop()
    return

# Get the files to be converted from the user
def GetFiles():
    # Clear the text frame
    textFrame.delete("1.0", END)
    
    # If converting a single file
    if cOpt.get() == 1:
        # If converting from a BoreholeLASLogData Model to a LAS 2.0 file get the Excel file
        if cType.get() == 1:
            cFile = tkFileDialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls")])
        # If converting from a LAS 2.0 file to the WellLog Content Model get the LAS file
        else:
            cFile = tkFileDialog.askopenfilename(filetypes=[("LAS Files","*.las")])
        # If cancel was pressed
        if cFile == "":
            return
        cFiles = []
        cFiles.append(cFile)
        
    # If converting all files in a folder
    else:
        path = tkFileDialog.askdirectory()
        # If cancel was pressed 
        if path == "":
            return
        
        # If converting from a BoreholeLASLogData Model to a LAS 2.0 file get the Excel files
        if cType.get() == 1:
            cFiles = glob.glob(path + "/*.xls") + glob.glob(path + "/*.xlsx")
        # If converting from a LAS 2.0 file to the WellLog Content Model get the LAS files
        else:
            cFiles = glob.glob(path + "/*.las")
        
        # Check for duplicate file names (not including the file extension)
        if DupFileNames(cFiles):
            return
  
    Convert(cFiles)    
    return

# Check for duplicate file names (not including the file extension)
def DupFileNames(files):

    # Make a list of the file names
    fileNames = []
    for f in files:
        tempBaseFile = os.path.basename(f)
        tempFileName = os.path.splitext(tempBaseFile)[0]
        fileNames.append(tempFileName)
    
    # Find any duplicate file names (this could happen if extensions are different)
    if len(fileNames) != len(set(fileNames)):
        dupes = []
        for i in fileNames:
            if fileNames.count(i) > 1:
                if i not in dupes:
                    dupes.append(i)
        baseFile = ', '.join(dupes)
        
        # Duplicate file names were found
        if len(dupes) == 1:
            Message(baseFile + " is used more than once as a file name. Make sure all file names are unique.")
            return True
        else:
            Message(baseFile + " are used more than once as file names. Make sure all file names are unique.")
            return True

    # No duplicate file names were found
    return False

# Start the conversion
def Convert(files):

    # If converting from a LAS 2.0 file to the WellLog Content Model open the WellLogsTemplate file
    if cType.get() == 2:
        try:
            # Open the template and preserve the formatting of the file
            # !!!! The colors are not slightly different though !!!
            wbk = xlrd.open_workbook("WellLogsTemplate.xls", formatting_info = True)
            # Get the field names from the first row of the first sheet
            fields = wbk.sheet_by_index(0).row_values(0)
            # Determine the first empty row by counting the number of rows
            firstEmptyRow = wbk.sheet_by_index(0).nrows
            w = copy(wbk)
            sht = w.get_sheet(0)
        except:
            Message("Error: Can't find WellLogsTemplate.xls")
            Message("A copy of the Well Logs Content Model must be in the same folder as this converter.")
            return
    
    # For each inputted file            
    for f in files:
        baseFile = os.path.basename(f)
        fileName = os.path.splitext(baseFile)[0]
        path = os.path.splitext(os.path.dirname(f))[0]

        output = ""
        # If converting from a BoreholeLASLogData Model to a LAS 2.0 file
        if cType.get() == 1:
            # Create output file
            lasFile = path + '\\' + fileName + '.las'      
            fileOut = open(lasFile, 'w')                     
            fileOut.write(output)
            
            # Read the BoreholeLASLogData Model and convert
            output = ReadBoreholeLASLogData(f, output)
            
            # Write the output file
            fileOut.write(output)
            fileOut.close()
            if output == "":
                os.remove(lasFile)
        
        # If converting from a LAS 2.0 file to the WellLog Content Model
        else:
            # Read the LAS file and convert
            output = ReadLAS(f, output)
            
            # Write the output
            if output != "":              
                WriteWellLogsCM(output, fields, firstEmptyRow, sht)
                firstEmptyRow += 1
                Message(baseFile + ": Converted successfully.")
     
    # If converting from a LAS 2.0 file to the WellLog Content Model           
    if cType.get() == 2:
        try:
            # Save the output file
            if os.path.isfile(path + "\\" + "WellLogs.xls"):
                res = tkMessageBox.askokcancel("Overwrite", "WellLogs.xls already exists. Overwrite?")
                if res:
                    w.save(path + "\\" + "WellLogs.xls")
                    Message("WellLogs.xls saved in " + path)
                else:
                    Message("WellLogs.xls not saved.")
            else:
                w.save(path + "\\" + "WellLogs.xls")
                Message("WellLogs.xls saved in " + path) 
        except:
            Message("Unable to save WellLogs.xls. If it is already open, close it first.")

    return

# Exit the program
def ExitConvert():
    exit()
    return

# Read the BoreholeLASLogData Excel file
def ReadBoreholeLASLogData(inExcel, output):
    try:
        try:
            wb = xlrd.open_workbook(inExcel)
            wb._path = os.path.basename(inExcel)
        except:
            Message("Error: Unable to open workbook. Terminating.")
            raise Exception
        
        sheets = [sht.name for sht in wb.sheets()]
        
        shts = {"~V": None, "~W": None, "~C": None, "~P": None, "~O": None, "~A": None}
        
        # Get the names of the sheets containg LAS sections
        for sheet in sheets:
            if sheet.startswith("~V") == True:
                shts["~V"] = sheet
            if sheet.startswith("~W") == True:
                shts["~W"] = sheet
            if sheet.startswith("~C") == True:
                shts["~C"] = sheet
            if sheet.startswith("~P") == True:
                shts["~P"] = sheet
            if sheet.startswith("~O") == True:
                shts["~O"] = sheet
            if sheet.startswith("~A") == True:
                shts["~A"] = sheet   
    
        # Print an error if can't find the necessary sheets and end
        for key, value in shts.iteritems():
            if value is None:
                # Sheets ~P and ~O are optional
                if key != "~P" and key != "~O":
                    Message(wb._path + ": Can't find the sheet starting with " + key + ". Terminating.")
                    raise Exception
        
        output += GetVersionInfo(wb, shts["~V"])
        output += GetWellInfo(wb, shts["~W"])
        output += GetCurveInfo(wb, shts["~C"])
        if shts["~P"] != None:
            output += GetParameterInfo(wb, shts["~P"])
        if shts["~O"] != None:
            output += GetOtherInfo(wb, shts["~O"])
        output += GetAsciiLogData(wb, shts["~A"])
    
        Message(wb._path + ": Conversion Successful")
        
    except Exception:
        Message(wb._path + ": Conversion Failed")
        output = ""
    
    return output

# Read the Excel file and write the ~Version Information Section
def GetVersionInfo(wb, shtName):

    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # Read the first four columns of the Excel file
    versInfo = sht.col_values(0)
    
    # Data starts in the 3rd row
    dataStartRow = 2
    
    # Output the headers
    output = "~Version Information Section"
    output += "\n"
    
    for i in range(dataStartRow, len(versInfo)):
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            versInfo[i] = str(versInfo[i]).encode('ascii')
        except:
            Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception  
        
        # Remove trailing whitespace and check for empty string
        versInfo[i] = versInfo[i].rstrip()
        if versInfo[i] == "":
            Message(wb._path + ": Row " + str(i + 1) + " in the " + shtName + " sheet is missing data. Terminating.")
            raise Exception
        
        # Output data
        output += versInfo[i]
        output += "\n"
    
    return output   

# Read the Excel file and write the ~Well Information Section
def GetWellInfo(wb, shtName):

    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # Read the first and second columns of the Excel file
    mnems = sht.col_values(0)
    values = sht.col_values(1)
    types = sht.col_types(1)
    
    # Required fields
    desc = ["START DEPTH", "STOP DEPTH", "STEP", "NULL VALUE", "COMPANY", "WELL", "FIELD", "LOCATION", "COUNTY", "STATE", "COUNTRY", "SERVICE COMPANY", "DATE", "UNIQUE WELL ID", "API NUMBER"]
    
    # Fields required because they are in the content model
    desc.extend(["LOG URI", "LOG TYPE", "WELL TYPE", "LAT DEGREE", "LONG DEGREE", "SRS", "LOCATION UNCERTAINTY", "TOTAL DEPTH", "ELEVATION GL", "LINK", "SOURCE", "NOTE"])
    
    # Determine the row with the headers and the row with the data immediately follows
    try:
        headerRow = mnems.index("LAS Mnemonic")
    except:
        try:
            headerRow = mnems.index("Mnemonic")
        except:
            Message(wb._path + ": Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the " + shtName + " sheet. Terminating.")
            raise Exception
    dataStartRow = headerRow + 1
    
    # Output the headers
    output = "~Well Information Section"
    output += "\n"
    output += "#MNEM.UNIT      VALUE/NAME      DESCRIPTION"
    output += "\n"
    output += "#--------     --------------   ---------------------"
    output += "\n"
    rowsOutputed = 1
    
    # For each item in the list representing the first column
    for i in range(dataStartRow, len(mnems)):
        
        # If the type of the cell is 3 a date type is indicated
        if types[i] == 3:
            values[i] = ConvertToDate(values[i], wb)
            if values[i] == -1:
                Message(wb._path + ": Unrecognized date in row " + str(i + 1) + " of the " + shtName + ". Terminating.")
                raise Exception

        # Remove decimal and trailing zeros that were added on Excel import
        try:
            values[i] = float(values[i])
            if values[i] == int(values[i]):
                values[i] = '%d'%values[i]
        except:
            pass      
        
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            values[i] = str(values[i]).encode('ascii')
        except:
            Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception

        # Check for missing fields and values
        mnems[i] = mnems[i].rstrip()
        if mnems[i] == "":
            Message(wb._path + ": Row " + str(i + 1) + " must have a field name (Mnem) in the " + shtName + " sheet.")
            raise Exception
        values[i] = values[i].rstrip()

        # First 15 fields are required so fill in placeholder data if empty
        if values[i] == "" and i < 15 + dataStartRow:
            # First 4 fields (STRT, STOP, STEP, NULL) should have the value '-9999' if no data
            if i < 4 + dataStartRow:
                values[i] = '-9999'
            # The rest of the fields should have the value 'Missing' if no data, except for date (see below)
            else:
                values[i] = 'Missing'
            # Thirteenth field (DATE) should have the value 1/1/1900 00:00, if no data
            if i == 13 + dataStartRow - 1:
                values[i] = datetime.datetime(1900, 1, 1, 0, 0, 0)

        # Output the fields
        output += mnems[i]
        
        # Each Mnem value must be followed by a period so insert one if not already there
        if "." not in mnems[i]:
            output += "."
        
        # Output the values
        output += "          "
        output += str(values[i])
        output += "          "
        output += ":"
        output += " "
        
        # Output the corresponding description
        if i - dataStartRow < len(desc):
            output += desc[i - dataStartRow]
        else:
            output += ": "  
            output += mnems[i]
            
        output += "\n"
        rowsOutputed = rowsOutputed + 1
            
    return output 

# Read the Excel file and write the ~Well Information Section
def GetCurveInfo(wb, shtName):
   
    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # Read the first 5 columns of the Excel file
    mnems = sht.col_values(0)
    units = sht.col_values(1)
    apiCodes = sht.col_values(2)
    curveDescs = sht.col_values(3)
    exs = sht.col_values(4)

    # Determine the row with the headers and the row with the data immediately follows
    try:
        headerRow = mnems.index("LAS Mnemonic")
    except:
        try:
            headerRow = mnems.index("Mnemonic")
        except:
            Message(wb._path + ": Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the " + shtName + " sheet. Terminating.")
            raise Exception
    dataStartRow = headerRow + 1
    
    # Output the headers
    output = "~Curve Information Section"
    output += "\n"
    output += "#MNEM.UNIT           API CODE        Curve Description"
    output += "\n"
    output += "#----------        -----------    -------------------------------"
    output += "\n"
    rowsOutputed = 1
    
    # For each item in the list representing the first column
    for i in range(dataStartRow, len(mnems)):

        # Remove decimal and trailing zeros that were added on Excel import
        try:
            apiCodes[i] = float(apiCodes[i])
            if apiCodes[i] == int(apiCodes[i]):
                apiCodes[i] = '%d'%apiCodes[i]
        except:
            pass   

        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            units[i] = str(units[i]).encode('ascii')
            apiCodes[i] = str(apiCodes[i]).encode('ascii')
            curveDescs[i] = str(curveDescs[i]).encode('ascii')
        except:
            Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception
        
        # If the first row of the data is an example, first data row the starts in the next row
        if i == dataStartRow and "example" in str(exs[i]):
            dataStartRow = dataStartRow + 1
        
        # Check for missing fields
        mnems[i] = mnems[i].rstrip()
        if mnems[i] == "" and i == dataStartRow:
            Message(wb._path + ": Row " + str(i + 1) + " needs a value. There must have at least one curve on the " + shtName + " sheet.")
            raise Exception
        
        # If the value in the 1st column is not empty and does not contain the words "value" or "units"
        # And if the value in the 5th column does not have the words "example"
        if mnems[i] != "" and "value" not in mnems[i] and "units" not in units[i] and "example" not in str(exs[i]):
            # Output the values
            output += mnems[i]
            output += "  ."
            output += units[i]
            output += "          "
            output += apiCodes[i]  
            output += "          "
            output += ": " + str(rowsOutputed) + " "
            output += curveDescs[i] 
            output += "\n"
            rowsOutputed = rowsOutputed + 1  
    
    return output

# Read the Excel file and write the ~Parameter Information Section
def GetParameterInfo(wb, shtName):

    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # Read the first four columns of the Excel file
    mnems = sht.col_values(0)
    units = sht.col_values(1)
    values = sht.col_values(2)
    descs = sht.col_values(3)
    
    # Determine the row with the headers and the row with the data immediately follows
    try:
        headerRow = mnems.index("LAS Mnemonic")
    except:
        try:
            headerRow = mnems.index("Mnemonic")
        except:
            Message(wb._path + ": Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the in the " + shtName + " sheet. Terminating.")
            raise Exception
    dataStartRow = headerRow + 1
    
    # Output the headers
    output = "~Parameter Information Section"
    output += "\n"
    output += "#MNEM.UNIT                  Value                Description"
    output += "\n"
    output += "#-----------------        ------------    ------------------------------"
    output += "\n"
    rowsOutputed = 1
    
    # For each item in the list representing the first column
    for i in range(dataStartRow, len(mnems)):
        
        # Remove decimal and trailing zeros that were added on Excel import
        try:
            values[i] = float(values[i])
            if values[i] == int(values[i]):
                values[i] = '%d'%values[i]
        except:
            pass
        
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            units[i] = str(units[i]).encode('ascii')
            values[i] = str(values[i]).encode('ascii')
            descs[i] = str(descs[i]).encode('ascii')
        except:
            Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception       
    
        # If the Menm field is not blank output the values
        mnems[i] = mnems[i].rstrip()
        if mnems[i] != "":
            output += mnems[i]
            output += "  ."
            output += units[i]
            output += "          "
            output += values[i]  
            output += "          "
            output += ": "
            output += descs[i] 
            output += "\n"
            rowsOutputed = rowsOutputed + 1
        
    # If no rows were outputed don't output the section header
    if rowsOutputed == 1:
        output = ""

    return output

# Read the Excel file and write the ~Other Information Section
def GetOtherInfo(wb, shtName):
    
    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # Read the first four columns of the Excel file
    otherInfo = sht.col_values(0)
    
    dataStartRow = 2
    
    output = "~Other Information Section"
    output += "\n"
    rowsOutputed = 1
    
    for i in range(dataStartRow, len(otherInfo)):
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            otherInfo[i] = str(otherInfo[i]).encode('ascii')
        except:
            Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception  
        
        # If the other info field is not blank output the values
        if otherInfo[i] != "":   
            output += otherInfo[i]
            output += "\n"
            rowsOutputed = rowsOutputed + 1
    
    # If no rows were outputed don't output the section header
    if rowsOutputed == 1:
        output = ""
        
    return output

# Read the Excel file and write the ~Ascii Log Data Section
def GetAsciiLogData(wb, shtName):

    # Get the sheet from the Excel file
    sht = wb.sheet_by_name(shtName)
    
    # The headers are in the 3rd row
    headersRow = 2
    
    # Figure out the last column that is supposed to have data
    # Look at the first data row and find the last column with data
    row = sht.row_values(headersRow + 1)
    lastCol = len(row) - 1
    for i in range(lastCol, 0, -1):
        if str(row[i]).strip() == "":
            lastCol = lastCol - 1
    del i
    
    # Read the columns of the Excel file which have data
    cols = []
    for i in range(lastCol + 1):
        col = sht.col_values(i)
        cols.append(col)
    del i, col

    rowsOutputed = 0
    
    output = "~A"
    output += "    "

    # For each row
    for i in range(headersRow, len(cols[0])):          
        # For each col
        for col in cols:

            # Remove decimal and trailing zeros that were added on Excel import
            try:
                col[i] = float(col[i])
                if col[i] == int(col[i]):
                    col[i] = '%d'%col[i]
            except:
                pass

            # Make sure all of the characters that will be in the LAS are ascii
            try:
                col[i] = str(col[i]).encode('ascii')
            except:
                Message(wb._path + ": Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
                raise Exception
            
            # Remove trailing whitespace and check for empty string
            col[i] = col[i].rstrip()
            if col[i] == "":
                Message(wb._path + ": Row " + str(i + 1) + " of the " + shtName + " sheet is missing data. Terminating.")
                raise Exception
             
            # Output the values
            output += col[i]
            output += "    "
            
        output += "\n"
        rowsOutputed = rowsOutputed + 1
        
    # If the only output was the section header "~A"
    if rowsOutputed == 1:
        Message(wb._path + ": Enter some data in the " + shtName + " sheet.")
        raise Exception

    return output

# Convert dates (stored in Excel as floats) back to dates
def ConvertToDate(val, wb):
    try:
        if val >= 61:
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(val, wb.datemode)
            date = datetime.datetime(year, month, day, hour, minute, second)
            # Excel treats the first 60 days of 1900 as ambiguous (see Microsoft documentation)
            # Assume the dates are what is indicated in the cell
        else:
            date = datetime.datetime(1900, 1, 1, 0, 0, 0) + datetime.timedelta(days = val - 1)
    except:
        return -1
    
    return date

# Read the LAS file
def ReadLAS(inLAS, output):
    # Open the LAS file for reading
    f = open(inLAS, 'r')
    
    # Read lines into an array
    lines = f.readlines()

    # Mapping for the field names
    # Key (left) is from LAS file
    # Value (right) is from Well Log Content Model
    mapping = {'STRT': 'TopLoggedInterval_ft',
               'STOP': 'BottomLoggedInterval_ft',
               'COMP': '',
               'WELL': 'DisplayName',
               'FLD': 'Field',
               'LOC': 'OtherLocationName',
               'CNTY': 'County',
               'STAT': 'State',
               'CTRY': '',
               'SRVC': '',
               'DATE': 'DateTimeLogRun',
               'UWI': 'WellBoreURI',
               'API': 'APINo',
               'LOGURI': 'LogURI',
               'LOGTYPE': 'LogTypeName',
               'WELLTYPE': 'WellType',
               'LATDEG': 'LatDegree',
               'LONGDEG': 'LongDegree',
               'SRS': 'SRS',
               'LOCUNCERT': 'LocationUncertaintyStatement',
               'TD': 'DrillerTotalDepth_ft',
               'ELGL': 'ElevationGL_ft',
               'LINK': 'RelatedResource',
               'SOURCE': 'Source',
               'NOTE': 'Notes'}

    # Regex Code for the search pattern to find the data values
    #    \s    Match any whitespace characters (space, tab etc.)
    #    {2,}  Match the preceeding character 2 or more times
    #    \S    Match any character NOT whitespace
    #    *     Match the preceding character 0 or more times
    #    ?     Match the preceding character occurs 0 or 1 times
    #    |     Or
    # So match a string bounded on either side by 2 or more sequential whitespaces
    # The string itself can contain up to 1 whitespace character separating words
    # In the example: CNTY.          Rio Arriba          : COUNTY
    # The match is: Rio Arriba
    searchPattern = "\s{2,}([\S*|\s?]*)\s{2,}"

    # Create a variable to hold the corresponding WellLogs CM field name and its value
    data = dict()
    
    # Look at each line and see if the line starts with a field which has a mapping
    for line in lines:
        for k, v in mapping.iteritems():
            if line.split('.')[0].strip() == k:
                units = line.split('.')[1].split(" ")[0]
                d = re.search(searchPattern, line).group(1).strip()
                
                # If the units are in meters convert to feet
                if units == "M" or units == "m":
                    try:
                        d = float(d) * 3.28084
                    except:
                        Message("Warning: It looks like the units of " + d + " for the " + k + " field are in Meters but conversion to feet failed.")
                        Message("This problem may need to be correctly manually.")
                
                # If a data item was found and the field name has a mapping
                if d != "" and mapping[k] != "":
                    data[v] = d

    return data

# Write the data from the LAS file as a row in the WellLogs Content Model
def WriteWellLogsCM(data, fields, row, sht):
    # For each field in the WellLogs CM write the data if the field name matches
    # the field name for which data was found in the LAS file
    for i, field in enumerate(fields):
        for d in data:
            if field == d:
                sht.write(row, i, data[d])

    return

# Write a message in the message box
def Message(message):
    textFrame.insert(END, message + "\n")
    return

if __name__ == "__main__":
    main()