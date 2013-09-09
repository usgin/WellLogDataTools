'''
Created on Jun 12, 2013

@author: Jessica Good Alisdairi, AZGS jalisdairi@azgs.az.gov

This script takes data inputed into the BoreholeLASLogData Content Model
and creates a LAS 2.0 file.

The following assumptions are made about the Excel workbook being read:

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
    "API NUMBER" "LOG URI", "LOG TYPE", "WELL TYPE", "LAT DEGREE", "LONG DEGREE", "SRS", "LOCATION UNCERTAINTY", "TOTAL DEPTH", 
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
import os, glob
try:
    import xlrd
except:
    print "Import of XLRD module failed.\nThe XLRD module can be downloaded from: http://pypi.python.org/pypi/xlrd"
import Tkinter, tkFileDialog
from Tkinter import *


# Main function for the Excel to NGDS Feature ArcGIS Tool
def main(argv=None):
    root = Tkinter.Tk()
    root.title("Excel to LAS Converter")
    root.minsize(0, 0)
    
    # Create the frame for the conversion option buttons
    labelframe1 = LabelFrame(root, text = "Conversion Options")
    labelframe1.pack(fill = "both", expand = "yes")
    # Create a frame for both buttons
    bframe = Frame(labelframe1)
    bframe.pack()
    # Create a frame for the 1st button
    b1frame = Frame(bframe, bd = 5)
    b1frame.pack(side = LEFT)
    # Create the 1st button
    B = Tkinter.Button(b1frame, text = "Convert Single Excel File", command = ConvertSingleFile)
    B.pack()
    # Create a frame for the 2nd button
    b2frame = Frame(bframe, bd = 5)
    b2frame.pack(side = LEFT)
    # Create the 2nd button
    B2 = Tkinter.Button(b2frame, text = "Convert All Excel Files in a Folder", command = ConvertFolder)
    B2.pack()
    
    # Create the frame for the messages list
    global textFrame
    msgFrame = LabelFrame(root, text="Messages")
    # Create the frame for the text
    textFrame = Tkinter.Text(msgFrame, height = 30)
    # Create the scrollbar
    yscrollbar = Scrollbar(msgFrame)
    yscrollbar.pack(side = RIGHT, fill = Y)
    textFrame.pack(side = LEFT, fill = Y)
    yscrollbar.config(command = textFrame.yview)
    textFrame.config(yscrollcommand = yscrollbar.set)
    msgFrame.pack(fill = X, expand = "yes")

    # Create the frame for the Exit button
    b3frame = Frame(root, bd = 5)
    b3frame.pack()
    # Create the Exit button
    B3 = Tkinter.Button(b3frame, text = "Exit", command = ExitConvert)
    B3.pack()
    
    root.mainloop()
    return
   
# Prompt for file to convert
def ConvertSingleFile():
    # Clear the text frame
    textFrame.delete("1.0", END)
    
    xlFile = tkFileDialog.askopenfilename(filetypes=[("All Files","*"), ("Excel Files","*.xlsx;*.xls")])
    # If cancel was pressed
    if xlFile == "":
        return
  
    xlFiles = []
    xlFiles.append(xlFile)
    Convert(xlFiles)
    return

# Prompt for folder which holds all the Excel files to convert 
def ConvertFolder():
    # Clear the text frame
    textFrame.delete("1.0", END)
    
    global xlBaseFile
    path = tkFileDialog.askdirectory()
    # If cancel was pressed
    if path == "":
        return

    # Find all Excel files in the folder
    xls = glob.glob(path + "/*.xls")
    xlsx = glob.glob(path + "/*.xlsx")
    xlFiles = xls + xlsx

    # Make a list of the file names
    xlFileNames = []
    for xlFile in xlFiles:
        xlTempBaseFile = os.path.basename(xlFile)
        xlTempFileName = os.path.splitext(xlTempBaseFile)[0]
        xlFileNames.append(xlTempFileName)
    
    # Find any duplicate file names
    if len(xlFileNames) != len(set(xlFileNames)):
        dupes = []
        for i in xlFileNames:
            if xlFileNames.count(i) > 1:
                if i not in dupes:
                    dupes.append(i)
        xlBaseFile = ', '.join(dupes)
        if len(dupes) == 1:
            Message("is used more than once as a file name. Make sure all file names are unique.")
        else:
            Message("are used more than once as file names. Make sure all file names are unique.")
    Convert(xlFiles)
    return

# Exit the program
def ExitConvert():
    exit()
    return

# Start the conversion
def Convert(xlFiles):
    global xlBaseFile
    for xlFile in xlFiles:
        xlBaseFile = os.path.basename(xlFile)
        xlFileName = os.path.splitext(xlBaseFile)[0]
        path = os.path.splitext(os.path.dirname(xlFile))[0]
        lasFile = path + '\\' + xlFileName + '.las'      
        fileOut = open(lasFile, 'w')                     # Create output file
        output = ""
        fileOut.write(output)
        output = ReadExcel(xlFile, output)
        fileOut.write(output)
        fileOut.close()
        if output == "":
            os.remove(lasFile)    
    return

# Message Box
def Message(message):
    textFrame.insert(END, xlBaseFile + ": "+ message + "\n")
    return

# Read the Excel file
def ReadExcel(inExcel, output):
    try:
        try:
            wb = xlrd.open_workbook(inExcel)
        except:
            Message("Unable to open workbook. Terminating.")
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
                    Message("Can't find the sheet starting with " + key + ". Terminating.")
                    raise Exception
        
        output += GetVersionInfo(wb, shts["~V"])
        output += GetWellInfo(wb, shts["~W"])
        output += GetCurveInfo(wb, shts["~C"])
        if shts["~P"] != None:
            output += GetParameterInfo(wb, shts["~P"])
        if shts["~O"] != None:
            output += GetOtherInfo(wb, shts["~O"])
        output += GetAsciiLogData(wb, shts["~A"])
    
        Message("Conversion Successful")
        
    except Exception:
        Message("Conversion Failed")
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
            Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception  
        
        # Remove trailing whitespace and check for empty string
        versInfo[i] = versInfo[i].rstrip()
        if versInfo[i] == "":
            Message("Row " + str(i + 1) + " in the " + shtName + " sheet is missing data. Terminating.")
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
            Message("Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the " + shtName + " sheet. Terminating.")
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
  
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            values[i] = str(values[i]).encode('ascii')
        except:
            Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception
        
        # Check for missing fields and values
        mnems[i] = mnems[i].rstrip()
        if mnems[i] == "":
            Message("Row " + str(i + 1) + " must have a field name (Mnem) in the " + shtName + " sheet.")
            raise Exception
        values[i] = values[i].rstrip()
        if values[i] == "" and i < 15 + dataStartRow:
            Message("Row " + str(i + 1) + " must have a value in the " + shtName + " sheet.")
            raise Exception
        
        # Output the fields
        output += mnems[i]
        
        # Each Mnem value must be followed by a period so insert one if not already there
        if "." not in mnems[i]:
            output += "."
        
        # Output the values
        output += "          "
        output += values[i]
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
            Message("Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the " + shtName + " sheet. Terminating.")
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

        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            units[i] = str(units[i]).encode('ascii')
            apiCodes[i] = str(apiCodes[i]).encode('ascii')
            curveDescs[i] = str(curveDescs[i]).encode('ascii')
        except:
            Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
            raise Exception
        
        # If the first row of the data is an example, first data row the starts in the next row
        if i == dataStartRow and "example" in str(exs[i]):
            dataStartRow = dataStartRow + 1
        
        # Check for missing fields
        mnems[i] = mnems[i].rstrip()
        if mnems[i] == "" and i == dataStartRow:
            Message("Row " + str(i + 1) + " needs a value. There must have at least one curve on the " + shtName + " sheet.")
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
            Message("Unable to find the row labeled \"LAS Mnemonic\" or \"Mnemonic\" in the first column of the in the " + shtName + " sheet. Terminating.")
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
        
        # Make sure all of the characters that will be in the LAS are ascii
        try:
            mnems[i] = str(mnems[i]).encode('ascii')
            units[i] = str(units[i]).encode('ascii')
            values[i] = str(values[i]).encode('ascii')
            descs[i] = str(descs[i]).encode('ascii')
        except:
            Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
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
            Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
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
    row = sht.row_values(2)
    lastCol = len(row) - 1
    for i in range(lastCol, 0, -1):
        if "value" in str(row[i]) and "units" in str(row[i]):
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

            # Make sure all of the characters that will be in the LAS are ascii
            try:
                col[i] = str(col[i]).encode('ascii')
            except:
                Message("Non-ascii character in row " + str(i + 1) + " of the " + shtName + " sheet. Terminating.")
                raise Exception
            
            # Remove trailing whitespace and check for empty string
            col[i] = col[i].rstrip()
            if col[i] == "":
                Message("Row " + str(i + 1) + " of the " + shtName + " sheet is missing data. Terminating.")
                raise Exception
             
            # Output the values
            output += col[i]
            output += "    "
            
        output += "\n"
        rowsOutputed = rowsOutputed + 1
        
    # If the only output was the section header "~A"
    if rowsOutputed == 1:
        Message("Enter some data in the " + shtName + " sheet.")
        raise Exception

    return output

if __name__ == "__main__":
    main()