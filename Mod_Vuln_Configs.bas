Attribute VB_Name = "Mod_Vuln_Configs"
'This module provides one place to set all configurations
'
Option Explicit
'
Private Sub Change_Header_Names_Config()
'
    Dim Header_Cell_1 As String, Header_Cell_2 As String, Header_Cell_3 As String, Header_Cell_4 As String, Header_Cell_5 As String, Header_Cell_6 As String, Header_Cell_7 As String
    Dim New_Name_1 As String, New_Name_2 As String, New_Name_3 As String, New_Name_4 As String, New_Name_5 As String, New_Name_6 As String, New_Name_7 As String
'
  'Define the headings to change and what to change them to ("vbNewLine" = New Line)
   'App ID
    Header_Cell_1 = "B1"
    New_Name_1 = "App" & Chr(10) & "ID"
   'Process Owner
    Header_Cell_2 = "C1"
    New_Name_2 = "Process" & Chr(10) & "Owner"
   'WhiteHat ID
    Header_Cell_3 = "J1"
    New_Name_3 = "WhiteHat" & Chr(10) & "ID"
   'Remediation Plan Owner(s)
    Header_Cell_3 = "J1"
    New_Name_3 = "Remediation" & Chr(10) & "Plan Owner(s)"
   'Remediation QA Date
    Header_Cell_4 = "K1"
    New_Name_4 = "Remediation" & Chr(10) & "QA Date"
   'Remediation Prod Date
    Header_Cell_5 = "L1"
    New_Name_5 = "Remediation" & Chr(10) & "Prod Date"
   'Business Risk
    Header_Cell_6 = "M1"
    New_Name_6 = "Business" & Chr(10) & "Risk"
   'Remediation Prod Date
    Header_Cell_7 = "N1"
    New_Name_7 = "Business" & Chr(10) & "Priority"
'
   'Call the procedure that changes the specified header names (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Change_Header_Names", Header_Cell_1, Header_Cell_2, Header_Cell_3, Header_Cell_4, Header_Cell_5, Header_Cell_6, Header_Cell_7, New_Name_1, New_Name_2, New_Name_3, New_Name_4, New_Name_5, New_Name_6, New_Name_7
'
End Sub
'
'
Private Sub Set_Column_Widths_Config()
'
'Procedure to define fixed column widths for specified columns
'
    Dim Column_Width_A As Single, Column_Width_B As Single, Column_Width_C As Single
    Dim Column_Width_D As Single, Column_Width_E As Single, Column_Width_F As Single
    Dim Column_Width_G As Single, Column_Width_H As Single, Column_Width_I As Single
    Dim Column_Width_J As Single, Column_Width_K As Single, Column_Width_L As Single
    Dim Column_Width_M As Single, Column_Width_N As Single
'
  'Define the column widths
    Column_Width_A = 30.57
    Column_Width_B = 4.35  'App ID
    Column_Width_C = 18.57
    Column_Width_D = 6.86
    Column_Width_E = 7.71
    Column_Width_F = 28.43
    Column_Width_G = 11
    Column_Width_H = 10
    Column_Width_I = 12
    Column_Width_J = 20  'Remediation Plan Owner
    Column_Width_K = 11.71
    Column_Width_L = 11.71
    Column_Width_M = 8
    Column_Width_N = 8
'
   'Call the procedure that changes the specified column widths (Do not modify)
    Application.Run "Mod_Vuln_Formatting.Set_Column_Widths", Column_Width_A, Column_Width_B, Column_Width_C, Column_Width_D, Column_Width_E, Column_Width_F, Column_Width_G, Column_Width_H, Column_Width_I, Column_Width_J, Column_Width_K, Column_Width_L, Column_Width_M, Column_Width_N
'
End Sub
'
'
'CONVERT EACH IDENTIFIED COLUMN TO TEXT AND SET THE DEFINED NUMBER FORMAT
Private Sub Convert_Column_Number_Format_Config()
'
   'To add a new column to be formatted, add a new set of variables defining the column and number format as below
    'Column 1
    Dim column1 As String, format1 As String
    column1 = "B"
    format1 = "0"
    'Column 2
    Dim column2 As String, format2 As String
    column2 = "D"
    format2 = "0"
    'Column 3
    Dim column3 As String, format3 As String
    column3 = "G"
    format3 = "m/d/yyyy"
    'Column 4
    Dim column4 As String, format4 As String
    column4 = "H"
    format4 = "0"
    'Column 5
    Dim column5 As String, format5 As String
    column5 = "I"
    format5 = "\RP-0"
    'Column 6
    Dim column6 As String, format6 As String
    column6 = "K"
    format6 = "m/d/yyyy"
    'Column 7
    Dim column7 As String, format7 As String
    column7 = "L"
    format7 = "m/d/yyyy"
'
   'Then add the new column information here
    Dim columnArray() As Variant, formatArray() As Variant
    columnArray = Array(column1, column2, column3, column4, column5, column6, column7)
    formatArray = Array(format1, format2, format3, format4, format5, format6, format7)
'
   'Call the procedure that formats the data as specified (Do not modify)
    Application.Run "Mod_Vuln_Formatting.Convert_Column_Number_Format", columnArray, formatArray
'
End Sub
'
'
Private Sub Format_Data_Config()
'
    Dim columnRange1 As String, columnRange2 As String, columnRange3 As String, columnRange4 As String
'
   'Center data text within specific column ranges
    columnRange1 = "B:B"
    columnRange2 = "D:E"
    columnRange3 = "G:I"
    columnRange4 = "K:N"
    '
   'Apply vertical alignment to the full range of data cells (performed by Format_Data procedure)
   'Auto-fit all rows with data in them (performed by Format_Data procedure)
    
   'Call the procedure that formats the data as specified (Do not modify)
    Application.Run "Mod_Vuln_Formatting.Format_Data", columnRange1, columnRange2, columnRange3, columnRange4
'
End Sub
'
'
'REPLACING DATA
Private Sub Replace_Data_Config()
'
' Procedure finds data and replaces data within the same cell, and optionally applies text formatting
' For example, Search Column "F:F" for the value "cat", and if Found, Replace it with "dog" in a Bold font
'
    Dim rdColumn As String, rdFindValue As String, rdReplaceValue As String, rdFormat As String
'
  'User defined search / replace values and options
   'Modifying the remediatino plan owner field for all remediation plans with no owners provided
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "" 'The value being searched for
    rdReplaceValue = "No RP Owner Provided" 'The value to replace the found value with
    rdFormat = "italic" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
  '
   'Durga Prasad's name change from name in Archer to name in AD/Outlook
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "Raipet, Jaichander" 'The value being searched for
    rdReplaceValue = "Prasad, Durga" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
  '
   'Ravi Venkat's names change from name in Archer to name in AD/Outlook
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "Venkataramani, Ravichandar" 'The value being searched for
    rdReplaceValue = "Venkat, Ravi" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
  '
   'Vijay Vemuri's names change from name in Archer to name in AD/Outlook
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "Vemuri, Naga Venkata Vijay" 'The value being searched for
    rdReplaceValue = "Vemuri, Vijay" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
  '
   'Sai Minnekanti's names change from name in Archer to name in AD/Outlook
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "Minnekanti, Sai Varahari Prasad" 'The value being searched for
    rdReplaceValue = "Minnekanti, Sai" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
  '
   'Sravan Pasumarthi's names change from name in Archer to name in AD/Outlook
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "Pasumarthi, Pola Sravan" 'The value being searched for
    rdReplaceValue = "Pasumarthi, Sravan" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Replace_Data", rdColumn, rdFindValue, rdReplaceValue, rdFormat
'
  'Find text and change it to text using Title Case
   'Rajeev Malik's names change from from all uppercase to title case
    rdColumn = "J:J" 'The column to search in
    rdFindValue = "MALIK, RAJEEV" 'The value being searched for
    rdReplaceValue = "Malik, Rajeev" 'The value to replace the found value with
    rdFormat = "normal" 'Optional text formatting of the replacement value
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Title_Case_Data", rdColumn, rdFindValue, rdFormat
  '
End Sub
'
'
'ADDING DATA
Private Sub Add_Data_Config()
'
' Procedure finds data and replaces data in a diffrent column of the same row
' For example, Find the value "cat" in column "F:F" and replace the current value in column "R:R" of that same row with "dog"
'
'FIND DATA AND REPLACE DATA IN A DIFFERENT COLUMN OF THE SAME ROW
  'Definitions (Do not modify)
    Dim rdColumn As String 'rdColumn is the columns to search in (only specify a single column is rdOffset is not 0)
    Dim rdFindValue As String 'rdFindValue is the data to search for
    Dim rdReplaceValue As String 'rdReplaceValue is the data that will be added
    Dim rdOffset As Integer 'rdOffset is the number of columns away from the column that the data was found in to add the replacement data
'
  'User defined data
   'CVM 449 - Risk Transferred
    rdColumn = "D:D"
    rdFindValue = "449"
    rdReplaceValue = "Risk Transferred"
    rdOffset = 8
    'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_Data", rdColumn, rdFindValue, rdReplaceValue, rdOffset
'
   'CVM 1580 - Risk Accepted
    rdColumn = "D:D"
    rdFindValue = "1580"
    rdReplaceValue = "Risk Accepted"
    rdOffset = 8
    'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_Data", rdColumn, rdFindValue, rdReplaceValue, rdOffset
'
   'CVM 1581 - Risk Accepted
    rdColumn = "D:D"
    rdFindValue = "1581"
    rdReplaceValue = "Risk Accepted"
    rdOffset = 8
    'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_Data", rdColumn, rdFindValue, rdReplaceValue, rdOffset
'
   'CVM 1582 - Risk Accepted
    rdColumn = "D:D"
    rdFindValue = "1582"
    rdReplaceValue = "Risk Accepted"
    rdOffset = 8
    'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_Data", rdColumn, rdFindValue, rdReplaceValue, rdOffset
'
   'CVM 1583 - Risk Accepted and Client to Assume Risk
    rdColumn = "D:D"
    rdFindValue = "1583"
    rdReplaceValue = "Risk Accepted and Client to Assume Risk"
    rdOffset = 8
    'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_Data", rdColumn, rdFindValue, rdReplaceValue, rdOffset
'
End Sub
'
'
'ADDING DATA FROM A DIFFERENT WORKSHEET
Private Sub Add_External_Data_Config()
'
' Procedure finds data in the active worobook that is specified in a different workbook, then replaces data in the active
' workbok with data from the different workbook in a diffrent column (as specified in the different workbook) of the same row
' of the active workbook
'
'   Example: The saved workbook named Animal_Characteristics has the value "Cat" in Column A, "Pointy Ears" in Column B, the number 2
'   in Column C, and "F:F" in Column D. This procedure searches for the value "Cat" in the active workbook's Column "F:F", and if
'   found, replaces the current value in column "H:H" (2 columns to the right) of that same row with the value "Pointy Ears".
'
    Dim fLocation As String, fName As String, fExtension As String, wsName As String
'
   'Define file path
    fLocation = "C:\Users\benvenua\Documents\Moody's\Projects\sSDLC\Open Vulnerabilities" 'Full folder location, (Example: "C:\Users\doej\Documents\VulnMngmnt\Spreadsheets\")
    fName = "Vulnerability Report Data Modification" 'Full file name without the extension (Example: "MySpreadsheet1")
    fExtension = ".xlsx" 'Full file extension including initial period (Default is ".xlsx")
    '
   'Define the name of the worksheet (the Tab) within the workbook to look for the data within
    wsName = "Add Data" 'The full name of the worksheet

   'Call the procedure that loads the external worksheet and commits each change (Do not modify)
    Application.Run "Mod_Vuln_Data_Changes.Add_External_Data", fLocation, fName, fExtension, wsName
'
End Sub
'
'
'HIGHLIGHTING SPECIFIED RANGES OF CELLS
Private Sub Highlight_Cells_Config()
'
'Procedure to highlight cells within a specified range
'   Case 1 is to highlight the QA and Prod Remediation cells if either date is before today
'   Case 2 is to highlight the Remediation Plan cells if any are blank
'
    Dim searchRange1 As String, highlightRange1 As String, searchRange2 As String, highlightRange2 As String
    Dim color1 As Integer
    Dim tintAndShade2 As Single
    '
   'Case 1: Highlight the QA and Prod Remediation cells if either date is before today
    searchRange1 = "L:L"  'The dates to compare to today (not necessarily what will be highlighted
    highlightRange1 = "K:L"
    color1 = 6  'Bright yellow highlighting using the standard Excel color
    '
   'Case 2: Highlight the Remediation Plan cells if any are blank
    searchRange2 = "I:L"
    highlightRange2 = "I:L"
    'themeColor2 = "xlThemeColorAccent2"  'Red  'Setting a theme must be done within the "Mod_Vuln_Formatting.Hightlight_Cells" procedure, as it does not accept a variable
    tintAndShade2 = 0.4  'Lighter by 60% of the specified built-in Excel color theme
    '
   'Call the procedure that highlights the ranges of cells as specified (Only modify to add/remove colors)
    Application.Run "Mod_Vuln_Formatting.Highlight_Cells", searchRange1, highlightRange1, color1, searchRange2, highlightRange2, tintAndShade2
'
End Sub
'
'
'CREATING NEW SHEETS
Private Sub Create_New_Sheets_Config()
'
' Procedure creates new sheets within the same workbook, and names them
'
    Dim SheetName1 As String, SheetName2 As String, SheetName3 As String
'
   'Sheet names to be used
    SheetName1 = "By Application"
    SheetName2 = "By Process Owner"
    SheetName3 = "By Remediation Owner"
'
   'Call the procedure that commits each change (Do not modify)
    Application.Run "Mod_Vuln_Workbook.Create_New_Sheets", SheetName1, SheetName2, SheetName3
'
End Sub
'
'
'ENBOLDEN COLUMNS
'
' Procedure applies the Bold text format to each column specified per worksheet
'
Private Sub Bold_Columns_Config(wrkSheet As String)
'
    Dim ws1BoldColumns As String, ws2BoldColumns As String, ws3BoldColumns As String
'
   'User defined colums to bold per specified worksheet
    ws1BoldColumns = "A:A"  'Worksheet 1
    ws2BoldColumns = "A:B"  'Worksheet 2
    ws3BoldColumns = "A:B"  'Worksheet 3
'
   'Call the procedure that bolds the specified column(s) (Do not modify)
    Application.Run "Mod_Vuln_Formatting.Bold_Columns", wrkSheet, ws1BoldColumns, ws2BoldColumns, ws3BoldColumns
'
End Sub
'
'
'SORTING DATA
Private Sub Sort_Data_Config(wrkSheet As String)
'
' Procedure sorts all data (not the header) based on a hierarchical order of columns
'   - This procedure currently must include five columns to sort by.  If less columns are
'     required, set the remaining columns to a column that is already listed.
'   - This procedure currently sorts all columns in ascending order
'
    Dim sortRange1 As String, sortRange2 As String, sortRange3 As String, sortRange4 As String, sortRange5 As String
    '
    Select Case wrkSheet
        '
       'Sort order for Worksheet 1
        Case "wrkSheet1" 'Ranges must be listed in sort hierarchy
            sortRange1 = "A:A"
            sortRange2 = "E:E"
            sortRange3 = "G:G"
            sortRange4 = "L:L"
            sortRange5 = "A:A"
            'Call the procedure that commits each change (Do not modify)
            Application.Run "Mod_Vuln_Data_Changes.Sort_Data", wrkSheet, sortRange1, sortRange2, sortRange3, sortRange4, sortRange5
            '
       'Sort order for Worksheet 2
        Case "wrkSheet2" 'Ranges must be listed in sort hierarchy
            sortRange1 = "A:A"
            sortRange2 = "B:B"
            sortRange3 = "E:E"
            sortRange4 = "G:G"
            sortRange5 = "L:L"
            'Call the procedure that commits each change (Do not modify)
            Application.Run "Mod_Vuln_Data_Changes.Sort_Data", wrkSheet, sortRange1, sortRange2, sortRange3, sortRange4, sortRange5
            '
       'Sort order for Worksheet 3
        Case "wrkSheet3" 'Ranges must be listed in sort hierarchy
            sortRange1 = "A:A"
            sortRange2 = "B:B"
            sortRange3 = "F:F"
            sortRange4 = "H:H"
            sortRange5 = "L:L"
            'Call the procedure that commits each change (Do not modify)
            Application.Run "Mod_Vuln_Data_Changes.Sort_Data", wrkSheet, sortRange1, sortRange2, sortRange3, sortRange4, sortRange5
            '
        Case Else
           'Message box to be displayed if there is an error sorting the data (Do not modify)
            MsgBox "Undefined worksheet in Mod_Vuln_Configs.Sort_Data_Config.", Buttons:=vbCritical, Title:="Error Sorting Data"
        '
    End Select
'
End Sub
'
'
'SAVE THE FILE
Private Sub Save_Workbook_Config()
'
    Dim currentDate As String, Path As String, Title As String, Extension As String, fileName As String
'
   'Define the elements of the file name
    Path = "C:\Users\benvenua\Documents\Moody's\Projects\sSDLC\Open Vulnerabilities\"  'Include backslash ("\")at the end of the path
    Title = "Open Application Vulnerabilities "
    currentDate = Format(Date, "yyyy-mm-dd")  'Define the date format
    Extension = ".xlsx"
    '
   'Define the concatonated file name
    fileName = Path & Title & currentDate & Extension
'
   'Call the procedure that saves and closes the workbook (Do not modify)
    Application.Run "Mod_Vuln_Workbook.Save_Workbook", fileName
'
End Sub
'
'
'E-MAIL THE FILE
Private Sub Mail_Workbook_Config()
'
    Dim currentDateTime As String
    Dim messageTo As String, messageCC As String, messageBCC As String, messageSubject As String, messageBody As String
'
   'Define the format for the date to be displayed in the message
    currentDateTime = Format(Now(), "m/d/yyyy h:nn AM/PM")
'
   'Define the message attributes
    messageTo = "Benvenuti, Adrian;"  'Use the full "user@company.com" email address if outside of corporate network
    messageCC = ""
    messageBCC = ""
    messageSubject = "Weekly Open Application Vulnerabilities"
    messageBody = "This week's list of open application vulnerabilities is attached." & vbNewLine & vbNewLine & _
                  "These are current as of " & currentDateTime & "." & vbNewLine & vbNewLine & vbNewLine & _
                  "Adrian"
'
   'Call the procedure that creates and sends the e-mail message (Do not modify)
    Application.Run "Mod_Vuln_Workbook.Mail_Workbook", messageTo, messageCC, messageBCC, messageSubject, messageBody
'
End Sub
