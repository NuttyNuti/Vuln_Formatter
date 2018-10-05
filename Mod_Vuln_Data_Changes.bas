Attribute VB_Name = "Mod_Vuln_Data_Changes"
'This module contains the components to make changes to the data
'
Option Explicit
'
Private Sub Change_Header_Names(Header_Cell_1 As String, Header_Cell_2 As String, Header_Cell_3 As String, Header_Cell_4 As String, Header_Cell_5 As String, Header_Cell_6 As String, Header_Cell_7 As String, New_Name_1 As String, New_Name_2 As String, New_Name_3 As String, New_Name_4 As String, New_Name_5 As String, New_Name_6 As String, New_Name_7 As String)
'
' Procedure changes user defined column header values with new user defined values
'
   'Change header names
    Range(Header_Cell_1).Value = New_Name_1
    Range(Header_Cell_2).Value = New_Name_2
    Range(Header_Cell_3).Value = New_Name_3
    Range(Header_Cell_4).Value = New_Name_4
    Range(Header_Cell_5).Value = New_Name_5
    Range(Header_Cell_6).Value = New_Name_6
    Range(Header_Cell_7).Value = New_Name_7
'
End Sub
'
'
Private Sub Replace_Data(rdColumn As String, rdFindValue As String, rdReplaceValue As String, rdFormat As String)
'
' Procedure finds and replaces data with specified formating changes
' For example, Find the value "cat" in column "F:F" and replace it with "dog" italicised
'   Note: This is a direct Find/Replace procedure, and does not provide for an offset data replacement location
'
    Dim rdRange As Range
'
    Set rdRange = Range(rdColumn)
'
    Application.ReplaceFormat.Clear
    With Application.ReplaceFormat.Font
        .Name = "Calibri"
        .FontStyle = rdFormat
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    '
    rdRange.Replace What:=rdFindValue, Replacement:=rdReplaceValue, LookAt:=xlPart _
        , SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=True
    '
    Application.ReplaceFormat.Clear
'
End Sub
'
'
Private Sub Title_Case_Data(tcdColumn As String, tcdFindValue As String, tcdFormat As String)
'
' Procedure finds data and rewrites it using (proper) title case
'
'   Example: Find the value "apple, ORANGE, BaNaNnA" in column "F:F" and replace it with "Apple, Orange, Bananna"
'
    Dim tcdRange As Range, tcdFoundValueLocations As Range
    Dim firstAddress As String, errorMessage As String
'
    tcdColumn = Column_Name(tcdColumn)  'Correct an improperly formatted single column name
    Set tcdRange = ws1.Range(tcdColumn)
'
    With tcdRange
        Set tcdFoundValueLocations = .Find( _
            What:=tcdFindValue, _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
        '
        If Not tcdFoundValueLocations Is Nothing Then 'If the value being searched for is found at least once within the specified range
            firstAddress = tcdFoundValueLocations.Address
            Do
                tcdFoundValueLocations.Value = Application.Proper(tcdFoundValueLocations.Value)
                Set tcdFoundValueLocations = .FindNext(tcdFoundValueLocations)
            Loop While tcdFoundValueLocations.Address <> firstAddress
        Else 'If the value being searched is not found within the specified range
            errorMessage = "The search value " & tcdFindValue & " was not found in " & tcdRange.Address & "." & vbNewLine & vbNewLine & "Please edit the macro's Replace Data configuration located within" & vbNewLine & vbTab & "Macro Module: Mod_Vuln_Data_Changes" & vbNewLine & "to remove or correct this entry."
            MsgBox errorMessage, Buttons:=vbExclamation, Title:="Replace Data Error"
        End If
    End With
'
End Sub
'
'
Private Function Column_Name(columnName As String) As String
'
' Procedure corrects an improperly formatted single column name
' For example, it changes column name "D" to "D:D"
'
    Dim charPosition As Integer
    '
    charPosition = InStr(2, columnName, ":")
    '
   'If the file location does not include a "\" as the last character, add it
    If charPosition = 0 Then
        columnName = columnName & ":" & columnName
    End If
    '
   'Set the full file path including the file name and extension
    Column_Name = columnName
'
End Function
'
'
Private Sub Add_Data(adColumn As String, adFindValue As String, adReplaceValue As String, adOffset As Integer, exFilePath As String)
'
' Procedure finds data and replaces data in a diffrent column of the same row
'
'   Example: Find the value "cat" in column "F:F" and replace the current value in column "R:R" of that same row with "dog"
'   Note: In this procedure, no formatting changes are applied to the data
'
    Dim adRange As Range, adFoundValueLocations As Range
    Dim firstAddress As String, errorMessage As String
'
    adColumn = Column_Name(adColumn)  'Correct an improperly formatted single column name
    Set adRange = ws1.Range(adColumn)
'
    With adRange
        Set adFoundValueLocations = .Find( _
            What:=adFindValue, _
            LookIn:=xlFormulas, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
        '
        If Not adFoundValueLocations Is Nothing Then 'If the value being searched for is found at least once within the specified range
            firstAddress = adFoundValueLocations.Address
            Do
                adFoundValueLocations.Offset(, adOffset).Value = adReplaceValue
                Set adFoundValueLocations = .FindNext(adFoundValueLocations)
            Loop While adFoundValueLocations.Address <> firstAddress
        Else 'If the value being searched is not found within the specified range
            errorMessage = "CVM# " & adFindValue & " was not found in column " & adRange.Address & ".  Maybe it was remediated?" & vbNewLine & vbNewLine & _
                "Please edit the macro's replacement data located within:" & vbNewLine & vbNewLine & _
                exFilePath & vbNewLine & vbNewLine & _
                "to remove or correct this entry."
            MsgBox errorMessage, Buttons:=vbExclamation, Title:="Add Data Error"
            'Use the following error message if searching against data within a config file and not an external file
            '"CVM# " & adFindValue & " was not found in " & adRange.Address & "." & vbNewLine & vbNewLine & "Macro Module: Mod_Vuln_Configs" & vbNewLine & "to remove or correct this entry."
        End If
    End With
'
End Sub
'
'
Private Function Full_File_Name(fName As String, fExtension As String) As String
'
   'If the file extension does not include a "." as the first character, add it
    If Not Left(fExtension, 1) = "." Then
        fExtension = "." & fExtension
    End If
    '
   'Set the full file path including the file name and extension
    Full_File_Name = fName & fExtension
'
End Function
'
'
Private Function Full_File_Path_and_Name(fLocation As String, fName As String, fExtension As String) As String
'
    Dim fullFileName As String
'
    'If the file location does not include a "\" as the last character, add it
    If Not Right(fLocation, 1) = "\" Then
        fLocation = fLocation & "\"
    End If
    '
    'If the file extension does not include a "." as the first character, add it
    fullFileName = Full_File_Name(fName, fExtension)
    '
   'Set the full file path including the file name and extension
    Full_File_Path_and_Name = fLocation & fullFileName
'
End Function
'
'
'CALCULATE THE ACTUAL RANGE OF THE DATA IN AN EXTERNAL WORKSHEET IN A RELIABLE WAY (DO NOT MODIFY)
Private Function External_Data_Range(ws As Worksheet) As String
'    Dim wb As Workbook
'    Dim ws As Worksheet
    Dim keyColumn As Range, firstColumnCell As Range, keyRow As Range, firstRowCell As Range, rowRange As Range, columnRange As Range, LastCellRange As Range
    Dim LastCol As Integer, LastRow As Integer
    '
   'Set the named workbook and worksheet  (For debugging purposes)
'    Set wb = Workbooks.Open(fullFileName)
'    Set ws = wb.Worksheets(wsName)
    '
   'Set the key column and row for which data must exist in all cells
    Set keyColumn = ws.Range("B:B")
    Set keyRow = ws.Range("1:1")
    '
   'Set the first cell in each of the column and row ranges
    Set firstColumnCell = ws.Range("B1")
    Set firstRowCell = ws.Range("A1")
    '
   'Find the bottom-most cell in the key column range
    Set rowRange = keyColumn.Find(What:="*", _
                After:=firstColumnCell, _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False)
    '
   'Find the right-most cell in the key row range
    Set columnRange = keyRow.Find(What:="*", _
                After:=firstRowCell, _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False)
    '
   'Set the data range based on the bottom-most and right-most cells on the sheet
    If rowRange Is Nothing Then
        LastCol = columnRange.Column
        LastRow = 1
    Else
        LastCol = columnRange.Column
        LastRow = rowRange.Row
    End If
    Set LastCellRange = Cells(LastRow, LastCol)
'
   'Set the function return value
    External_Data_Range = "$A$2" & ":" & LastCellRange.Address
    '
    'Debug.Print LastCol & ", " & LastRow & ", " & dataRange.Address & vbNewLine  'Test the results against the actual size of the worksheet
'
End Function
'
'
Private Sub Add_External_Data(fLocation As String, fName As String, fExtension As String, wsName As String)
'
' Procedure finds data in the active worobook that is specified in a different workbook, then replaces data in the active
' workbok with data from the different workbook in a diffrent column (as specified in the different workbook) of the same row
' of the active workbook
'
'   Example: The saved workbook named Animal_Characteristics has the value "Cat" in Column A, "Pointy Ears" in Column B, the number 2
'   in Column C, and "F:F" in Column D. This procedure searches for the value "Cat" in the active workbook's Column "F:F", and if
'   found, replaces the current value in column "H:H" (2 columns to the right) of that same row with the value "Pointy Ears".
'
    Dim exFilePath As String, exFileName As String, exDataRangeString As String
    Dim exWb As Workbook
    Dim exWs As Worksheet
    Dim exDataRange As Range
    Dim adArray() As Variant
    Dim exRow As Integer, exColumn As Integer, i As Integer
    Dim adColumn As String, adFindValue As String, adReplaceValue As String
    Dim adOffset As Integer
'
   'Set the full file path including the file name and extension
    exFilePath = Full_File_Path_and_Name(fLocation, fName, fExtension)
'    exFileName = Full_File_Name(fName, fExtension)
    '
   'Set the workbook and worksheet objects as specified by the file path and worksheet name
    Set exWb = Workbooks.Open(exFilePath)  'Needs error handling to be added
    Set exWs = exWb.Sheets(wsName)  'Needs error handling to be added
    '
   'Calculate the data range in the external worksheet
    Set exDataRange = Range(External_Data_Range(exWs))
'    Set exDataRange = exWs.Range("$A$2", Range("A1").SpecialCells(xlLastCell))  'This method does not work if the worksheet has been modified
    '
   'Count rows and columns
    exRow = exDataRange.Rows.count
    exColumn = exDataRange.Columns.count
    '
   'Redimension the array with the actual calculated size
    ReDim adArray(exRow, exColumn) As Variant
    '
   'Load the values from all cells within the range into the array
    adArray = exDataRange.Value
    '
   'Close the external workbook and explicitly activate the main worksheet
   'A possible alternative would be to move this section until after the FOR loop, throw a flag if there is an error locating one of the replacement values, and exit the SUB without closing this file
   'A further possibility would be to highlight the external cells with errors
    exWb.Close SaveChanges:=False
    ws1.Activate
    '
   'Catch any errors around loading the data
    On Error GoTo Errhandler
    '
   'Add data as specified in the external worksheeet
    For i = 1 To UBound(adArray, 1)
    '
       'Load the data from the array into the variables to be sent to the Add_Data procedure
        adColumn = adArray(i, 1)
        adFindValue = adArray(i, 2)
        adOffset = adArray(i, 3)
        adReplaceValue = adArray(i, 4)
        '
       'Correct an improperly formatted single column name
        adColumn = Column_Name(adColumn)
        '
       'Perform the procedure to add the data as specified in each row of the external worksheet
        Call Add_Data(adColumn, adFindValue, adReplaceValue, adOffset, exFilePath)
    '
    Next i
    '
   'Reset error handling
    On Error GoTo 0
'
   'Release the worksheet and workbook objects
    Set exWs = Nothing
    Set exWb = Nothing
'
'End the procedure, so that the error handler is not executed when there are no errors
Exit Sub
'
Errhandler:
    '
    Dim message As String
    '
    Select Case Err
    '
        Case 9:  ' Error 9: "ArgumentNullException": Array is Nothing, OR "RankException": Rank is less than 1 or Rank is greater than the rank of Array.
            message = MsgBox("Either the workbook or worksheet containing the data could not be accessed, or " & _
                    "there was less then the expected amount of data found in the spreadsheet." & vbCr & vbCr & _
                    "Error in Mod_Vuln_Data_Changes.Add_External_Data", Buttons:=vbExclamation, Title:="Error Loading External Data")
        '
        Case Else:  ' An error other than 9 has occurred.
            ' Display the error number and the error text.
            message = MsgBox("Error # " & Err & " : " & Error(Err), Buttons:=vbExclamation, Title:="Error Loading External Data")
    '
    End Select
    '
   'Release the worksheet and workbook objects
    Set exWs = Nothing
    Set exWb = Nothing
'
End Sub
'
'
Private Sub Sort_Data(wrkSheet As String, sortRange1 As String, sortRange2 As String, sortRange3 As String, sortRange4 As String, sortRange5 As String)
'
' Procedure to sort the data (not the header) by the columns specified by the user in a hierarchical order
'   - This procedure currently sorts by up to five columns
'   - This procedure currently sorts all columns in ascending order
'
    Dim ws As Worksheet
    Dim range1 As Range, range2 As Range, range3 As Range, range4 As Range, range5 As Range, currentDataRange As Range
'
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
        Case Else
           'Message box to be displayed if the Sort_Data procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Sort_Data", Buttons:=vbCritical, Title:="Error Sorting Data"
    End Select
'
    Set range1 = Range(sortRange1)
    Set range2 = Range(sortRange2)
    Set range3 = Range(sortRange3)
    Set range4 = Range(sortRange4)
    Set range5 = Range(sortRange5)
'
   'SORT DATA
    Set currentDataRange = ws.Range(dataRange.Address)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=range1, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=range2, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=range3, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=range4, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=range5, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange currentDataRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'
   'Clear the sort fields
    ws.Sort.SortFields.Clear
'
End Sub
'
'
Private Sub Move_Column(wrkSheet As String)
'
    Application.ScreenUpdating = False
    '
    Select Case wrkSheet
        Case "wrkSheet2"  'Worksheet 2
            ws2.Columns("C:C").Cut
            ws2.Columns("A:A").Insert Shift:=xlToRight
        Case "wrkSheet3"  'Worksheet 3
            ws3.Columns("J:J").Cut
            ws3.Columns("A:A").Insert Shift:=xlToRight
        Case Else
           'Message box to be displayed if the Move_Column procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Move_Column", Buttons:=vbCritical, Title:="Error Moving Columns"
    End Select
    '
    Application.ScreenUpdating = True
'
End Sub
