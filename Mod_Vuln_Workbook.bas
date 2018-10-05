Attribute VB_Name = "Mod_Vuln_Workbook"
'This module contains the components that interact with the worksheet itself, not the data contained within it
'
Option Explicit
'
'DECELARE THE GLOBAL VARIABLES TO BE USED
Public mainWorkbook As Workbook  'The main workbook
Public ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet  'Worksheets within the main workbook
Public headerRange As Range, dataRange As Range, totalRange As Range  'Ranges for worksheet ws1
Public rowCount As Integer  'Total number of rown in worksheet ws1, including the header
'
'
'DEFINE ANY VARIABLES THAT WILL BE SHARED BETWEEN DIFFERENT MODULES AND PROCEDURES
Private Sub Define_Global_Variables()
'
   'Set the main workbook and worksheet objects
    Set mainWorkbook = ActiveWorkbook  'Declare mainWorkbook as the active workbook
    Set ws1 = mainWorkbook.ActiveSheet  'Declare ws1 as the active sheet
    
   'Call the procedure that calculates all worksheet sizes that can be used by any other module/procedure
    Call Calculate_Worksheet_Sizes(ws1)
'
End Sub
'
'
Private Function Full_File_Name(fLocation As String, fName As String, fExtension As String) As String
'
    'If the file location does not include a "\" as the last character, add it
    If Not Right(fLocation, 1) = "\" Then
        fLocation = fLocation & "\"
    End If
    '
    'If the file extension does not include a "." as the first character, add it
    If Not Left(fExtension, 1) = "." Then
        fExtension = "." & fExtension
    End If
    '
   'Set the full file path including the file name and extension
    Full_File_Name = fLocation & fName & fExtension
'
End Function
'
'
'CALCULATE THE ACTUAL RANGE OF THE DATA IN AN EXTERNAL WORKSHEET IN A RELIABLE WAY (DO NOT MODIFY)
Private Function External_Data_Range(exWs As Worksheet) As Range
'
    Dim keyColumn As Range, firstColumnCell As Range, keyRow As Range, firstRowCell As Range, rowRange As Range, columnRange As Range, LastCellRange As Range
    Dim LastCol As Integer, LastRow As Integer
    '
   'Set the key column and row for which data must exist in all cells
    Set keyColumn = exWs.Range("B:B")
    Set keyRow = exWs.Range("1:1")
    '
   'Set the first cell in each of the column and row ranges
    Set firstColumnCell = exWs.Range("B1")
    Set firstRowCell = exWs.Range("A1")
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
    Set External_Data_Range = Range("$A$2", LastCellRange.Address)
    '
    'Debug.Print LastCol & ", " & LastRow & ", " & dataRange.Address & vbNewLine  'Test the results against the actual size of the worksheet
'
End Function
'
'
'CALCULATE THE ORIGINAL RANGE OF THE DATA (DO NOT MODIFY)
Private Sub Calculate_Worksheet_Sizes(ws As Worksheet)
'
' Procedure calculates the range values for the header area and the data area
'
   'Range value for the header area
    Set headerRange = ws.Range("$A$1", Range("A1").End(xlToRight))
'
   'Range value for the data area
    Set dataRange = ws.Range("$A$2", Range("A1").SpecialCells(xlLastCell))  'This calculation may yield incorrect ranges if data had been added/deleted or moved.  Only use this if the sheet has not yet been modified or is protected.
'
   'Range value for the total area (header area + data area)
    Set totalRange = ws.Range("$A$1", dataRange.Address)
'
   'Count of rows including header
    rowCount = ws.Cells(Rows.count, "A").End(xlUp).Row
'
End Sub
'
'
'CREATE NEW SHEETS
Private Sub Create_New_Sheets(SheetName1 As String, SheetName2 As String, SheetName3 As String)
'
' Procedure creates three new sheets with defined names as copies of the original sheet
'
   'Set and name the current worksheet
    Set ws1 = ActiveSheet  'Declare the active sheet (Sheet1) as worksheet "ws1"
    ws1.Name = SheetName1  'Name Sheet1 as defined in config
    '
   'Set the zoom level of the worksheet based on the width of the data space
    headerRange.Columns.Select  'Select the width of the data space (only select one row)
    ActiveWindow.Zoom = True
    If ActiveWindow.Zoom > 110 Then 'Set the zoom level on the worksheet to fit the width of the data space, but be no greater than 110%
        ActiveWindow.Zoom = 110
    End If
    Range("A1").Select  'Return the selection to the top left-hand cell
    '
   'Create, name, and resize other worksheets
    ws1.Copy After:=ws1    'Copy Sheet1 to a new Sheet2, after Sheet1, and implicitly make Sheet2 the active sheet
    Set ws2 = ActiveSheet  'Declare the active sheet (Sheet2) as worksheet "ws2"
    ws2.Name = SheetName2  'Name Sheet2 as defined in config
    ws1.Copy After:=ws2    'Copy Sheet1 to a new Sheet3, after Sheet2, and implicitly make Sheet3 the active sheet
    Set ws3 = ActiveSheet  'Declare the active sheet (Sheet3) as worksheet "ws3"
    ws3.Name = SheetName3  'Name Sheet3 as defined in config
'
End Sub
'
'
'SAVE THE FILE
Private Sub Save_Workbook(fileName As String)
'
' Procedure saves the file using the user-defined location and file name
'
   'Commit the save
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
'
End Sub
'
'
'EMAIL THE FILE
Private Sub Mail_Workbook(messageTo As String, messageCC As String, messageBCC As String, messageSubject As String, messageBody As String)
'
'Procedure mails the entire workbook with a defined recipient, subject, and body message.
'  Note: this procedure uses Early Binding to gain efficiency.  When running this procedure on a different PC from which it was originally developed,
'  either enable "Tools \ References...\ Microsoft Outlook x.x Object Library" or switch to a Late Binding method.  For more information on Late Binding
'  methods, please read the last section in "http://www.rondebruin.nl/win/s1/outlook/amail1.htm"
'
    Dim emailApplication As Outlook.Application
    Dim newMessage As Outlook.MailItem
'
   'Set the message objects
    Set emailApplication = CreateObject("Outlook.Application")
    Set newMessage = emailApplication.CreateItem(olMailItem)
'
   'Create the message attributes
    On Error Resume Next
    With newMessage
        .To = messageTo
        .CC = messageCC
        .BCC = messageBCC
        .Subject = messageSubject
        .Body = messageBody
        .BodyFormat = olFormatHTML
        .Attachments.Add ActiveWorkbook.FullName
        .Send  'Use ".Display" for testing and Use ".Send" for production
    End With
    On Error GoTo 0
'
   'Clear the message objects
    Set newMessage = Nothing
    Set emailApplication = Nothing
'
End Sub
'
'
'CLOSE THE FILE
Private Sub Close_Workbook()
'
' Procedure closes the file
'
   'Close the workbook
    mainWorkbook.Close SaveChanges:=False  'File should be saved by the "Save_Workbook_Config" procedure
'
End Sub
'
'
'CLEAR GLOBALLY DEFINED VARIABLES FROM MEMORY
Private Sub Cleanup()
'
' Procedure clears all globally defined variables from memory
'
   'Clear objects (Example: "Set rngMyRange = Nothing")
    'Ranges
    Set headerRange = Nothing
    Set dataRange = Nothing
    Set totalRange = Nothing
    'Worksheets
    Set ws1 = Nothing
    Set ws2 = Nothing
    Set ws3 = Nothing
    'Workbooks
    Set mainWorkbook = Nothing
    '
   'Clear strings (Example: "strMyString = Empty")
    '
   'Clear arrays (Example: "Erase arrMyArray")
    '
   'Zero-out integers and floating point numbers (Example: "myNumber = 0")
    rowCount = 0
    '
'
End Sub
'
'
'ALERT THAT THE MACRO COMPLETED SUCCESSFULLY
Private Sub Completed_Successfully_Alert()
'
' Procedure displays a message box that the macro completed successfully
'
    Dim msgBoxButtons As Integer, messageBox As Integer
    Dim message As String, msgBoxTitle As String
'
   'Define the components of the message box
    msgBoxTitle = "Macro Completed Successfully"  'Title
    message = "The  Vulnerability Formatting macro completed successfully."  'Message
    msgBoxButtons = vbInformation  'Button constant (https://msdn.microsoft.com/en-us/library/aa445082%28v=vs.60%29.aspx)
'
   'Display the message box with the defined components
    messageBox = MsgBox(message, msgBoxButtons, msgBoxTitle)
'
End Sub

