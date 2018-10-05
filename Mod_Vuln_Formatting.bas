Attribute VB_Name = "Mod_Vuln_Formatting"
'This module contains the components to format worksheets, data and headers
'
Option Explicit
'
'
'EXPLICITLY SET THE WIDTH OF ALL COLUMNS
Private Sub Set_Column_Widths(Column_Width_A As Single, Column_Width_B As Single, Column_Width_C As Single, Column_Width_D As Single, Column_Width_E As Single, Column_Width_F As Single, Column_Width_G As Single, Column_Width_H As Single, Column_Width_I As Single, Column_Width_J As Single, Column_Width_K As Single, Column_Width_L As Single, Column_Width_M As Single, Column_Width_N As Single)
'
   'Set defined column widths
    Columns("A:A").ColumnWidth = Column_Width_A
    Columns("B:B").ColumnWidth = Column_Width_B
    Columns("C:C").ColumnWidth = Column_Width_C
    Columns("D:D").ColumnWidth = Column_Width_D
    Columns("E:E").ColumnWidth = Column_Width_E
    Columns("F:F").ColumnWidth = Column_Width_F
    Columns("G:G").ColumnWidth = Column_Width_G
    Columns("H:H").ColumnWidth = Column_Width_H
    Columns("I:I").ColumnWidth = Column_Width_I
    Columns("J:J").ColumnWidth = Column_Width_J
    Columns("K:K").ColumnWidth = Column_Width_K
    Columns("L:L").ColumnWidth = Column_Width_L
    Columns("M:M").ColumnWidth = Column_Width_M
    Columns("N:N").ColumnWidth = Column_Width_N
'
End Sub
'
'
'CONVERT EACH IDENTIFIED COLUMN TO TEXT AND SET THE DEFINED NUMBER FORMAT
Private Sub Convert_Column_Number_Format(columnArray() As Variant, formatArray() As Variant)
'
    Dim columnRange As Range
    Dim i As Integer
'
   'Change the column letter to Excel's column designation format
    For i = 0 To UBound(columnArray)
        columnArray(i) = columnArray(i) & ":" & columnArray(i)
    Next i
'
   'Convert each identified column to text and set the defined number format
    For i = 0 To UBound(columnArray)
        Set columnRange = Range(columnArray(i))  'Define the range based on the parameter value
        columnRange.NumberFormat = formatArray(i)  'Set the number format
        columnRange.TextToColumns  'Convert the text to the defined number format
    Next i
'
End Sub
'
'
'DISPLAY EACH IDENTIFIED COLUMN USING BOLD TEXT
Private Sub Bold_Columns(wrkSheet As String, ws1BoldColumns As String, ws2BoldColumns As String, ws3BoldColumns As String)
'
    Dim ws As Worksheet
    Dim boldColumns As String
    Dim columnRange As Range
'
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
            boldColumns = ws1BoldColumns
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
            boldColumns = ws2BoldColumns
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
            boldColumns = ws3BoldColumns
        Case Else
           'Message box to be displayed if the Add_Header_Borders procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Formatting.Bold_Columns", Buttons:=vbCritical, Title:="Error Bolding Columns"
    End Select
'
    Set columnRange = ws.Range(boldColumns)
'
   'Make the identified columns bold
    columnRange.Font.Bold = True
'
End Sub
'
'
'SET VERTICAL AND HORIZONTAL ALIGNMENT ON SPECIFIED COLUMNS
Private Sub Format_Data(columnRange1 As String, columnRange2 As String, columnRange3 As String, columnRange4 As String)
'
    Dim range1 As Range, range2 As Range, range3 As Range, range4 As Range, totalColumnRange As Range
'
   'Center specific columns
    Set range1 = Range(columnRange1)
    Set range2 = Range(columnRange2)
    Set range3 = Range(columnRange3)
    Set range4 = Range(columnRange4)
    Set totalColumnRange = Union(range1, range2, range3, range4)
    totalColumnRange.HorizontalAlignment = xlCenter  'Center specified data text
    '
   'Apply vertical alignment to data cells
    With dataRange
        .VerticalAlignment = xlTop  'Top justify header text
        .WrapText = True  'Word wrap data text
    End With
'
   'Auto-fit all rows
    dataRange.EntireRow.AutoFit  'Auto-fit all data rows
'
End Sub
'
'
'HIGHLIGHT CELLS WITHIN A SPECIFIED RANGE
Private Sub Highlight_Cells(SearchString1 As String, highlightString1 As String, color1 As Integer, SearchString2 As String, highlightString2 As String, tintAndShade2 As Single)
'
    Dim searchRange1 As Range, highlightRange1 As Range, searchRange2 As Range, highlightRange2 As Range
    Dim highlightColumnCount1 As Integer, highlightColumnCount2 As Integer, i As Integer, j As Integer
    '
    Set searchRange1 = Range(SearchString1)
    Set highlightRange1 = Range(highlightString1)
    Set searchRange2 = Range(SearchString2)
    Set highlightRange2 = Range(highlightString2)
    '
    highlightColumnCount1 = highlightRange1.Columns.count
    highlightColumnCount2 = highlightRange2.Columns.count
    '
   'Case 1: Highlight the QA and Prod Remediation cells in bright yellow if either date is before today
    For i = 2 To rowCount
    '
       'Only process the cell if it contains a date
        If IsDate(searchRange1.Cells(i, 1)) Then
        '
           'Perform the following if the date listed in the cell is before today's date
            If DateValue(searchRange1.Cells(i, 1)) < Date Then
            '
               'Highlight all cells in that row within the specified highlight range
                For j = 1 To highlightColumnCount1
                '
                    highlightRange1.Cells(i, j).Interior.ColorIndex = 6
                '
                Next j
            '
            End If
        '
        End If
    '
    Next i
'
   'Case 2: Highlight the Remediation Plan cells in light red if any are blank
    For i = 2 To rowCount
    '
       'Check that the risk hasn't already been accepted
        If searchRange2.Cells(i, 4).Value = "" Then
        '
            For j = 1 To highlightColumnCount2
            '
               'Perform the following if any cell within the specified highlight range within the row is blank
                If searchRange2.Cells(i, j).Value = "" Then
                '
                   'Highlight the entire row within the specified highlight range
                    With searchRange2.Rows(i).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent2
                        .TintAndShade = tintAndShade2
                        .PatternTintAndShade = 0
                    End With
                    '
                    Exit For
                '
                End If
            '
            Next j
        '
        End If
    '
    Next i
'
End Sub
'
'
'FORMAT THE HEADER ROW WITH TEXT POSITION AND COLOR
Private Sub Format_Header()
'
   'Format header text - Centered, Bottom-aligned, Wrapped
    With headerRange
        .HorizontalAlignment = xlCenter  'Center header text
        .VerticalAlignment = xlBottom  'Bottom justify header text
        .WrapText = True  'Word wrap header text
        .Orientation = 0  'Header text should be displayed horizontally
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    '
   'Make header background dark blue
    With headerRange.Interior
        .Pattern = xlSolid  'Make header background a solid color
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2  'Make header background dark blue (over)
        .TintAndShade = 0  'Make header background dark blue (down)
        .PatternTintAndShade = 0
    End With
    '
   'Format header text bold and white
    With headerRange.Font
        .Bold = True  'Make headers bold
        .ThemeColor = xlThemeColorDark1  'Make header text white (over)
        .TintAndShade = 0  'Make header text white (down)
    End With
    '
   'Auto-fit the header row based on new names
    Rows("1:1").EntireRow.AutoFit
'
End Sub
'
'
'FREEZE THE HEADER ROW
Private Sub Freeze_Header()
'
   'Freeze header
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
'
End Sub
'
'
'ADD BORDERS TO THE INSIDE AND OUTSIDE OF THE HEADER ROW
Private Sub Add_Header_Borders(wrkSheet As String)
'
    Dim ws As Worksheet
    Dim currentHeaderRange As Range
'
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
        Case Else
           'Message box to be displayed if the Add_Header_Borders procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Add_Data_Borders", Buttons:=vbCritical, Title:="Error Adding Data Borders"
    End Select
'
   'Add header borders
    Set currentHeaderRange = ws.Range(headerRange.Address)
    currentHeaderRange.Borders(xlDiagonalDown).LineStyle = xlNone
    currentHeaderRange.Borders(xlDiagonalUp).LineStyle = xlNone
    currentHeaderRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    With currentHeaderRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentHeaderRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentHeaderRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentHeaderRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentHeaderRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
'
End Sub
'
'
'ADD BORDERS TO THE INSIDE AND OUTSIDE OF THE DATA SECTION
Private Sub Add_Data_Borders(wrkSheet As String)
'
    Dim ws As Worksheet
'
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
        Case Else
           'Message box to be displayed if the Add_Data_Borders procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Add_Data_Borders", Buttons:=vbCritical, Title:="Error Adding Data Borders"
    End Select
'
   'Add borders to data
    Dim currentDataRange As Range
'
    Set currentDataRange = ws.Range(dataRange.Address)
'
    currentDataRange.Borders(xlDiagonalDown).LineStyle = xlNone
    currentDataRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With currentDataRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentDataRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentDataRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentDataRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With currentDataRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With currentDataRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.35
        .Weight = xlThin
    End With
'
End Sub
'
'
'ADD FILTERS IN THE HEADER ROW FOR ALL COLUMNS
Private Sub Add_Filters(wrkSheet As String)
'
    Dim ws As Worksheet
    '
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
        Case Else
           'Message box to be displayed if the Add_Filters procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Add_Filter", Buttons:=vbCritical, Title:="Error Adding Filters"
    End Select
    '
   'Add auto-filter
    Dim currentHeaderRange As Range
    Set currentHeaderRange = ws.Range(headerRange.Address)
    currentHeaderRange.AutoFilter
'
End Sub
'
'
'FORMAT THE PAGE FOR PRINTING
Private Sub Page_Setup(wrkSheet As String)
'
    Dim ws As Worksheet
'
    Select Case wrkSheet
        Case "wrkSheet1"  'Worksheet 1
            Set ws = ws1
        Case "wrkSheet2"  'Worksheet 2
            Set ws = ws2
        Case "wrkSheet3"  'Worksheet 3
            Set ws = ws3
        Case Else
           'Message box to be displayed if the Page_Setup procedure was passed an invalid worksheet name
            MsgBox "Undefined worksheet in Mod_Vuln_Data_Changes.Page_Setup", Buttons:=vbCritical, Title:="Error During Page Setup"
    End Select
'
   'Set the title row(s) to print
    With ws.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With
'
   'Set the print area
    Dim currentTotalRange As Range
    Set currentTotalRange = ws.Range(totalRange.Address)
    ws.PageSetup.PrintArea = currentTotalRange.Address
'
   'Set the page header and footer
    With ws.PageSetup
        .LeftHeader = "&G"  'Left header contains a graphic
        .CenterHeader = "" & Chr(10) & ""  'Center header is blank
        .RightHeader = "&""Calibri,Bold""&14BBDI Open Vulnerabilities" & Chr(10) & "&A"  'Right header contains the project name and the worksheet name, all in Font Name: Calibri, Font Style: Bold, Font Size: 14
        .LeftFooter = "Broadridge Confidential"  'Left footer text
        .CenterFooter = "Page &P of &N"  'Center footer contains the current page number out of the total page numbers
        .RightFooter = "&D"  'Right footer text
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.86)
        .BottomMargin = Application.InchesToPoints(0.69)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True  'Center the table horizontally on the page
        .CenterVertically = False  'Top justify the table on the page
        .Orientation = xlLandscape  'Set the page to print in a landscape layout
        .Draft = False
        .PaperSize = xlPaperLetter  'Set the paper to 8.5"x11"
        .FirstPageNumber = xlAutomatic  'Set page numbering to start at "1"
        .Order = xlDownThenOver
        .BlackAndWhite = False  'Allow color printing by default if able
        .Zoom = False
        .FitToPagesWide = 1  'Set to fit all columns on a single page
        .FitToPagesTall = 99  'Set to allow an unlimited amount of pages to print all rows
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False  'Keep the header and footer at their original size, regardless of the content shrinking to fit
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
'
End Sub
'
'
'1)RESET THE FIND FUNCTION, 2) SET ALL WORKSHEETS BACK TO THEIR CELL "A1", 3) DISPLAY THE FIRST WORKSHEET
Private Sub Reset_Find_and_Cell_Selection()
'
   'Reset the Find function to clear the "Match entire cell contents" checkbox
    Range("A:A").Find(What:="", LookAt:=xlPart, MatchCase:=False, SearchFormat:=False).Activate
    '
   'Move back to the beginning of the third sheet
    ws3.Activate
    Range("$A$2").Activate
    Range("$A$1").Activate
'
   'Move back to the beginning of the second sheet
    ws2.Activate
    Range("$A$1").Activate
'
   'Move back to the beginning of the first sheet
    ws1.Activate
    Range("$A$1").Activate
'
End Sub
