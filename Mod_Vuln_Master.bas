Attribute VB_Name = "Mod_Vuln_Master"
'Macro: Open Application Vulnerabilities Formatter
'
'This is the master module that calls the configuration file, and the data change, formatting and workbook procedures
'
'Ordered list of procedures performed by the "Open Application Vulnerabilities Formatter" macro:
'
'1. Make any necessary changes to shared data
    'a. Convert each identified column to text and set the defined number format
    'b. Change header names
    'c. Set defined column widths
    'd. Format the data
    'e. Find/replace blank RP Owner data
    'f. Add Risk Treatment data based on CVM
    'g. Highlight specified data
'2. Format the header
    'a. Set header background and font color
    'b. Freeze header
'3. Copy Sheet1 to a new Sheet2 and a new Sheet3 and name and resize them all
'4. Format the data in Sheet1
    'a. Enbolden the first column
    'b. Sort data
    'c. Add borders to header
    'd. Add borders to data
    'e. Add data filters across all headers
    'f. Add Page Setup information for printing
'5. Format the data in Sheet2
    'a. Move column
    'b. Enbolden the first two columns
    'c. Sort data
    'd. Add borders to header
    'f. Add borders to data
    'f. Add data filters across all headers
    'g. Add Page Setup information for printing
'6. Format the data in Sheet3
    'a. Move column
    'b. Enbolden the first two columns
    'c. Sort data
    'd. Add borders to header
    'f. Add borders to data
    'f. Add data filters across all headers
    'g. Add Page Setup information for printing
'7. Save, e-mail, and close the file
    'a. Reset all sheets to top left cell
    'b. Save the file
    'c. E-mail the file
    'd. Close file
    'e. Clear the objects used by the macro from memory
    'f. Alert when completed successfully
'
'
'For use by Task Counter procedure
Public taskCounter As Integer
'
'
Public Sub Open_Application_Vulnerabilities_Formatter()
'
    Application.ScreenUpdating = False
    taskCounter = 1
    '
   'Define any global variables
    Application.Run "Mod_Vuln_Workbook.Define_Global_Variables"
        Task_Counter
   '1. Make any necessary changes to shared data
    Application.Run "Mod_Vuln_Configs.Convert_Column_Number_Format_Config"  'a. Convert each identified column to text and set the defined number format
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Change_Header_Names_Config"  'b. Change header names
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Set_Column_Widths_Config"  'c. Set defined column widths
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Format_Data_Config"  'd. Format the data
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Replace_Data_Config"  'e. Find/replace blank RP Owner data
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Add_External_Data_Config"  'f. Add Risk Treatment data based on CVM (Use Add_External_Data_Config to load data from a different workbook)
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Highlight_Cells_Config"  'g. Highlight specified data
        Task_Counter
   '2. Format the header
    Application.Run "Mod_Vuln_Formatting.Format_Header"  'a. Set header background and font color
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Freeze_Header"  'b. Freeze header
        Task_Counter
   '3. Copy Sheet1 to a new Sheet2 and a new Sheet3 and name and resize them all
    Application.Run "Mod_Vuln_Configs.Create_New_Sheets_Config"
        Task_Counter
   '4. Format the data in Sheet1
    Application.Run "Mod_Vuln_Configs.Bold_Columns_Config", "wrkSheet1"  'a. Enbolden the first column
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Sort_Data_Config", "wrkSheet1"  'b. Sort data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Header_Borders", "wrkSheet1"  'c. Add borders to header
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Data_Borders", "wrkSheet1"  'd. Add borders to data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Filters", "wrkSheet1"  'e. Add data filters across all headers
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Page_Setup", "wrkSheet1"  'f. Add Page Setup information for printing
        Task_Counter
   '5. Format the data in Sheet2
    Application.Run "Mod_Vuln_Data_Changes.Move_Column", "wrkSheet2"  'a. Move column
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Bold_Columns_Config", "wrkSheet2"  'b. Enbolden the first two columns
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Sort_Data_Config", "wrkSheet2"  'c. Sort data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Header_Borders", "wrkSheet2"  'd. Add borders to header
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Data_Borders", "wrkSheet2"  'e. Add borders to data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Filters", "wrkSheet2"  'f. Add data filters across all headers
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Page_Setup", "wrkSheet2"  'g. Add Page Setup information for printing
        Task_Counter
   '6. Format the data in Sheet3
    Application.Run "Mod_Vuln_Data_Changes.Move_Column", "wrkSheet3"  'a. Move column
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Bold_Columns_Config", "wrkSheet3"  'b. Enbolden the first two columns
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Sort_Data_Config", "wrkSheet3"  'c. Sort data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Header_Borders", "wrkSheet3"  'd. Add borders to header
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Data_Borders", "wrkSheet3"  'e. Add borders to data
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Add_Filters", "wrkSheet3"  'f. Add data filters across all headers
        Task_Counter
    Application.Run "Mod_Vuln_Formatting.Page_Setup", "wrkSheet3"  'g. Add Page Setup information for printing
        Task_Counter
   '7. Save, e-mail, and close the file
    Application.Run "Mod_Vuln_Formatting.Reset_Find_and_Cell_Selection"  'a. Reset Find settings and select the top left cell of all sheets
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Save_Workbook_Config"  'b. Save the file
        Task_Counter
    Application.Run "Mod_Vuln_Configs.Mail_Workbook_Config"  'c. E-mail the file
        Task_Counter
    Application.Run "Mod_Vuln_Workbook.Close_Workbook"  'd. Close the file
        Task_Counter
    Application.Run "Mod_Vuln_Workbook.Cleanup"  'e. Clear the objects used by the macro from memory
        Task_Counter
    Application.Run "Mod_Vuln_Workbook.Completed_Successfully_Alert"  'f. Alert when completed successfully
        Application.StatusBar = ""
    '
    Application.ScreenUpdating = True
'
End Sub
'
'
'Progress display within the Status Bar
Private Sub Task_Counter()
'
    Dim totalTaskCount As Integer
    totalTaskCount = 36
    
    Application.StatusBar = "Task " & taskCounter & " of " & totalTaskCount & " tasks complete (" & Round(taskCounter / totalTaskCount * 100, 0) & "%)."
    taskCounter = taskCounter + 1
'
End Sub
'
'
'Procedures/enhancements to add:
  '1. Add a Legend as the last sheet in the workbook
    'If no worksheet called "Legend", Then
        'Create new worksheet object as the last worksheet
        'Name new owrksheet object "Legend"
        'Add header cell
        'Format header cell
    'Set new worksheet object
    'Count rows in Column A
    'Set new range object as the first unused cell in Column A
    'Add legend text to that cell
    'Apply any specified cell shading to that cell
    'Apply any text coloring to that cell
    'Set new range object to data range (A2:A[last used cell])
    'Add borders to that range
    'Auto-fit Column A
    'Select cell A1
    'Selection worksheet 1, cell A1
  '2. Add a procedure to create a Summary sheet that includes Application Name, App ID, Process Owner, Count of Open Vulnerabilities, _
      'Remediation Plans Missing?, and counts of each vulnerability by Severity
  '3. Add error handling to the Add_External_Data, Save_Workbook, Mail_Workbook, and Close_Workbook procedures
'
'
