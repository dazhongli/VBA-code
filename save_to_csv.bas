'This function save the selection of a excel file to a csv file
'taken from https://stackoverflow.com/questions/32237153/exporting-selection-to-csv
'Modified by Dazhong Nov 6 2017
Sub save_to_csv()

ChDrive "P:" '// <~~ change current drive to P:\
Dim copyRng As Excel.Range
Dim ThisWB  As Excel.Workbook
Dim OtherWB As Excel.Workbook
Dim sName   As String

'// set reference to the 'Master' workbook
Set ThisWB = ActiveWorkbook

'// assign selected range to 'copyRng'
Set copyRng = Application.InputBox(Prompt:="Select range to convert to CSV", Type:=8)

'// If the user selected a range, then proceed with rest of code:
If Not copyRng Is Nothing Then
    '// Create a new workbook with 1 sheet.
    Set OtherWB = Workbooks.Add(1)

    '// Get A1, then expand this 'selection' to the same size as copyRng. 
    '// Then assign the value of copyRng to this area (similar to copy/paste)
    OtherWB.Sheets(1).Range("A1").Resize(copyRng.Rows.Count, copyRng.Columns.Count).Value = copyRng.Value

    '// Get save name for CSV file.
    sName = Application.GetSaveAsFilename(FileFilter:="CSV files (*.csv), *.csv")

    '// If the user entered a save name then proceed:
    If Not LCase(sName) = "false" Then
        '// Turn off alerts
        Application.DisplayAlerts = False
        '// Save the 'copy' workbook as a CSV file
        OtherWB.SaveAs sName, xlCSV
        '// Close the 'copy' workbook
        OtherWB.Close
        '// Turn alerts back on
        Application.DisplayAlerts = True
    End If

    '// Make the 'Master' workbook the active workbook again
    ThisWB.Activate

    MsgBox "Conversion complete", vbInformation
End If

End Sub
