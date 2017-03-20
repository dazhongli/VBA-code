'This function is to plot hte CPT directly for the ASCII file
'By Dazhong 13 May 2016

Option Explicit

Sub add_a_CPT_record()
    'we disable the automatic calculation first
    Application.Calculation = xlCalculationManual
    'key words
    Dim str_date_flag, str_seperator, date_of_testing As String
    str_date_flag = "Date of Testing"
    str_seperator = ":"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim fd As FileDialog
    Dim sSourcefile, stargetfile As String
    Dim filepicked As Boolean
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim target_location As String
    target_location = "Your location where file is to be saved"
    Dim vwp_filename, sp_filename As String
    fd.InitialFileName = "Initial Path for the source file"
    filepicked = fd.Show
    
    If Not filepicked Then
        Exit Sub
    End If
    
    Dim n_file_selected As Integer
    n_file_selected = fd.SelectedItems.Count
    sSourcefile = fd.SelectedItems(1)
    
    Dim selected_filename_full, selected_filename_short As String
    selected_filename_full = sSourcefile
    
    selected_filename_short = fso.GetBaseName(selected_filename_full)
    ThisWorkbook.Sheets("Template").Copy , After:=Worksheets(Worksheets.Count)
    ThisWorkbook.Sheets(Worksheets.Count).Name = selected_filename_short
    '----------------------Open the file as file stream----------------------------------------------
    Dim objTextStream
    Set objTextStream = fso.OpenTextFile(selected_filename_full, ForReading)
    Dim txtLine As String
    '---find the date of the testing------------------------
    Do While objTextStream.AtEndOfStream <> True
        txtLine = objTextStream.ReadLine()
        If InStr(1, txtLine, str_date_flag, vbTextCompare) > 0 Then 'if we found the keywords for the date
            Exit Do
        End If
    Loop
    'Trim the line to clean string
    txtLine = Trim(txtLine)
    Dim n As Integer
    n = InStr(1, txtLine, str_seperator, vbTextCompare)
    date_of_testing = Right(txtLine, Len(txtLine) - n) 'we got the date fo testing here~~
    
    'we will read the data here
    If n_file_selected > 1 Then ' if we have selected a file for the reading, do it here
        Set objTextStream = fso.OpenTextFile(fd.SelectedItems(2), ForReading)
        Do While objTextStream.AtEndOfStream <> True
            txtLine = objTextStream.ReadLine()
            If InStr(1, txtLine, "Data table", vbTextCompare) > 0 Then 'if we found the keywords for the date
                Exit Do
            End If
        Loop
    End If
    'read the title
    txtLine = objTextStream.ReadLine()
    'read the unit
    txtLine = objTextStream.ReadLine()
    'read first line
    txtLine = objTextStream.ReadLine()

    '-----------get a reference to the excel--------------------------------------
    Dim ws As Worksheet
    Dim Raw_data_rng As Range
    Set ws = ThisWorkbook.Sheets(selected_filename_short)
    'assign the range
    Set Raw_data_rng = ws.Range("B43:H5000")
    ws.Range("E25") = "_" & date_of_testing
    Raw_data_rng.ClearContents
    Dim n_line As Integer ' keep a track of the line number
    n_line = 1
    '--------------------------------we will read some serious stuff here ---------------------------
    'We will need a another loop here
    Do While objTextStream.AtEndOfStream <> True
        Dim LineItems As Variant
        txtLine = objTextStream.ReadLine()
        LineItems = Split(Trim(txtLine), "   ")
        Dim i As Integer
        For i = 0 To 6
            On Error GoTo Error_handler
            Raw_data_rng(n_line, i + 1) = LineItems(i)
        Next i
        n_line = n_line + 1
    Loop
Error_handler:
Application.Calculation = xlCalculationAutomatic
    ' We add the curve in the plot
    Dim Plot_chart As ChartObject
    Set Plot_chart = ThisWorkbook.Sheets("Plots").ChartObjects(1)
    Plot_chart.Activate
    ActiveChart.SeriesCollection.NewSeries
    Dim LastIndex As Integer
    LastIndex = ActiveChart.SeriesCollection.Count
    ActiveChart.SeriesCollection(LastIndex).Name = "=" & "'" & selected_filename_short & "'" & "!F19"
    ActiveChart.SeriesCollection(LastIndex).XValues = "=" & "'" & selected_filename_short & "'" & "!$W$44:$W$6000"
    ActiveChart.SeriesCollection(LastIndex).Values = "=" & "'" & selected_filename_short & "'" & "!$X$44:$X$6000"
    
End Sub
