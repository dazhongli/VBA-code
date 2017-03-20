 Sub ChangeChartData()
 start_time = Timer

    Application.DisplayAlerts = False
    Dim pptChart As Chart
    Dim pptChartData As ChartData
    Dim pptWorkbook As Object
    Dim sld As Slide
    Dim Shp As Shape
    Dim xlobject As Object
    Set xlobject = CreateObject("Excel.application")
    Dim PreFileName, CurrentFileName
    PreFileName = "Null"
    For Each sld In ActivePresentation.Slides
        For Each Shp In sld.Shapes
            If Shp.HasChart Then
            Call SetChartSizeAndPosition(Shp, -0.34, 1.68, 27.5, 16.89)
            End If
        Next
    Next
    On Error Resume Next
    xlobject.Close True
    Set pptWorkbook = Nothing
    Set pptChartData = Nothing
    Set pptChart = Nothing
    Application.DisplayAlerts = True
    MsgBox "VWP PPT Update completed!"
    end_time = Timer
    MsgBox "Total Time Elapsed = " & Format(end_time - start_time, "0") & " Seconds!", vbInformation
End Sub


Private Function WorkbookIsOpen(wbname) As Boolean
'   Returns TRUE if the workbook is open
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function

Sub format_chart()
    For Each sld In ActivePresentation.Slides
    If Shp.HasChart Then
        Set pptChart = Shp.Chart
        pptChart.Axes(xlValue).MinimumScale = 0
    End If
End Sub


Sub SetChartSizeAndPosition(Shp As Shape, Left As Single, Top As Single, Width As Single, Height As Single)
factor = 28.338 'this factor is used to convert to pixel to the cm
With Shp
     .Left = Left * factor
     .Top = Top * factor
     .Width = Width * factor
     .Height = Height * factor
 End With

End Sub
