Sub label_bubble_chart()
'
    Dim starting_row, end_row, column_number As Integer
    offset_column = CInt(InputBox("Please input the offset for the label,negative - left, positive - right, referenced to x value"))
    Dim n_series As Integer ' number_of seies in the bubble chart
    n_series = ActiveChart.SeriesCollection.Count
    Dim str_formula As String
    Dim var As Variant
    For i = 1 To n_series
        str_formula = ActiveChart.SeriesCollection(i).Formula
        var = Split(str_formula, ",")
        Dim temp As String
        Dim sheet_name As String
        sheet_name = Replace(Split(var(2), "!")(0), "'", "")
        temp = Split(var(2), "!")(1)
        Dim str_column As String
        'The column Letter would be the one between two dollar signs
        str_column = Mid(temp, FindN("$", temp, 1) + 1, FindN("$", temp, 2) - FindN("$", temp, 1) - 1)
        Dim row_number As Variant
        
        row_number = GetNums(temp)
        
        
        starting_row = row_number(0)
        end_row = row_number(1)
        column_number = Sheets(1).Range(str_column & 1).Column
        On Error GoTo Error_handler
        On Error GoTo 0
        Dim j As Integer
        j = 1
        For j = 1 To ActiveChart.SeriesCollection(i).Points.Count
            ActiveChart.SeriesCollection(i).Points(j).DataLabel.Formula = "=" & "'" & sheet_name & "'" & "!" & Sheets(sheet_name).Cells(starting_row + j - 1, column_number + offset_column).Address
        Next j
    Next i
Error_handler:
End Sub
Function FindN(sFindWhat As String, sInputString As String, N As Integer) As Integer
     Dim j As Integer
     Application.Volatile
     FindN = 0
     For j = 1 To N
         FindN = InStr(FindN + 1, sInputString, sFindWhat)
         If FindN = 0 Then Exit For
    Next
End Function

Function GetNums(ByVal strIn As String) As Variant  'Array of numeric strings
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .Pattern = "[^\d]+"
        NumStr = .Replace(strIn, " ")
    End With

    GetNums = Split(Trim(NumStr), " ")
End Function
