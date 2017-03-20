'This function is a sample to quick expend the data of curve,
'by Dazhong 26/06/2016
Option Explicit

Sub ChartRangeAdd()
    On Error Resume Next
    Dim oCht As Chart
    Dim s As Integer
    Dim sTmp As String
    Dim split_formula As Variant
    Set oCht = ActiveChart
    oCht.Select
    Dim x_range, y_range As String
    For s = 1 To oCht.SeriesCollection.Count
        sTmp = oCht.SeriesCollection(s).Formula 'Get the forumlar of the curve
        split_formula = Split(sTmp, ",")
        x_range = split_formula(1)
        y_range = split_formula(2)
        Dim end_row As Integer
        end_row = GetNums(x_range)(1)
        oCht.SeriesCollection(s).Formula = Replace(oCht.SeriesCollection(s).Formula, CStr(end_row), "170")
        Debug.Print oCht.SeriesCollection(s).Formula
    Next s
End Sub
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
