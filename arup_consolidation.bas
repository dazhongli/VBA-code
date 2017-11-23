'The subroutines down here are simply from the spreadsheets for the consolidation analysis. The spreadsheet developed for the HKBCF. 
'No changes are made, a wrapper is made for easier run. 

Sub run()
    Call Primary
    Call Secondary
End Sub

Private Sub Primary()
'
' Primary Macro
'
Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets("primary")
With ws
.Range("R170").Select
    .Range("R170").GoalSeek Goal:=0, ChangingCell:=.Range("F168")
    
    .Range("R366").GoalSeek Goal:=0, ChangingCell:=.Range("F364")
    
    
    .Range("R677").GoalSeek Goal:=0, ChangingCell:=.Range("F676")
    
    .Range("R944").GoalSeek Goal:=0, ChangingCell:=.Range("F943")
     
 
    .Range("R164").Select
End With
End Sub

Private Sub Secondary()
'
' Secondary Macro
'
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("secondary")
    
    With ws
        If (.Range("P195") >= (0.2) And .Range("P195") < 0.3) Then
        .Range("G195") = "0.2-0.3"
        Else
        If (.Range("P195") >= 0.3 And .Range("P195") < 0.4) Then
        .Range("G195") = "0.3-0.4"
        Else
        If (.Range("P195") >= 0.4 And .Range("P195") < 0.5) Then
        .Range("G195") = "0.4-0.5"
        Else
        If .Range("P195") < 0.2 Then
        .Range("G195") = "<0.2"
        End If
        End If
        End If
        End If
        
        If (.Range("P196") >= (0.2) And .Range("P196") < 0.3) Then
        .Range("G196") = "0.2-0.3"
        Else
        If (.Range("P196") >= 0.3 And .Range("P196") < 0.4) Then
        .Range("G196") = "0.3-0.4"
        Else
        If (.Range("P196") >= 0.4 And .Range("P196") < 0.5) Then
        .Range("G196") = "0.4-0.5"
        Else
        If .Range("P196") < 0.2 Then
        .Range("G196") = "<0.2"
End If
End If
End If
End If
    
    If (.Range("P197") >= (0.2) And .Range("P197") < 0.3) Then
.Range("G197") = "0.2-0.3"
Else
If (.Range("P197") >= 0.3 And .Range("P197") < 0.4) Then
.Range("G197") = "0.3-0.4"
Else
If (.Range("P197") >= 0.4 And .Range("P197") < 0.5) Then
.Range("G197") = "0.4-0.5"
Else
If .Range("P197") < 0.2 Then
.Range("G197") = "<0.2"
End If
End If
End If
End If
    
    If (.Range("P198") >= (0.2) And .Range("P198") < 0.3) Then
.Range("G198") = "0.2-0.3"
Else
If (.Range("P198") >= 0.3 And .Range("P198") < 0.4) Then
.Range("G198") = "0.3-0.4"
Else
If (.Range("P198") >= 0.4 And .Range("P198") < 0.5) Then
.Range("G198") = "0.4-0.5"
Else
If .Range("P198") < 0.2 Then
.Range("G198") = "<0.2"
End If
End If
End If
End If
    
    If (.Range("P199") >= (0.2) And .Range("P199") < 0.3) Then
.Range("G199") = "0.2-0.3"
Else
If (.Range("P199") >= 0.3 And .Range("P199") < 0.4) Then
.Range("G199") = "0.3-0.4"
Else
If (.Range("P199") >= 0.4 And .Range("P199") < 0.5) Then
.Range("G199") = "0.4-0.5"
Else
If .Range("P199") < 0.2 Then
.Range("G199") = "<0.2"
End If
End If
End If
End If
    
    If (.Range("P200") >= (0.2) And .Range("P200") < 0.3) Then
.Range("G200") = "0.2-0.3"
Else
If (.Range("P200") >= 0.3 And .Range("P200") < 0.4) Then
.Range("G200") = "0.3-0.4"
Else
If (.Range("P200") >= 0.4 And .Range("P200") < 0.5) Then
.Range("G200") = "0.4-0.5"
Else
If .Range("P200") < 0.2 Then
.Range("G200") = "<0.2"
End If
End If
End If
End If
    
        If (.Range("P201") >= (0.2) And .Range("P201") < 0.3) Then
    .Range("G201") = "0.2-0.3"
    Else
    If (.Range("P201") >= 0.3 And .Range("P201") < 0.4) Then
    .Range("G201") = "0.3-0.4"
    Else
    If (.Range("P201") >= 0.4 And .Range("P201") < 0.5) Then
    .Range("G201") = "0.4-0.5"
    Else
    If .Range("P201") < 0.2 Then
    .Range("G201") = "<0.2"
    End If
    End If
    End If
    End If
        
        If (.Range("P202") >= (0.2) And .Range("P202") < 0.3) Then
    .Range("G202") = "0.2-0.3"
    Else
    If (.Range("P202") >= 0.3 And .Range("P202") < 0.4) Then
    .Range("G202") = "0.3-0.4"
    Else
    If (.Range("P202") >= 0.4 And .Range("P202") < 0.5) Then
    .Range("G202") = "0.4-0.5"
    Else
    If .Range("P202") < 0.2 Then
    .Range("G202") = "<0.2"
    End If
    End If
    End If
    End If
    
        If (.Range("P203") >= (0.2) And .Range("P203") < 0.3) Then
    .Range("G203") = "0.2-0.3"
    Else
    If (.Range("P203") >= 0.3 And .Range("P203") < 0.4) Then
    .Range("G203") = "0.3-0.4"
    Else
    If (.Range("P203") >= 0.4 And .Range("P203") < 0.5) Then
    .Range("G203") = "0.4-0.5"
    Else
    If .Range("P203") < 0.2 Then
    .Range("G203") = "<0.2"
    End If
    End If
    End If
    End If
    
        If (.Range("P204") >= (0.2) And .Range("P204") < 0.3) Then
    .Range("G204") = "0.2-0.3"
    Else
    If (.Range("P204") >= 0.3 And .Range("P204") < 0.4) Then
    .Range("G204") = "0.3-0.4"
    Else
    If (.Range("P204") >= 0.4 And .Range("P204") < 0.5) Then
    .Range("G204") = "0.4-0.5"
    Else
    If .Range("P204") < 0.2 Then
    .Range("G204") = "<0.2"
    End If
    End If
    End If
    End If
    End With
End Sub




