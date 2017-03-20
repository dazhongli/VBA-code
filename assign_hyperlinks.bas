'this function below shows an example of how the handle the hyper link
Sub assgin_hyperlinks()
    Dim ws As Worksheet
    Set ws = Sheets("Index")
    Dim rng_instrument_ID As Range
    Set rng_instrument_ID = ws.Range("Instrument_ID")
    Dim cell As Range
    For Each cell In rng_instrument_ID
        ws.Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="Master!" & Cells(2, cell.Offset(0, 1)).Address, TextToDisplay:=cell.Value
    Next cell
    
End Sub
