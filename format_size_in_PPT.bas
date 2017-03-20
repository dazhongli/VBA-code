Sub format_size()

    Local_Left = 0.2
    Local_Top = 0.2
    Local_Width = 26.72
    Local_Height = 7
    factor = 28.338 'this factor is used to convert to pixel to the cm
    With ActiveWindow.Selection.ShapeRange
        .Left = Local_Left * factor
        .Top = Local_Top * factor
        .Width = Local_Width * factor
        .Height = Local_Height * factor
    End With
End Sub

