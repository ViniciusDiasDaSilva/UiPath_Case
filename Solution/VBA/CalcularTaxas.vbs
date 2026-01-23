Sub CalcularTaxas(UF As String)

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Range("G27").Value > 2000000 Then
        Range("G28").Value = Range("G27").Value * 0.1
    End If
    
    Select Case estado
        Case "SP"
            taxa = 5
        Case "RJ"
            taxa = 2
        Case "MG"
            taxa = 1
        Case Else
            taxa = 0
    End Select
    Range("G29").Value = taxa

    ws.Cells.EntireColumn.AutoFit

    With ws.PageSetup
        .Orientation = xlLandscape      ' Paisagem
        .Zoom = False                   ' Desativa zoom manual
        .FitToPagesWide = 1             ' 1 página de largura
        .FitToPagesTall = 1             ' 1 página de altura
    End With

End Sub
