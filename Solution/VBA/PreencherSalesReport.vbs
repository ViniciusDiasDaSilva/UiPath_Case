Sub PreencherSalesReport(dt_Sales As Object, dt_Vendor As Object, vendor_id As String)

    Dim ws As Worksheet
    Dim startRow As Long: startRow = 18
    Dim i As Long
    ws.Name = vendor_id
    ' Limpar campos
    ws.Range("B7:B13,C7,G27:G31").ClearContents
    
    With dt_Vendor.Rows(0)
        ws.Range("B9").Value = .Item("Vendor Name")
        ws.Range("B12").Value = .Item("Phone Number")
        ws.Range("B13").Value = .Item("e-Mail")
    
    ' Cabeçalho (primeira linha do DataTable)
    With dt_Sales.Rows(0)
        ws.Range("B7").Value = .Item("VENDOR ID")
        ws.Range("B10").Value = .Item("Endereço")
        ws.Range("B11").Value = .Item("Bairro") & "," & .Item("Localidade/UF")
        ws.Range("C7").Value = .Item("DATE")
    End With
    
    ' Itens
    For i = 0 To dt.Rows.Count - 1
        ws.Cells(startRow + i, "B").Value = dt.Rows(i)("INVOICE") & "/" & dt.Rows(i)("ITEM TYPE")
        ws.Cells(startRow + i, "D").Value = dt.Rows(i)("Valor Unitário (BRL)")
        ws.Cells(startRow + i, "F").Value = dt.Rows(i)("QTY")
        
    Next i

    ' Totais
    ws.Range("G27").Formula = "=SUM(G18:G" & startRow + dt.Rows.Count - 1 & ")"
    ws.Range("G28").Formula = "=IF(G27>2000000,G27*0.1,0)"
    uf = Split(dt.Rows(0)("STATE"), "/")
    ws.Range("G29").Value = _
        IIf(uf = "SP", 0.05, _
        IIf(uf = "RJ", 0.02, _
        IIf(uf = "MG", 0.01, 0)))

    ws.Range("G30").Formula = "=G27*G29"
    ws.Range("G31").Formula = "=G27-G28+G30"

End Sub
