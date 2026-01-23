Sub PreencherCabecalhoSalesReport(pVendorId As String, pVendorName As String, pPhone As String, pEmail As String, pEndereco As String, pBairro As String, pCidadeUF As String)

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim startRow As Long: startRow = 18
    Dim i As Long


    On Error Resume Next
    ws.Name = pVendorId
    On Error GoTo 0

    ws.Range("C7").Value = Format(Date, "MM/dd/yyyy")
    ws.Range("B9").Value = pVendorName 
    ws.Range("B12").Value = pPhone 
    ws.Range("B13").Value = pEmail 

    
    ws.Range("B7").Value = pVendorId
    ws.Range("B10").Value = pEndereco
    ws.Range("B11").Value = pBairro & ", " & pCidadeUF

    ws.Range("B35").Value = Replace(ws.Range("B35").Value, "MM/DD/YYYY", Format(Date, "MM/dd/yyyy"))
End Sub
