Sub PreencherCabecalhoSalesReport(pVendorId As String, pVendorName As String, pPhone As String, pEmail As String, pEndereco As String, pBairro As String, pCidadeUF As String)

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim startRow As Long: startRow = 18
    Dim i As Long

    ' Renomear aba
    On Error Resume Next
    ws.Name = Left(vendor_id, 31)
    On Error GoTo 0

   
    ws.Range("B9").Value = pVendorName 
    ws.Range("B12").Value = pPhone 
    ws.Range("B13").Value = pEmail 

    
    ws.Range("B7").Value = pVendorId
    ws.Range("B10").Value = pEndereco
    ws.Range("B11").Value = pBairro & ", " & pCidadeUF


End Sub
