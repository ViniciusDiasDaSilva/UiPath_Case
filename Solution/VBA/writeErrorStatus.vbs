Sub writeErrorStatus(invoice As String, VendorId As String, status As String)
    Set celula = Columns("A").Find(What:=invoice, LookAt:=xlWhole)
    If Not celula Is Nothing Then
        primeiraOcorrencia = celula.Address
        Do
            If Cells(celula.Row, "D").Value = VendorId Then
                Cells(celula.Row, "O").Value = status
                
                Exit Do
            End If

            Set celula = Columns("A").FindNext(celula)
        Loop While celula.Address <> primeiraOcorrencia
    End If

End Sub
