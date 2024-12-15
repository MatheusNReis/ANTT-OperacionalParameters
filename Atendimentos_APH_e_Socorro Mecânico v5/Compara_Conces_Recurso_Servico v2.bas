Function Compara_Conces_Recurso_Serviço(ROpWorksheet As Object, lastRowROpWorksheet As Long, ByVal Concessionaria As String, ByVal Recurso As String) As String
' Concessionaria e Recurso são valores das células B e F de objWorksheet
    
    Dim rng As Range
    Dim cell As Range

    ' Procurar valor correspondente na coluna E da planilha ROpWorksheet
    Set rng = ROpWorksheet.Range("E2:E" & lastRowROpWorksheet)

    For Each cell In rng 'Percorrendo a planilha ROpWorksheet
        If (cell.Offset(0, -4).Value = Concessionaria) And (cell.Offset(0, 1).Value = Recurso) Then
            Compara_Conces_Recurso_Serviço = cell.Value 'Ambulância C, Guincho Leve, etc.
            Exit Function
        End If
    Next cell

    'Caso não seja encontrado valor correspondente, a variável mantém o mesmo nome que já estava em Recurso
    Compara_Conces_Recurso_Serviço = Recurso
    
End Function
