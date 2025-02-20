Function Compara_Conces_Recurso_Serviço(ROpWorksheet As Object, lastRowROpWorksheet As Long, ByVal Concessionaria As String, ByVal Recurso As String) As String
' Concessionaria e Recurso são valores das células B e F de objWorksheet

'Atendimento_APH_e_Socorro Mecânico.xlsm

'Function used in Module1 to relate resource with operational service in the Resource Spreadsheet according to
'concessionaire and vehicle type, keeping the same resource name when no match found

'Função utilizada no Módulo1 para relacionar recurso com serviço operacional na Planilha de recursos conforme
'concessionária e tipo de veículo, mantendo o mesmo nome de recurso quando não encontrado correspondente

'Version: 2.0

'Created by Matheus Nunes Reis on 16/09/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/05782e3211f77a3b453fbcf02b3e031513031ecf/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    
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
