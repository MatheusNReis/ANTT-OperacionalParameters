'Atendimento_APH_e_Socorro Mecânico.xlsm

'Function used in Module1 to relate resource with operational service in the Resource Spreadsheet according to
'concessionaire and vehicle type

'Função utilizada no Módulo1 para relacionar recurso com serviço operacional na Planilha de recursos conforme
'concessionária e tipo de veículo

'Version: 1.0

'Created by Matheus Nunes Reis on 20/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


Function Compara_Conces_Recurso_Serviço(ROpWorksheet As Object, lastRowROPWorksheet As Long, ByVal Concessionaria As String, ByVal Recurso As String, ByVal Serviço As String) As String
' Concessionaria e Recurso são valores das células B e F de objWorksheet
    
    Dim rng As Range
    Dim cell As Range

    ' Procurar valor correspondente na coluna E da planilha ROpWorksheet
    Set rng = ROpWorksheet.Range("E2:E" & lastRowROPWorksheet)

    For Each cell In rng 'Percorrendo a planilha ROpWorksheet
        If (cell.Offset(0, -4).Value = Concessionaria) And (cell.Offset(0, 1).Value = Recurso) And (cell.Offset(0, 2).Value = Serviço) Then
            Compara_Conces_Recurso_Serviço = cell.Value 'Ambulância C, Guincho Leve, etc.
            Exit Function
        End If
    Next cell

    'Caso não seja encontrado valor correspondente, a variável = vazio
    Compara_Conces_Recurso_Serviço = ""
    
End Function
