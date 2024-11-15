Sub CalculaDiferenca()
    
'Atendimento_APH_e_Socorro Mecânico.xlsm

'Test calculation conversion for different dates

'Testa conversão de cálculo para diferentes datas

'Created by Matheus Nunes Reis on 05/10/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/05782e3211f77a3b453fbcf02b3e031513031ecf/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long

    ' Define a planilha
    Set ws = ThisWorkbook.Sheets("Teste")

    ' Loop da linha 2 até a linha 57
    For i = 2 To 57
        ' Define a célula na coluna M (13ª coluna)
        Set rng = ws.Cells(i, 14)

        ' Calcula a diferença entre a coluna L (12ª coluna) e a coluna K (11ª coluna) em dias
        rng.Value = CDbl(ws.Cells(i, 12).Value - ws.Cells(i, 11).Value)
        Next i
End Sub
