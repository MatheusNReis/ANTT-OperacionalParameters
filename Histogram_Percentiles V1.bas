Sub Histograma_Percentis()
    
    'Calcula Percentis de lista e gera gráfico de histograma com frequências acumuladas
    'Dois métodos para determinação do número de blocos (Fórmula de sturges ou Raiz Quadrada)
    
    'Calculate percentiles from list and create Histogram chart including cumulative frequencies
    'Two methods to determine the number of bins (Sturges' formula or Square Root Choice)
    
    'Version: 1.0
    
    'Created by Matheus Nunes Reis on 18/12/2024
    'Copyright © 2024 Matheus Nunes Reis. All rights reserved.
    
    'GitHub: MatheusNReis
    'License:
    'MIT License. Copyright © 2024 MatheusNReis
    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    Dim ColunaTempos As String, ColunaBinAuxiliar As String, ColunaTituloPercentis As String
    Dim ColunaOutputRange As String, ColunaPercentis As String
    
    ColunaTempos = "S" '''
    ColunaOutputRange = "Z" '''
    ColunaBinAuxiliar = "Y" ''' Deve ser uma coluna antes da ColunaOutputRange
    ColunaTituloPercentis = "AD" '''
    ColunaPercentis = "AE" '''
    
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, ColunaTempos).End(xlUp).Row
    
    Dim dataRange As Range
    Set dataRange = ws.Range(ColunaTempos & "2:" & ColunaTempos & LastRow).SpecialCells(xlCellTypeConstants)
    
    
    
    ' Percentis
    Dim percentis As Variant
    percentis = Array(0.5, 0.6, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95) '''
    ws.Cells(1, ColunaTituloPercentis) = "Percentil"
    ws.Cells(1, ColunaPercentis) = "Valor"
    
    Dim i As Long
    For i = LBound(percentis) To UBound(percentis)
        ws.Cells(i + 2, ColunaTituloPercentis).Value = percentis(i) * 100 & "%"
        ws.Cells(i + 2, ColunaPercentis).Value = WorksheetFunction.Percentile(dataRange, percentis(i))
    Next i
    
    ws.Range(ColunaPercentis & "2:" & ColunaPercentis & UBound(percentis) + 2).NumberFormat = "hh:mm:ss"
    
    
    
    ' Cria histograma
    'Determine the number of data points
    Dim numDataPoints As Long
    numDataPoints = dataRange.Count
    
    
    Dim numBins As long
    'Calculate the number of bins using Sturges' formula
    'numBins = Application.WorksheetFunction.Ceiling_Math(Application.WorksheetFunction.Log(numDataPoints, 2) + 1)
    
    'Number of bins defined automatically by Square Root Choice
    numBins = Round(Sqr(numDataPoints))
    
    
    'Determine the minimum and maximum values in the data range
    Dim minValue As Double
    Dim maxValue As Double
    minValue = Application.WorksheetFunction.Min(dataRange)
    maxValue = Application.WorksheetFunction.Max(dataRange)
    
    'Calculate bin width
    Dim binWidth As Double
    binWidth = (maxValue - minValue) / numBins
    
    
    ' Define binRange
    'Define first bin value
    ws.Cells(2, ColunaBinAuxiliar).Value = minValue
    'Define the other bin values
    For i = 1 To (numBins)
        ws.Cells(i + 2, ColunaBinAuxiliar).Value = minValue + (binWidth * i)
    Next i
    
    Dim binRange As Range
    Set binRange = ws.Range(ColunaBinAuxiliar & "2:" & ColunaBinAuxiliar & numBins + 1)
    
    'Define outputRange insertion position
    Dim outputRange As Range
    Set outputRange = ws.Range(ColunaOutputRange & "1") ' Starting cell for the output
    
    
    'Clear previous output
    'outputRange.CurrentRegion.Clear
    
    'Create histogram using Analysis ToolPak
    Application.Run "ATPVBAEN.XLAM!Histogram", dataRange, outputRange, binRange, False, True, True, False 'binRange pode ser argumento vazio
    
    
    'Clear origin bins
    ws.Columns(ColunaBinAuxiliar).ClearContents
    
    'Convert chart datarange to Time format
    ws.Range(ColunaOutputRange & "2:" & ColunaOutputRange & LastRow).NumberFormat = "hh:mm:ss"
    
    
    'Add titles to the output
    ws.Range(ColunaOutputRange & "1").Value = "Bloco"
    ws.Cells(1, Range(ColunaOutputRange & "1").Column + 1).Value = "Frequência"
    ws.Cells(1, Range(ColunaOutputRange & "1").Column + 2).Value = "% cumulativo"
    
    
End Sub

