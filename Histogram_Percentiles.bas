Sub Histograma_Percentis()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    Dim ColunaTempos As String, ColunaTitPercentis As String, ColunaPercentis As String
    ColunaTempos = "S" '''
    ColunaTitPercentis = "Y" '''
    ColunaPercentis = "Z" '''
    
    Dim dataRange As Range
    Set dataRange = ws.Range("S2:S129686").SpecialCells(xlCellTypeConstants)
    
    'Percentis
    Dim percentis As Variant
    percentis = Array(0.5, 0.6, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95) '''
    
    Dim i As Long
    For i = LBound(percentis) To UBound(percentis)
        ws.Cells(i + 1, ColunaTitPercentis).Value = "Percentil " & percentis(i) * 100 & "%"
        ws.Cells(i + 1, ColunaPercentis).Value = WorksheetFunction.Percentile(dataRange, percentis(i))
    Next i
    
    
    
    'Cria histograma
   ' Determine the number of data points
    Dim numDataPoints As Long
    numDataPoints = dataRange.Count
    
    ' Calculate the number of bins using Sturges' formula
    Dim numBins As Integer
    'numBins = Application.WorksheetFunction.Ceiling_Math(Application.WorksheetFunction.Log(numDataPoints, 2) + 1)
    
    
    'Number of bins defined manually
    numBins = 50
    
    ' Determine the minimum and maximum values in the data range
    Dim minValue As Double
    Dim maxValue As Double
    minValue = Application.WorksheetFunction.Min(dataRange)
    maxValue = Application.WorksheetFunction.Max(dataRange)
    
    ' Calculate bin width
    Dim binWidth As Double
    binWidth = (maxValue - minValue) / numBins
    
    ' Create bin range values
    Dim binRange As Range
    Set binRange = ws.Range("T1:T" & numBins)
    
    For i = 2 To (numBins)
        ws.Cells(i, "T").Value = minValue + (binWidth * i)
    Next i
    
    Dim outputRange As Range
    Set outputRange = ws.Range("U1") ' Starting cell for the output
    
    ' Clear previous output
    'outputRange.CurrentRegion.Clear
    
    ' Create histogram using Analysis ToolPak
    Application.Run "ATPVBAEN.XLAM!Histogram", dataRange, outputRange, binRange, False, False, True, False
    
    'Clear origin bins
    ws.Columns("T").ClearContents
    
    ' Add titles to the output
    ws.Range("U1").Value = "Bin"
    ws.Range("V1").Value = "Frequency"
    'ws.Range("W1").Value = "Cumulative Frequency"
    
    
End Sub
