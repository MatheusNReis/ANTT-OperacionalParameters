Sub ModifyHistogramCharts()

    
    'Set the active histogram chart in the worksheet.
    'Shows data in chart refering to the last data line defined.
    'Last chart line defined automatically.
    'Charts set automatically from call of the main sub.
    
    'Configura o histograma ativo na planilha.
    'Mostra dados no gráfico referentes à linha de dado definida.
    'Última linha de dado para o gráfico definida automaticamente.
    'Gráficos configurados automaticamente a partir de chamada da rotina principal.
    
    'version: 2.0
    
    'Created by Matheus Nunes Reis on 27/12/2024
    'Copyright © 2024 Matheus Nunes Reis. All rights reserved.
    
    'GitHub: MatheusNReis
    'License:
    'MIT License. Copyright © 2024 MatheusNReis
    

    Dim chart As chart
    Dim newRange As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    
    For i = 1 To 4
    
        Set chart = ws.ChartObjects(i).chart
        
        Dim ResourceTitle(1 To 4) As String
        ResourceTitle(1) = "Ambulância C"
        ResourceTitle(2) = "Ambulância D"
        ResourceTitle(3) = "Guincho Leve"
        ResourceTitle(4) = "Guincho Pesado"
        
        Dim ColunaOutputRange(1 To 4) As String
        ColunaOutputRange(1) = "AB" '''
        ColunaOutputRange(2) = "AP" '''
        ColunaOutputRange(3) = "BD" '''
        ColunaOutputRange(4) = "BR" '''
        
        
        LastRow = ws.Cells(ws.Rows.Count, ColunaOutputRange(i)).End(xlUp).Row
        Dim a As Long
        For a = 2 To LastRow
            If ws.Cells(a, ColunaOutputRange(i)).Value >= 0.0625 Then ''' 00:01:30
                Exit For
            End If
        Next a
        
        ChartRangeFinalLine = a
        
        
        'Set chart = ActiveChart
        
    
        'Change x-axis title to "Tempo"
        
        chart.Axes(xlCategory, xlPrimary).HasTitle = True
        chart.ChartTitle.Text = "HISTOGRAMA " & ws.Name & " - " & ResourceTitle(i)
        chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Tempo"
    
        'Change chart data range to especified range
        'Define separately data ranges for x and y axis
        Set xRange = ws.Range(ColunaOutputRange(i) & "2:" & ColunaOutputRange(i) & ChartRangeFinalLine)
        
        Set yRange = ws.Range(ws.Cells(2, ws.Cells(1, ColunaOutputRange(i)).Column + 1), _
                                ws.Cells(ChartRangeFinalLine, ws.Cells(1, ColunaOutputRange(i)).Column + 1))
     
        Set wRange = ws.Range(ws.Cells(2, ws.Cells(1, ColunaOutputRange(i)).Column + 2), _
                                ws.Cells(ChartRangeFinalLine, ws.Cells(1, ColunaOutputRange(i)).Column + 2))
    
        chart.SeriesCollection(1).XValues = xRange
        chart.SeriesCollection(1).Values = yRange
        
        chart.SeriesCollection(2).XValues = xRange
        chart.SeriesCollection(2).Values = wRange
        
        'Set chart legends, linegrids and maximum value of the secondary vertical axis
        chart.Legend.Position = xlLegendPositionBottom
        
        With chart.Axes(xlValue, xlSecondary)
            .MaximumScale = 1
            .HasMajorGridlines = True
            .TickLabels.NumberFormat = "0%"
        End With
        
        
        'Set chart dimensions
        chart.Parent.Width = 680
        chart.Parent.Height = 255
    
       
    Next i
    
    
End Sub
