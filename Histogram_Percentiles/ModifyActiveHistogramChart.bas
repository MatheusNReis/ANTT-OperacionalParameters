Sub ModifyActiveHistogramChart()


    'Configura o histograma ativo na planilha
    'Mostra dados no gráfico referentes à linha de dado definida
    
    'Set the active histogram chart in the worksheet
    'Shows data in chart refering to the last data line defined
    
    'Created by Matheus Nunes Reis on 18/12/2024
    'Copyright © 2024 Matheus Nunes Reis. All rights reserved.
    
    'GitHub: MatheusNReis
    'License:
    'MIT License. Copyright © 2024 MatheusNReis
    

    Dim chart As chart
    Dim newRange As Range
    Dim ws As Worksheet
    
    ChartRangeFinalLine = 47 '''
    ColunaOutputRange = "Z" '''
    
    Set chart = ActiveChart
    Set ws = ThisWorkbook.Sheets(1)


    'Change x-axis title to "Tempo"
    
    chart.Axes(xlCategory, xlPrimary).HasTitle = True
    chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Tempo"

    'Change chart data range to especified range
    'Define separately data ranges for x and y axis
    Set xRange = ws.Range(ColunaOutputRange & "2:" & ColunaOutputRange & ChartRangeFinalLine)
    
    Set yRange = ws.Range(ws.Cells(2, ws.Cells(1, ColunaOutputRange).Column + 1), _
                            ws.Cells(ChartRangeFinalLine, ws.Cells(1, ColunaOutputRange).Column + 1))
 
    Set wRange = ws.Range(ws.Cells(2, ws.Cells(1, ColunaOutputRange).Column + 2), _
                            ws.Cells(ChartRangeFinalLine, ws.Cells(1, ColunaOutputRange).Column + 2))

    chart.SeriesCollection(1).XValues = xRange
    chart.SeriesCollection(1).Values = yRange
    
    chart.SeriesCollection(2).XValues = xRange
    chart.SeriesCollection(2).Values = wRange
    
    'Set chart legends and linegrids
    chart.Legend.Position = xlLegendPositionBottom
    
    With chart.Axes(xlValue, xlSecondary)
        .HasMajorGridlines = True
    End With
    
    'Set chart dimensions
    chart.Parent.Width = 680
    chart.Parent.Height = 255
    
End Sub
