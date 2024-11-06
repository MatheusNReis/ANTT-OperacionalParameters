Sub AtendimentoMetodo1()
   
'Atendimento_APH_e_Socorro Mecânico.xlsm

'Prepares data from the source spreadsheet to calculate pencentiles and averages taking into account the
'1st service for each Resource (vehicle type) in the same code (service)

'Prepara dados da planilha de origem para cálculo de pencentis e média levando em consideração 1º atendimento
'para cada Recurso (tipo de veículo) no mesmo código (atendimento)

'Version: 1.0

'Created by Matheus Nunes Reis on 25/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
   
   
   '''Primeiro o usuário confirma se vai iniciar o processamento
    Dim resposta As VbMsgBoxResult
    Dim i As Long
    
    resposta = MsgBox("Processar dados da planilha " & ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value & "?", vbYesNo + vbQuestion, "Confirmação de Processamento")
    
    If resposta = vbYes Then
    
        Application.DisplayAlerts = False ' Desativa os alertas

        
        'For i = ThisWorkbook.Sheets.Count To 1 Step -1
        '    If (ThisWorkbook.Sheets(i).Name <> "1.Instruções") And (ThisWorkbook.Sheets(i).Name <> "2.Compilado Método1") Then
        '        ThisWorkbook.Sheets(i).Delete
        '    End If
        'Next i
    
        Application.DisplayAlerts = True ' Reativa os alertas
    Else
        Exit Sub
        
    End If
'''----------

'Armazenamento dos Caminhos dos arquivos da pasta
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(ThisWorkbook.Sheets("1.Instruções").Cells(1, "B").Value) ''''''''Necessário informar o caminho da pasta que contém planilhas

    Dim file As Object
    
    i = 1
    
    Dim Caminho3 As String

    For Each file In folder.Files
        If (Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Parâmetros Operacionais")) = "Parâmetros Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Parâmetros Operacionais")) = "Parâmetros Operacionais") Then
            Caminho3 = file.Path 'Caminho do arquivo Parâmetros Operacionais
        End If
    Next file
'''--------------------
    
    
    ' Define a planilha de dados da concessionária já com tipos de veículo tratados
    Dim ConcesWorksheet As Worksheet
    Set ConcesWorksheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value)
    
    'Define a nova planilha DestWorksheet que terá cópia dos dados da ConcesWorksheet
    Dim DestWorksheet As Worksheet
    ' Copia a planilha ConcesWorksheet para DestWorksheet
    ConcesWorksheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ' Renomeia a nova planilha DestWorksheet
    Set DestWorksheet = ActiveSheet
    DestWorksheet.Name = "M1." & ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value
    
    
    ' Copiar dados coluna por coluna da planilha origem para a nova aba recém-criada
    'Dim lastcolumn As Long
    'lastcolumn = objWorksheet.Cells(1, objWorksheet.Columns.Count).End(xlToLeft).Column
    'For i = 1 To lastcolumn
    '    Dim rngOrigem As Range
    '    Dim rngDestino As Range
        ' Definir os intervalos de origem e destino
    '    Set rngOrigem = objWorksheet.Range(objWorksheet.Cells(1, i), objWorksheet.Cells(LastRowObjWorksheet, i))
    '    Set rngDestino = DestWorksheet.Cells(1, i)
        ' Copiar os valores
    '    rngDestino.Resize(rngOrigem.Rows.Count, rngOrigem.Columns.Count).Value = rngOrigem.Value
    'Next i
    
    
    
    ' Classificar pelas colunas E (Serviço), coluna F (Recurso - Tipo de veículo), coluna A (Cod) e coluna D (Atendimento)
    'Última linha com dados na planilha de destino
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    With DestWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=DestWorksheet.Range("E1:E" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("F1:F" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("A1:A" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("D1:D" & LastRowDestWorksheet), Order:=xlAscending
        .SetRange DestWorksheet.Range("A1:N" & LastRowDestWorksheet) 'Ordena a tabela da coluna A até N
        .Header = xlYes
        .Apply
    End With
    
    
    ' Copiar linha por linha (linhas não expurgadas)
    'Continua criação de cabeçalhos (1ª linha)
    DestWorksheet.Range("M1").Value = "t. Ocorrência"
    DestWorksheet.Range("N1").Value = "t. Acionamento"
    DestWorksheet.Range("O1").Value = "Mês"
    'Cópia das linhas com t.Ocorrência e t.Acionamento não nulos na planilha destino (2ª linha em diante) e quantifica expurgos
    Dim t_ocorrencia As Double, t_acionamento As Double
    k = 2
    Dim QtdeExpurgo As Long, t_ocorrenciaZero As Long, t_acionamentoZero As Long, Chegada_antes_Ocorrencia As Long, Chegada_antes_Acionamento As Long
    QtdeExpurgo = 0
    t_ocorrenciaZero = 0
    t_acionamentoZero = 0
    Chegada_antes_Ocorrencia = 0
    Chegada_antes_Acionamento = 0
    ForaDoPrimeiroAtendimento = 0
    ServiçoAnterior = ""
    RecursoAnterior = ""
    CodAnterior = ""
    
    For i = 2 To LastRowDestWorksheet
        t_ocorrencia = DestWorksheet.Cells(i, "L").Value - DestWorksheet.Cells(i, "J").Value ' coluna L-J
        t_acionamento = DestWorksheet.Cells(i, "L").Value - DestWorksheet.Cells(i, "K").Value 'coluna L-K
        ValorK_J = DestWorksheet.Cells(i, "K").Value - DestWorksheet.Cells(i, "J").Value ' coluna K-J
        
        If t_ocorrencia = 0 Then
            t_ocorrenciaZero = t_ocorrenciaZero + 1
            QtdeExpurgo = QtdeExpurgo + 1
        ElseIf t_acionamento = 0 Then
            t_acionamentoZero = t_acionamentoZero + 1
            QtdeExpurgo = QtdeExpurgo + 1
        ElseIf t_ocorrencia < 0 Then
            Chegada_antes_Ocorrencia = Chegada_antes_Ocorrencia + 1
            QtdeExpurgo = QtdeExpurgo + 1
        ElseIf t_acionamento < 0 Then
            Chegada_antes_Acionamento = Chegada_antes_Acionamento + 1
            QtdeExpurgo = QtdeExpurgo + 1
        ElseIf ValorK_J < 0 Then
            Acionamento_antes_Ocorrencia = Acionamento_antes_Ocorrencia + 1
            QtdeExpurgo = QtdeExpurgo + 1
        Else 'Manter somente as linhas do 1º atendimento para cada Recurso (tipo de veículo) no mesmo código
             'Manter 1º atendimento por recurso - Levar em consideração a primeira ocorrencia de cada recurso (tipo de veículo) para um mesmo Cod, a primeira ocorrencia será avaliado pelo Inicio da ocorrencia, e caso ocorrencia seja igual pega só o primeiro
            If (DestWorksheet.Cells(i, "E") = ServiçoAnterior) And (DestWorksheet.Cells(i, "F") = RecursoAnterior) And (DestWorksheet.Cells(i, "A") = CodAnterior) Then
                'Se as variáveis comparadas permanecem iguais, entende-se que não é o primeiro atendimento e a linha é expurgada
                ForaDoPrimeiroAtendimento = ForaDoPrimeiroAtendimento + 1
                QtdeExpurgo = QtdeExpurgo + 1
            Else
                'Teve mudança nas variáveis comparadas, a linha é considerada como 1º atendimento
                'Aplica informações das colunas M e N na planilha destino
                DestWorksheet.Cells(k, "M").Value = Format(t_ocorrencia, "hh:mm:ss")
                DestWorksheet.Cells(k, "N").Value = Format(t_acionamento, "hh:mm:ss")
                DestWorksheet.Cells(k, "O").Value = Format(DestWorksheet.Cells(i, "D").Value, "mm") 'Cria coluna com número do mês. OBS: esta coluna é apagada ao final do processamento
                'Copiar valor de A:L para a planilha de destino usando Range
                DestWorksheet.Range(DestWorksheet.Cells(k, "A"), DestWorksheet.Cells(k, "L")).Value = DestWorksheet.Range(DestWorksheet.Cells(i, "A"), DestWorksheet.Cells(i, "L")).Value 'As linhas são substituídas na mesma planilha, de forma que no após as linha k sobram apenas as linhas desnecessárias, sendo deletadas mais adiante
                k = k + 1
                'Atualização das variáveis devido a mudança em ao menos uma delas
                ServiçoAnterior = DestWorksheet.Cells(i, "E")
                RecursoAnterior = DestWorksheet.Cells(i, "F")
                CodAnterior = DestWorksheet.Cells(i, "A")
            End If
        End If
    Next i
    'Deleta conteúdo das coluna A até O, das linhas k até DestWorksheet.Rows.Count (linhas desnecessárias que sobraram)
    DestWorksheet.Range("A" & k & ":O" & DestWorksheet.Rows.Count).ClearContents
    
    
    ' Apresentação de resumo dos expurgos na planilha destino
    Percent_Expurgo = ((QtdeExpurgo + ConcesWorksheet.Cells(8, "P").Value) / (ConcesWorksheet.Cells(2, "P").Value))
    'P1 e P2 estão no outro módulo
    DestWorksheet.Range("P4").Value = "Expurgo"
    DestWorksheet.Range("P5").Value = Format(Percent_Expurgo, "Percent")
    'P7 e P8 estão no outro módulo
    DestWorksheet.Range("P10").Value = "t. Ocorrência zero"
    DestWorksheet.Range("P11").Value = t_ocorrenciaZero
    DestWorksheet.Range("P13").Value = "t. Acionamento zero"
    DestWorksheet.Range("P14").Value = t_acionamentoZero
    DestWorksheet.Range("P16").Value = "Chegada antes de ocorrência"
    DestWorksheet.Range("P17").Value = Chegada_antes_Ocorrencia
    DestWorksheet.Range("P19").Value = "Chegada antes de acionamento"
    DestWorksheet.Range("P20").Value = Chegada_antes_Acionamento
    DestWorksheet.Range("P22").Value = "Acionamento antes de ocorrência"
    DestWorksheet.Range("P23").Value = Acionamento_antes_Ocorrencia
    DestWorksheet.Range("P25").Value = "Fora do 1º Atendimento"
    DestWorksheet.Range("P26").Value = ForaDoPrimeiroAtendimento
    
    'Copia célula por célula - 'Processo funciona bem, porém mais lentamente
    ' Copia os dados da primeira planilha para a planilha de destino
    'For j = 1 To 12 'Coluna L = 12
    '   For i = 1 To LastRowObjWorksheet ' Copiar a quantidade de linhas da planilha sharepoint
            ' Copia os dados da coluna A da primeira planilha para a coluna A da planilha de destino
    '        DestWorksheet.Cells(i, j).Value = objWorksheet.Cells(i, j).Value
    '    Next i
    'Next j
    
    ' Classificar pelas colunas E (Serviço), coluna F (Recurso - Tipo de veículo), coluna O (Mês)-apagada pós-processamento e M (t. Ocorrência)
    'Última linha com dados na planilha de destino
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    With DestWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=DestWorksheet.Range("E1:E" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("F1:F" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("O1:O" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("M1:M" & LastRowDestWorksheet), Order:=xlAscending
        .SetRange DestWorksheet.Range("A1:O" & LastRowDestWorksheet) 'Ordena a tabela da coluna A até O
        .Header = xlYes
        .Apply
    End With
    
    ' Resultados concessionária
    'Adicionar nova planilha de resultados da concessionária na pasta de trabalho
    Dim PlanR As Worksheet
    Set PlanR = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) 'PlanR nesta etapa do código é a última planilha na contagem
    PlanR.Name = "RM1." & ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value
    'Cria cabeçalhos na planilha de Resultados PlanR
    PlanR.Range("A1").Value = "Código"
    PlanR.Range("B1").Value = "Parâmetro"
    PlanR.Range("C1").Value = "Amostra"
    PlanR.Range("D1").Value = "Ano"
    PlanR.Range("E1").Value = "Mês"
    PlanR.Range("F1").Value = "Mês/Ano"
    PlanR.Range("G1").Value = "Etapa"
    PlanR.Range("H1").Value = "Concessionária"
    PlanR.Range("I1").Value = "BR"
    PlanR.Range("J1").Value = "Ciclo de Inspeção"
    PlanR.Range("K1").Value = "Atendimento"
    PlanR.Range("L1").Value = "Veículo"
    PlanR.Range("M1").Value = "Absoluto/Médio"
    PlanR.Range("N1").Value = "%"
    PlanR.Range("O1").Value = "Tempo(min)"
    PlanR.Range("P1").Value = "Tempo Máx"
    PlanR.Range("Q1").Value = "Resultado"
    PlanR.Range("R1").Value = "Status_inicial"
    PlanR.Range("S1").Value = "Status_final"
    
    'Cálculo dos Percentis e média e preenchimento das
    'Dim LastRowPlanDados
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    Dim ServiçoInicial As String, RecursoInicial As String, MesInicial As String
    Dim LinhaInicial As Long, LinhaFinal As Long, LinhaPlanR As Long
    ServiçoInicial = DestWorksheet.Cells(2, "E").Value
    RecursoInicial = DestWorksheet.Cells(2, "F").Value
    MesInicial = DestWorksheet.Cells(2, "O").Value
    LinhaInicial = 2
    LinhaPlanR = 2
    
    'Abrir planilha "Parâmetros Operacionais"
    Dim POpWorkbook As Object, POpWorksheet As Object
    Set POpWorkbook = Workbooks.Open(Caminho3)
    Set POpWorksheet = POpWorkbook.Sheets(1)
    
    'Preenchimento das colunas da PlanR
    For i = 2 To LastRowDestWorksheet
        If (DestWorksheet.Cells(i, "E").Value <> ServiçoInicial) Or (DestWorksheet.Cells(i, "F").Value <> RecursoInicial) Or (DestWorksheet.Cells(i, "O").Value <> MesInicial) Then
            LinhaFinal = i - 1
           '1-----------------------
            'PlanR coluna B
            If DestWorksheet.Cells(LinhaInicial, "F").Value = "Ambulância C" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 1
            ElseIf DestWorksheet.Cells(LinhaInicial, "F").Value = "Ambulância D" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 2
            ElseIf DestWorksheet.Cells(LinhaInicial, "F").Value = "Guincho Leve" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 3
                 ElseIf DestWorksheet.Cells(LinhaInicial, "F").Value = "Guincho leve" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 3
            ElseIf DestWorksheet.Cells(LinhaInicial, "F").Value = "Guincho Pesado" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 4
            ElseIf DestWorksheet.Cells(LinhaInicial, "F").Value = "Guincho pesado" Then
                PlanR.Cells(LinhaPlanR, "B").Value = 4
            End If
            
            'PlanR coluna C
            'em branco
            'PlanR coluna D
            PlanR.Cells(LinhaPlanR, "D").Value = Format(DestWorksheet.Cells(LinhaInicial, "D").Value, "yyyy")
            'PlanR coluna E
            PlanR.Cells(LinhaPlanR, "E").Value = Format(DestWorksheet.Cells(LinhaInicial, "D").Value, "mm")
            'PlanR coluna F
            PlanR.Cells(LinhaPlanR, "F").Value = PlanR.Cells(LinhaPlanR, "E").Value & "/" & PlanR.Cells(LinhaPlanR, "D").Value
            'PlanR coluna A
            PlanR.Cells(LinhaPlanR, "A").Value = ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value & "_" & PlanR.Cells(LinhaPlanR, "B").Value & "_" & _
                                                    PlanR.Cells(LinhaPlanR, "E").Value & "_" & PlanR.Cells(LinhaPlanR, "D").Value
            'PlanR colunas G, I, J, O ''''M, N outra análise
            Dim rng As Range
            Dim cell As Range
            Dim lastRowPOpWorksheet As Long
            lastRowPOpWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
            ' Procurar valor correspondente na coluna A da planilha POpWorksheet
            Set rng = POpWorksheet.Range("A2:A" & lastRowPOpWorksheet)
            For Each cell In rng 'Percorrendo a planilha POpWorksheet
                If (cell.Offset(0, 1).Value = DestWorksheet.Cells(LinhaInicial, "B")) And (cell.Offset(0, 7).Value = DestWorksheet.Cells(LinhaInicial, "F")) And _
                                                (cell.Offset(0, 8).Value = "Absoluto") Then
                    PlanR.Cells(LinhaPlanR, "G") = cell.Value
                    PlanR.Cells(LinhaPlanR, "I") = cell.Offset(0, 11).Value
                    PlanR.Cells(LinhaPlanR, "J") = cell.Offset(0, 5).Value
                    PlanR.Cells(LinhaPlanR, "O") = cell.Offset(0, 10).Value
                    Exit For
                End If
            Next cell
            'PlanR coluna H
            PlanR.Cells(LinhaPlanR, "H").Value = ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value
            'PlanR coluna K
            PlanR.Cells(LinhaPlanR, "K").Value = DestWorksheet.Cells(LinhaInicial, "E").Value
            'PlanR coluna L
            PlanR.Cells(LinhaPlanR, "L").Value = DestWorksheet.Cells(LinhaInicial, "F").Value
            'PlanR coluna M
            PlanR.Cells(LinhaPlanR, "M").Value = "Médio"
            'PlanR coluna N
            PlanR.Cells(LinhaPlanR, "N").Value = "90%"
            'PlanR coluna P
            PlanR.Cells(LinhaPlanR, "P").Value = Format(TimeSerial(0, PlanR.Cells(LinhaPlanR, "O"), 0), "hh:mm:ss")
            'PlanR coluna Q
            'cálculo média de 90% dos dados iniciais para critérios na coluna "tempos (min)" considerando parâmetros operacionais = "Absoluto"
            TotalLinesBlock = LinhaFinal - LinhaInicial + 1
            x = PlanR.Cells(LinhaPlanR, "N").Value * TotalLinesBlock
            PartialLinesBlock = Application.WorksheetFunction.RoundUp(PlanR.Cells(LinhaPlanR, "N").Value * TotalLinesBlock, 0)
            PlanR.Cells(LinhaPlanR, "Q").Value = Format(Application.WorksheetFunction.Average(DestWorksheet.Range("M" & LinhaInicial & ":M" & (LinhaInicial + PartialLinesBlock))), _
                                                    "hh:mm:ss")
            
           '1-----------------------
           
           
            LinhaPlanR = LinhaPlanR + 1
            '-----------------------
            'Atualiza valor das variáveis devido a mudança de grupo de dados
            ServiçoInicial = DestWorksheet.Cells(i, "E").Value
            RecursoInicial = DestWorksheet.Cells(i, "F").Value
            MesInicial = DestWorksheet.Cells(i, "O").Value
            LinhaInicial = i
        End If
    Next i
    
    
    ' Apaga conteúdo coluna O (mês), usada para classificar
    'DestWorksheet.Range("O1" & ":O" & LastRowDestWorksheet).ClearContents
    
    'Fecha planilha
    POpWorkbook.Close savechanges:=False
    
    ' Fecha e salva a pasta atual de trabalho
    ThisWorkbook.Save

    
    ' Libera a memória
    
    Set DestWorksheet = Nothing
    Set DestWorksheet = Nothing
    Set POpWorksheet = Nothing
    Set POpWorkbook = Nothing
    Set PlanR = Nothing
    
    Set folder = Nothing
    Set fso = Nothing
    
    'ThisWorkbook.Save
    
    MsgBox "Método 1 concluído!"
    
End Sub
