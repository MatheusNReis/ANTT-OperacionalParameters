Sub TratarRecursos_TipoDeVeiculos()

'Atendimento_APH_e_Socorro Mecânico.xlsm

'Processes resources and vehicles types data, relating operacional resources spreadsheet to the spreadsheet
'of services performed, excludes unnecessary/incorrect information, organizes and presents operacional
'checks files in the folder, identifies the correct ones, prepares data source and operational resources spreadsheet
'and prepares a new spreadsheet to receive data processed by method 1 or method 2
'Considers method of single column time determination with recognizable date

'Trata dados de resurso e de tipos de veículos, relacionando planilha de resursos operacionais à planilha de
'atendimentos realizados, exlui informações desnecessárias/incorretas, organiza e apresenta dados operacionais
'conforme regra do primeiro recurso por atendimento
'Verifica arquivos da pasta, identifica os corretos, prepara planilha de origem de dados e de rescursos operacionais
'e prepara nova planilha para receber dados tratados pelo método 1 ou método2
'Considera método da determinação da coluna única de tempo com conversão para data reconhecível

'Version: 2.0

'Created by Matheus Nunes Reis on 18/09/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/05782e3211f77a3b453fbcf02b3e031513031ecf/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


'Armazenamento dos Caminhos dos arquiivos da pasta
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(ThisWorkbook.Sheets("1.Instruções").Cells(1, "B").Value) ''''''''Necessário informar o caminho da pasta que contém planilhas

    Dim file As Object
    Dim i As Long
    i = 1

    ' Cria um array dinâmico chamado "Caminho" para armazenar os caminhos somente de arquivos .xlsx ou .xls existentes na pasta
    Dim Caminho() As String
    Dim Caminho2 As String
   
    ReDim Caminho(1 To 1)

    For Each file In folder.Files
        If ((Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Recursos Operacionais")) <> "Recursos Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Recursos Operacionais")) <> "Recursos Operacionais")) And _
            ((Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Parâmetros Operacionais")) <> "Parâmetros Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Parâmetros Operacionais")) <> "Parâmetros Operacionais")) Then
            ' Redimensiona o array para acomodar mais um caminho de pasta
            ReDim Preserve Caminho(1 To i)
            ' Salva o caminho da planilha no array e o nome do arquivo na variável
            NomeArquivo = file.Name
            Caminho(i) = file.Path
            i = i + 1
        End If
        ' Se o nome do arquivo contém "Recursos Operacionais" seu caminho é guardado na variável Caminho2
        If (Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Recursos Operacionais")) = "Recursos Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Recursos Operacionais")) = "Recursos Operacionais") Then 'InStr(1, fileName, "Recursos Operacionais", vbTextCompare) > 0
            Caminho2 = file.Path 'Caminho do arquivo Recursos Operacionais
        End If
        
    Next file

'''----------

    'Primeiro o usuário confirma se vai iniciar Tratamento de veículos
    Dim resposta As VbMsgBoxResult
    'Dim i As Long
    
    resposta = MsgBox("Tratar dados dos Tipos de Veículos para " & NomeArquivo & "?", vbYesNo + vbQuestion, "Confirmação de Tratamento")
    
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


    Dim ObjWorkbook As Object, ObjWorksheet As Object
    Dim DestWorkbook As Workbook, DestWorksheet As Worksheet
  
    Dim LastRowObjWorksheet As Long
    
    ' Define a primeira planilha da pasta de trabalho de origem
    Set ObjWorkbook = Workbooks.Open(Caminho(1))
    Set ObjWorksheet = ObjWorkbook.Sheets(1)
    
    'Abrir planilha "Recursos Operacionais"
    Dim ROpWorkbook As Object, ROpWorksheet As Object
    Set ROpWorkbook = Workbooks.Open(Caminho2)
    Set ROpWorksheet = ROpWorkbook.Sheets(1)
    
    ' Acrescenta nova planilha na pasta atual de trabalho com nome da concessionária na pasta atual de trabalho recém aberta em objWorkbook
    Dim NomeDaNovaPlanilha As String
    Dim PosicaoInicial As Long
    Dim PosicaoFinal As Long
    PosicaoInicial = InStr(Caminho(1), "- ") + Len("- ")
    PosicaoFinal = InStr(Caminho(1), ".xl") - 1 '.xl seria para identificar .xlsx ou .xls
    NomeDaNovaPlanilha = Mid(Caminho(1), PosicaoInicial, PosicaoFinal - PosicaoInicial + 1)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = NomeDaNovaPlanilha
    End With
    
    ' Define a planilha de destino (Última planilha criada)
    Set DestWorksheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Última linha com dados na planilha de origem
    LastRowObjWorksheet = ObjWorksheet.Cells(ObjWorksheet.Rows.Count, "B").End(xlUp).Row

    Dim lastRowROpWorksheet As Long
    lastRowROpWorksheet = ROpWorksheet.Cells(ROpWorksheet.Rows.Count, "A").End(xlUp).Row


    ' Reorganização da planilha Recursos Operacionais "RoPWorksheet" para que tenha apenas dados da concessionária desejada
    ROpWorksheet.Range("A:L").Sort Key1:=ROpWorksheet.Range("A1"), Order1:=xlAscending, Header:=xlYes
    Dim firstLineNomeConces As Long
    Dim lastLineNomeConces As Long
    firstLineNomeConces = Application.Match(NomeDaNovaPlanilha, ROpWorksheet.Range("A1:A" & lastRowROpWorksheet), 0)
    'lastLineNomeConces = Application.Match(NomeDaNovaPlanilha, ROpWorksheet.Range("A1:A" & lastRowROPWorksheet), 1)
    'Cria lastLineNomeConces
    For i = (firstLineNomeConces + 1) To lastRowROpWorksheet
        If ROpWorksheet.Cells(i, "A") <> NomeDaNovaPlanilha Then
            lastLineNomeConces = i - 1
            Exit For
        End If
    Next i
    'Copiar o intervalo de A:L contendo "NomeDaNovaPlanilha" para a célula A2
    ROpWorksheet.Range("A" & firstLineNomeConces & ":L" & lastLineNomeConces).Cut ROpWorksheet.Range("A2")
    'Encontrar novamente última linha com nome concessionária
    lastLineNomeConces = lastLineNomeConces - firstLineNomeConces + 2 '+2 por conta do cabeçalho e o último dado contado que ficaria de fora
    'Deletar o conteúdo das demais linhas abaixo de A:L
    ROpWorksheet.Range("A" & lastLineNomeConces + 1 & ":L" & lastRowROpWorksheet).ClearContents
    

''''''' Etapa de expurgar dados com tempos <=0, substitiur os recursos pelos tipos de veículos e verificar Inconsistência_serviço_recurso
        'e copia da ObjWorksheet para Destworksheet
        
    ' Copiar linha por linha (linhas não expurgadas)
    'Criação de cabeçalhos (1ª linha)
    DestWorksheet.Range(DestWorksheet.Cells(1, "A"), DestWorksheet.Cells(1, "L")).Value = ObjWorksheet.Range(ObjWorksheet.Cells(1, "A"), ObjWorksheet.Cells(1, "L")).Value
    DestWorksheet.Range("M1").Value = "t. Ocorrência"
    DestWorksheet.Range("N1").Value = "t. Acionamento"
    DestWorksheet.Range("O1").Value = "Mês"
    
    'Converter os dados das colunas J, K e L em data reconhecível
    For Each r In ObjWorksheet.Range("J2:J" & LastRowObjWorksheet)
        r.Value = Format(r.Value, "dd/mm/yyyy hh:mm")
    Next r
    
    For Each r In ObjWorksheet.Range("K2:K" & LastRowObjWorksheet)
        r.Value = Format(r.Value, "dd/mm/yyyy hh:mm")
    Next r
    
    For Each r In ObjWorksheet.Range("L2:L" & LastRowObjWorksheet)
        r.Value = Format(r.Value, "dd/mm/yyyy hh:mm")
    Next r
    
    
    Dim t_ocorrencia As Double, t_acionamento As Double, ValorK_J As Double
    'Conta novamente lastRowROPWorksheet após reorganização da planilha Recursos Operacionais "RoPWorksheet"
    lastRowROpWorksheet = ROpWorksheet.Cells(ROpWorksheet.Rows.Count, "A").End(xlUp).Row 'Necessário na função Compara_Serviço_Recurso()
    Dim k As Long
    k = 2
    Dim QtdeExpurgo As Long, t_ocorrenciaZero As Long, t_acionamentoZero As Long, Chegada_antes_Ocorrencia As Long, Chegada_antes_Acionamento As Long
    QtdeExpurgo = 0
    t_ocorrenciaZero = 0
    t_acionamentoZero = 0
    Chegada_antes_Ocorrencia = 0
    Chegada_antes_Acionamento = 0
    Inconsistência_serviço_recurso = 0
    
    For i = 2 To LastRowObjWorksheet
    
        t_ocorrencia = ObjWorksheet.Cells(i, "L").Value - ObjWorksheet.Cells(i, "J").Value ' coluna L-J
        t_acionamento = ObjWorksheet.Cells(i, "L").Value - ObjWorksheet.Cells(i, "K").Value 'coluna L-K
        ValorK_J = ObjWorksheet.Cells(i, "K").Value - ObjWorksheet.Cells(i, "J").Value ' coluna K-J
        
        If t_ocorrencia >= 0 And t_ocorrencia <= 1.15740767796524E-05 Then
            t_ocorrenciaZero = t_ocorrenciaZero + 1
            QtdeExpurgo = QtdeExpurgo + 1
        ElseIf t_acionamento >= 0 And t_acionamento <= 1.15740767796524E-05 Then
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
        Else
                'Substituir Recursos pelos Tipos de Veículos Veículos 2 e Verifica Inconsistência_serviço_recurso
                'ValorRecurso retorna o código substituído pelo veículo ou mantém o mesmo Recurso caso não seja encontrado código correspondente em ROpWorksheet
                ValorRecurso = Compara_Conces_Recurso_Serviço(ROpWorksheet, lastRowROpWorksheet, ObjWorksheet.Cells(i, "B").Value, _
                                                        ObjWorksheet.Cells(i, "F").Value)
                'Substitui Recurso se encontrado correspondente, caso contrário mantém o anterior
                ObjWorksheet.Cells(i, "F").Value = ValorRecurso
                                                        
                'Até este ponto, todos os códigos/Recuros identificáveis em ROpWorksheet foram substituídos por Guincho Leve/Pesado ou Ambulância C/D
                
                If (ValorRecurso = "Guincho Leve" And ObjWorksheet.Cells(i, "E").Value <> "Mecânico") Or _
                    (ValorRecurso = "Guincho Pesado" And ObjWorksheet.Cells(i, "E").Value <> "Mecânico") Or _
                    (ValorRecurso = "Ambulância C" And ObjWorksheet.Cells(i, "E").Value <> "Médico") Or _
                    (ValorRecurso = "Ambulância D" And ObjWorksheet.Cells(i, "E").Value <> "Médico") Then
                    
                    Inconsistência_serviço_recurso = Inconsistência_serviço_recurso + 1
                    QtdeExpurgo = QtdeExpurgo + 1
                Else
                    'Aplica informações das colunas A:L na planilha destino usando Range
                    DestWorksheet.Range(DestWorksheet.Cells(k, "A"), DestWorksheet.Cells(k, "L")).Value = ObjWorksheet.Range(ObjWorksheet.Cells(i, "A"), ObjWorksheet.Cells(i, "L")).Value
                    'Aplica informações das colunas M, N e O na planilha destino
                    DestWorksheet.Cells(k, "M").Value = t_ocorrencia ''''''
                    DestWorksheet.Cells(k, "N").Value = t_acionamento '''''''
                    DestWorksheet.Cells(k, "O").Value = Format(DestWorksheet.Cells(k, "D").Value, "mm") 'Cria coluna com número do mês. OBS: esta coluna é apagada ao final do processamento
                    k = k + 1
                End If
        End If
    Next i
    
'''''''

    ' Classificar pelas colunas A (Código) e coluna L (Chegada)
    Dim LastRowDestWorksheet As Long
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    With DestWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=DestWorksheet.Range("A1:A" & LastRowDestWorksheet), Order:=xlAscending 'Tanto faz os dados serem texto ou número na coluna A, o importante é os mesmos atendimentos estarem juntos após esta classificação
        .SortFields.Add Key:=DestWorksheet.Range("L1:L" & LastRowDestWorksheet), Order:=xlAscending
        .SetRange DestWorksheet.Range("A1:O" & LastRowDestWorksheet) 'Ordena a tabela da coluna A até O
        .Header = xlYes
        .Apply
    End With
    
    
    ' Determinação da coluna única de tempo -> seleção do tempo entre t. ocorrência ou t. acionamento
    'Adiciona cabeçalho "Tempo" na planilha
    DestWorksheet.Cells(1, "P").Value = "Tempo"
    Dim Codigo As String, Med As Integer, Mec As Integer
    'Inicializa variável
    Codigo = ""
    
    For i = 2 To LastRowDestWorksheet
        
        If DestWorksheet.Cells(i, "A").Value <> Codigo Then
            'Tudo dentro deste If é para quando o código do atendimento acabou de mudar
            'Verifica qual o caso de 1º atendimento
            
            'Reinicializa Med e Mec somente quando o código do atendimento muda
            Mec = 0
            Med = 0
            
            If DestWorksheet.Cells(i, "F").Value = "Guincho Leve" Or DestWorksheet.Cells(i, "F").Value = "Guincho Pesado" Then
                'Se Guincho Leve/Pesado for 1º atendimento
                DestWorksheet.Cells(i, "P").Value = DestWorksheet.Cells(i, "M").Value '''''
                Mec = 1
            ElseIf DestWorksheet.Cells(i, "F").Value = "Ambulância C" Or DestWorksheet.Cells(i, "F").Value = "Ambulância D" Then
                'Se Ambulância for 1º atendimento
                DestWorksheet.Cells(i, "P").Value = DestWorksheet.Cells(i, "M").Value '''''
                Med = 1
            Else
                'Se 1º atendimento não é nem Guincho e nem Ambulância
                'Define Med e Mec iguais a "1" para que posteriormente, no mesmo código de atendimento, t.acionamento seja aplicado para Ambulância e Guincho
                Mec = 1
                Med = 1
            End If
            
        Else
            'Código do atendimento não mudou e o caso de 1º atendimento já foi determinado
            If Mec = 0 And (DestWorksheet.Cells(i, "F").Value = "Guincho Leve" Or DestWorksheet.Cells(i, "F").Value = "Guincho Pesado") Then
                'Se 1º atendimento geral foi médico e ainda não teve o 1º atendimento mecânico
                DestWorksheet.Cells(i, "P").Value = DestWorksheet.Cells(i, "M").Value ''''''
                Mec = 1
            ElseIf Med = 0 And (DestWorksheet.Cells(i, "F").Value = "Ambulância C" Or DestWorksheet.Cells(i, "F").Value = "Ambulância D") Then
                'Se 1º atendimento geral foi mecânico e ainda não teve o 1º atendimento médico
                DestWorksheet.Cells(i, "P").Value = DestWorksheet.Cells(i, "M").Value '''''''
                Med = 1
            Else
                'Se 1º atendimento geral não foi guincho nem ambulância, ou se não for mais o 1º atendimento
                DestWorksheet.Cells(i, "P").Value = DestWorksheet.Cells(i, "N").Value '''''''
            End If
               
        End If
        
        'Atualiza a variável Codigo
        Codigo = DestWorksheet.Cells(i, "A").Value
    Next i
    
    
    
    ' Expurgar veículos que não são Guincho e nem Ambulância
    Dim Fora_atend_mec_med As Long
    k = 2
    For i = 2 To LastRowDestWorksheet
        If DestWorksheet.Cells(i, "F") = "Guincho Leve" Or DestWorksheet.Cells(i, "F") = "Guincho Pesado" Or _
           DestWorksheet.Cells(i, "F") = "Ambulância C" Or DestWorksheet.Cells(i, "F") = "Ambulância D" Then
               
            DestWorksheet.Range(DestWorksheet.Cells(k, "A"), DestWorksheet.Cells(k, "P")) = DestWorksheet.Range(DestWorksheet.Cells(i, "A"), DestWorksheet.Cells(i, "P")).Value
            k = k + 1
        Else
            Nao_Guincho_nem_ambulancia = Nao_Guincho_nem_ambulancia + 1
            QtdeExpurgo = QtdeExpurgo + 1
        End If
    Next i
    'Deleta conteúdo das coluna A até P, das linhas k até DestWorksheet.Rows.Count (linhas desnecessárias que sobraram)
    DestWorksheet.Range("A" & k & ":P" & DestWorksheet.Rows.Count).ClearContents
    
    
    ' Classificar pelas colunas E (Serviço), F (Recurso), O (Mês) e P (Tempo)
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    With DestWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=DestWorksheet.Range("E1:E" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("F1:F" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("O1:O" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("P1:P" & LastRowDestWorksheet), Order:=xlAscending
        .SetRange DestWorksheet.Range("A1:P" & LastRowDestWorksheet) 'Ordena a tabela da coluna A até O
        .Header = xlYes
        .Apply
    End With
    
    
    ' Cria planilha para apresentação de resumo dos expurgos
    Dim DadosExpWorksheet As Worksheet
    Set DadosExpWorksheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    DadosExpWorksheet.Name = "Dados expurgos"
    'Apresentaçõa de resumo dos expurgos
    Percent_Expurgo = (QtdeExpurgo / (LastRowObjWorksheet - 1)) 'Subtrai por 1 por conta do cabeçalho
    DadosExpWorksheet.Range("A1").Value = "Concessionária"
    DadosExpWorksheet.Range("A2").Value = NomeDaNovaPlanilha 'subtrai por 1 desconsiderando cabeçalho
    DadosExpWorksheet.Range("B1").Value = "Nº Atendimentos sem expurgo"
    DadosExpWorksheet.Range("B2").Value = LastRowObjWorksheet - 1 'subtrai por 1 desconsiderando cabeçalho
    DadosExpWorksheet.Range("C1").Value = "Expurgo"
    DadosExpWorksheet.Range("C2").Value = Format(Percent_Expurgo, "Percent")
    DadosExpWorksheet.Range("D1").Value = "t. Ocorrência zero"
    DadosExpWorksheet.Range("D2").Value = t_ocorrenciaZero
    DadosExpWorksheet.Range("E1").Value = "t. Acionamento zero"
    DadosExpWorksheet.Range("E2").Value = t_acionamentoZero
    DadosExpWorksheet.Range("F1").Value = "Chegada antes de ocorrência"
    DadosExpWorksheet.Range("F2").Value = Chegada_antes_Ocorrencia
    DadosExpWorksheet.Range("G1").Value = "Chegada antes de acionamento"
    DadosExpWorksheet.Range("G2").Value = Chegada_antes_Acionamento
    DadosExpWorksheet.Range("H1").Value = "Acionamento antes de ocorrência"
    DadosExpWorksheet.Range("H2").Value = Acionamento_antes_Ocorrencia
    DadosExpWorksheet.Range("I1").Value = "Inconsistência serviço-recurso(Ex: ambulância em serviço mecânico)"
    DadosExpWorksheet.Range("I2").Value = Inconsistência_serviço_recurso
    DadosExpWorksheet.Range("J1").Value = "Não Guincho nem Ambulância"
    DadosExpWorksheet.Range("J2").Value = Nao_Guincho_nem_ambulancia


    'Coloca na planilha "1.Instruções" o nome da concessionária atual tratada
    ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value = NomeDaNovaPlanilha

    ObjWorkbook.Close savechanges:=False
    ROpWorkbook.Close savechanges:=False

    Set DestWorksheet = Nothing
    Set DestWorkbook = Nothing
    Set ObjWorksheet = Nothing
    Set ObjWorkbook = Nothing
    Set ROpWorksheet = Nothing
    Set ROpWorkbook = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    ThisWorkbook.Save
    
    MsgBox "Tratamento de veículos concluído!"


End Sub
