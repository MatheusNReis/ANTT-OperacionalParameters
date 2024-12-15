Sub TratarRecursos_TipoDeVeiculos()


'Atendimento_APH_e_Socorro Mecânico.xlsm

'Processes resources and vehicles types data, relating operacional resources spreadsheet to the spreadsheet
'of services performed, excludes unnecessary/incorrect information, organizes and presents operacional
'data according to the rule of first resource per service
'All kinds of purges are done after verification of first resource in each service
'Minimum time is not considered in purge processing
'Classification for counting of services according to highway and concession
'Includes latitude and longitude data

'Trata dados de resurso e de tipos de veículos, relacionando planilha de resursos operacionais à planilha de
'atendimentos realizados, exlui informações desnecessárias/incorretas, organiza e apresenta dados operacionais
'conforme regra do primeiro recurso por atendimento
'Todos os tipos de expurgos são feitos após verificação do primeiro recurso por atendimento
'Tempo mínimo não é considerado no processamento de expurgo
'Classificação para contagem dos atendimentos conforme rodovia e concessão
'Contempla dados de Latitude e Longitude

'Version: 5.0

'Created by Matheus Nunes Reis on 04/12/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
'MIT License. Copyright © 2024 MatheusNReis



'Armazenamento dos Caminhos dos arquivos da pasta
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
    
    
    
    ' Classificar pelas colunas A (Código) e coluna L (Chegada) - Não necessita classificar por recurso ou Serviço
    
    With ObjWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ObjWorksheet.Range("A1:A" & LastRowObjWorksheet), Order:=xlAscending 'Tanto faz os dados serem texto ou número na coluna A, o importante é os mesmos atendimentos estarem juntos após esta classificação
        .SortFields.Add Key:=ObjWorksheet.Range("L1:L" & LastRowObjWorksheet), Order:=xlAscending
        .SetRange ObjWorksheet.Range("A1:Z" & LastRowObjWorksheet) 'Ordena a tabela da coluna A até Z
        .Header = xlYes
        .Apply
    End With
    
    
    

''''''' Etapa de, linha por linha, realizar análise do 1º atendimento, expurgar dados com tempos <=0, substitiur os recursos pelos tipos de veículos, verificar nem guincho nem ambulância,
        'verificar Inconsistência_serviço_recurso
        'e copiar da ObjWorksheet para Destworksheet se não for expurgo
        
    ' Copiar linha por linha (linhas não expurgadas)
    'Criação de cabeçalhos (1ª linha)
    DestWorksheet.Range(DestWorksheet.Cells(1, "A"), DestWorksheet.Cells(1, "Q")).Value = ObjWorksheet.Range(ObjWorksheet.Cells(1, "A"), ObjWorksheet.Cells(1, "Q")).Value
    DestWorksheet.Range("R1").Value = "t. Ocorrência"
    DestWorksheet.Range("S1").Value = "t. Acionamento"
    DestWorksheet.Range("T1").Value = "Mês"
    DestWorksheet.Range("U1").Value = "Tempo"
    DestWorksheet.Range("V1").Value = "Ano"
    DestWorksheet.Range("W1").Value = "Rodovia Concat"
    DestWorksheet.Range("X1").Value = "Contagem"
    
    
    'Conta novamente lastRowROPWorksheet após reorganização da planilha Recursos Operacionais "RoPWorksheet"
    lastRowROpWorksheet = ROpWorksheet.Cells(ROpWorksheet.Rows.Count, "A").End(xlUp).Row 'Necessário na função Compara_Serviço_Recurso()
    Dim ValorRecurso As String
    Dim CódigoLinhaAnterior As String
    CódigoLinhaAnterior = ObjWorksheet.Cells(2, "A").Value
    Dim PrimeiroAtend As Long, AmbC As Long, AmbD As Long, GL As Long, GP As Long
    PrimeiroAtend = 0   'O atendimento atual é o 1º atendimento
    AmbC = 0    'Nenhum atendimento de Ambulância C realizado ainda
    AmbD = 0    'Nenhum atendimento de Ambulância D realizado ainda
    GL = 0      'Nenhum atendimento de Guincho Leve realizado ainda
    GP = 0      'Nenhum atendimento de Guincho Pesado realizado ainda
    Dim t_ocorrencia As Double, t_acionamento As Double, ValorK_J As Double
    Dim Base As String
    Dim QtdeExpurgo As Long, t_ocorrenciaZero As Long, t_acionamentoZero As Long, Chegada_antes_Ocorrencia As Long, Chegada_antes_Acionamento As Long
    Dim Acionamento_antes_Ocorrencia As Long
    QtdeExpurgo = 0
    t_ocorrenciaZero = 0
    t_acionamentoZero = 0
    Chegada_antes_Ocorrencia = 0
    Chegada_antes_Acionamento = 0
    Acionamento_antes_Ocorrencia = 0
    Inconsistência_serviço_recurso = 0
    AmbC_Repetida = 0
    AmbD_Repetida = 0
    GL_Repetido = 0
    GP_Repetido = 0
    Dim LinhaDestworksheet As Long 'Linha na DestWorksheet
    LinhaDestworksheet = 2
    
    For i = 2 To LastRowObjWorksheet 'i é linha na ObjWorksheet
    
        'Substituir Recurso da linha atual pelos Tipos de Veículos 2 da RopWorksheet e Verifica Inconsistência_serviço_recurso
        'Necessário já no início do loop pois tem que saber qual o tipo de veículo nas verificações seguintes
        'ValorRecurso retorna o código substituído pelo tipo de veículo ou mantém o mesmo Recurso caso não seja encontrado código correspondente em ROpWorksheet
        ValorRecurso = Compara_Conces_Recurso_Serviço(ROpWorksheet, lastRowROpWorksheet, ObjWorksheet.Cells(i, "B").Value, _
                                                        ObjWorksheet.Cells(i, "F").Value)
        'Substitui Recurso em ObjWorksheet se encontrado correspondente, caso contrário mantém o anterior
        ObjWorksheet.Cells(i, "F").Value = ValorRecurso
                                                        
        'Até este ponto, o tipo de veículo/Recurso identificável em ROpWorksheet foi substituído por Guincho Leve/Pesado ou Ambulância C/D na ObjWorksheet


        'Verificação se é mesmo atendimento (Código) da linha anterior
        If ObjWorksheet.Cells(i, "A").Value <> CódigoLinhaAnterior Then
            'O atendimento atual é o 1º, e recomeça a contagem de ambulância e guincho
            PrimeiroAtend = 0
            AmbC = 0
            AmbD = 0
            GL = 0
            GP = 0
        End If
        
        
        
        'Verifica se o veículo é AmbC, AmbD, GL ou GP e adiciona+1 à quantidade se algum for verdadeiro
        
        If ValorRecurso = "Ambulância C" Then
            AmbC = AmbC + 1
        ElseIf ValorRecurso = "Ambulância D" Then
            AmbD = AmbD + 1
        ElseIf ValorRecurso = "Guincho Leve" Then
            GL = GL + 1
        ElseIf ValorRecurso = "Guincho Pesado" Then
            GP = GP + 1
        End If
        
        
        'Cálculo dos Tempos
        t_ocorrencia = ObjWorksheet.Cells(i, "L").Value - ObjWorksheet.Cells(i, "J").Value ' coluna L-J (Tempo Ocorrência)
        t_acionamento = ObjWorksheet.Cells(i, "L").Value - ObjWorksheet.Cells(i, "K").Value 'coluna L-K (Tempo Acionamento)
        ValorK_J = ObjWorksheet.Cells(i, "K").Value - ObjWorksheet.Cells(i, "J").Value ' coluna K-J (Hora Acionamento - Hora ocorrência)
        
        
        'Regra 1º Atendimento: TempoOcorrência para 1º atendimento geral, demais são TempoAcionamento. (Qualquer tipo de veículo pode ser 1º atendimento)
        If PrimeiroAtend = 0 Then
            Tempo = t_ocorrencia
        Else 'PrimeiroAtend = 1
            Tempo = t_acionamento
        End If
        
        
        Base = ObjWorksheet.Cells(i, "M").Value 'Base é a coluna Classificação Tempos Zeros
        
        
        '   Verificação Expurgos, e se não expurgada a linha, os dados são aplicados na DestWorksheet
        'Excluir tempos negativos (Chegada antes de ocorrência, Chegada antes de acionamento e Acionamento antes de ocorrência).
        If t_ocorrencia < 0 Then
            Chegada_antes_Ocorrencia = Chegada_antes_Ocorrencia + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        ElseIf t_acionamento < 0 Then
            Chegada_antes_Acionamento = Chegada_antes_Acionamento + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        ElseIf ValorK_J < 0 Then
            Acionamento_antes_Ocorrencia = Acionamento_antes_Ocorrencia + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        'Tempos zeros: expurgados SOMENTE tempos zeros na base (t_ocorrencia=0 ou t_acionamento=0) e base = 2 (atendimento na base operacional)
        ElseIf PrimeiroAtend = 0 And t_ocorrencia = 0 And (Base = "2" Or Base = "") Then 'PrimeiroAtend = 0 relacion-ase ao t_ocorrencia
            t_ocorrenciaZero = t_ocorrenciaZero + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
        ElseIf PrimeiroAtend = 1 And t_acionamento = 0 And (Base = "2" Or Base = "") Then 'PrimeiroAtend = 1 relaciona-se ao t_acionamento
            t_acionamentoZero = t_acionamentoZero + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        'Tempo > 0 e <= 3 minutos
        'ElseIf Tempo > 0 And Tempo <= 0.00208333333333 Then
            'TempoReduzido = TempoReduzido + 1 'TempoReduzido é a contagem de Tempos >0 e <= 3 minutos
            'QtdeExpurgo = QtdeExpurgo + 1
            'PrimeiroAtend = 1
            
        'Nem guincho e nem ambulância
        ElseIf ValorRecurso <> "Ambulância C" And _
               ValorRecurso <> "Ambulância D" And _
               ValorRecurso <> "Guincho Leve" And _
               ValorRecurso <> "Guincho Pesado" Then
               
            Nao_Guincho_nem_ambulancia = Nao_Guincho_nem_ambulancia + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        'Incosistência Serviço/Recurso no atendimento médico/mecânico. 'tem que ser após verificação Nem Guincho Nem Ambulância
        ElseIf (ValorRecurso = "Guincho Leve" And ObjWorksheet.Cells(i, "E").Value <> "Mecânico") Or _
                (ValorRecurso = "Guincho Pesado" And ObjWorksheet.Cells(i, "E").Value <> "Mecânico") Or _
                (ValorRecurso = "Ambulância C" And ObjWorksheet.Cells(i, "E").Value <> "Médico") Or _
                (ValorRecurso = "Ambulância D" And ObjWorksheet.Cells(i, "E").Value <> "Médico") Then
                    
            Inconsistência_serviço_recurso = Inconsistência_serviço_recurso + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
        'Recurso Médico/Mecânico repetido (A linha é expurgada se tiver mais de um AmbC, AmbD, GL ou GP no mesmo atendimento)
        ElseIf (ValorRecurso = "Ambulância C" And AmbC > 1) Then
            AmbC_Repetida = AmbC_Repetida + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
        ElseIf (ValorRecurso = "Ambulância D" And AmbD > 1) Then
            AmbD_Repetida = AmbD_Repetida + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
        ElseIf (ValorRecurso = "Guincho Leve" And GL > 1) Then
            GL_Repetido = GL_Repetido + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
        ElseIf (ValorRecurso = "Guincho Pesado" And GP > 1) Then
            GP_Repetido = GP_Repetido + 1
            QtdeExpurgo = QtdeExpurgo + 1
            PrimeiroAtend = 1
            
        Else
            'Linha não expurgada - atribui dados da linha na planilha destino Destworksheet
            
            'Aplica informações das colunas A:Q na planilha destino usando Range
            DestWorksheet.Range(DestWorksheet.Cells(LinhaDestworksheet, "A"), DestWorksheet.Cells(LinhaDestworksheet, "Q")).Value = ObjWorksheet.Range(ObjWorksheet.Cells(i, "A"), ObjWorksheet.Cells(i, "Q")).Value
            
            'Aplica informações das colunas R, S, T, U e V na planilha destino
            DestWorksheet.Cells(LinhaDestworksheet, "R").Value = t_ocorrencia
            DestWorksheet.Cells(LinhaDestworksheet, "S").Value = t_acionamento
            DestWorksheet.Cells(LinhaDestworksheet, "T").Value = Format(DestWorksheet.Cells(LinhaDestworksheet, "D").Value, "mm")
            DestWorksheet.Cells(LinhaDestworksheet, "U").Value = Tempo
            DestWorksheet.Cells(LinhaDestworksheet, "V").Value = Format(DestWorksheet.Cells(LinhaDestworksheet, "D").Value, "yyyy")
            
            
            
            'Preenchimento da coluna 'Rodovia Concat'
            '--
            
            'ECO RIOMINAS
            If ObjWorksheet.Cells(i, "B").Value = "ECORIOMINAS" And ObjWorksheet.Cells(i, "H").Value = "BR-465" And ValorRecurso = "Ambulância C" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-465"
            ElseIf ObjWorksheet.Cells(i, "B").Value = "ECORIOMINAS" And ObjWorksheet.Cells(i, "H").Value <> "BR-465" And ValorRecurso = "Ambulância C" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-116/BR-493"
            ElseIf ObjWorksheet.Cells(i, "B").Value = "ECORIOMINAS" And ObjWorksheet.Cells(i, "H").Value = "BR-116" And (ValorRecurso = "Ambulância D" Or ValorRecurso = "Guincho Leve") Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-116"
            ElseIf ObjWorksheet.Cells(i, "B").Value = "ECORIOMINAS" And ObjWorksheet.Cells(i, "H").Value <> "BR-116" And (ValorRecurso = "Ambulância D" Or ValorRecurso = "Guincho Leve") Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-465/BR-493"
            ElseIf ObjWorksheet.Cells(i, "B").Value = "ECORIOMINAS" And (ValorRecurso = "Guincho Pesado") Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "Todas"
            
            'ECOVIAS DO ARAGUAIA
            ElseIf ObjWorksheet.Cells(i, "B").Value = "ECOVIAS DO ARAGUAIA" And ObjWorksheet.Cells(i, "H").Value = "BR-153" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-153"
                ElseIf ObjWorksheet.Cells(i, "B").Value = "ECOVIAS DO ARAGUAIA" And ObjWorksheet.Cells(i, "H").Value <> "BR-153" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "BR-080/BR-414"
                
            'RIO SP
            ElseIf ObjWorksheet.Cells(i, "B").Value = "RIOSP" And ValorRecurso = "Ambulância C" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = "Todas"
            ElseIf ObjWorksheet.Cells(i, "B").Value = "RIOSP" And ValorRecurso <> "Ambulância C" Then
                DestWorksheet.Cells(LinhaDestworksheet, "W").Value = ObjWorksheet.Cells(i, "H").Value 'Rodovia Atual
            
            End If
            '--
            
            
            
            LinhaDestworksheet = LinhaDestworksheet + 1
            PrimeiroAtend = 1
            
        End If
    
    
        CódigoLinhaAnterior = ObjWorksheet.Cells(i, "A").Value
    
    Next i
    
 '''''''
    
    
    
    
    ' Classificar pelas colunas F(Recurso), V(Ano), T(Mês), W(Rodovia Concat) e U(Tempo)
    Dim LastRowDestWorksheet As Long
    LastRowDestWorksheet = DestWorksheet.Cells(DestWorksheet.Rows.Count, "B").End(xlUp).Row
    With DestWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=DestWorksheet.Range("F1:F" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("V1:V" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("T1:T" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("W1:W" & LastRowDestWorksheet), Order:=xlAscending
        .SortFields.Add Key:=DestWorksheet.Range("U1:U" & LastRowDestWorksheet), Order:=xlAscending
        .SetRange DestWorksheet.Range("A1:Z" & LastRowDestWorksheet) 'Ordena a tabela da coluna A até Z
        .Header = xlYes
        .Apply
    End With
    
    
    
    'Realizar contagem (preencher coluna Contagem)
    'Inicialização
    Resource = DestWorksheet.Cells(2, "F").Value
    ano = DestWorksheet.Cells(2, "V").Value
    mes = DestWorksheet.Cells(2, "T").Value
    Rodovia_Concat = DestWorksheet.Cells(2, "W").Value
    Count = 0
    DestWorksheet.Cells(2, "X").Value = 1
    
    Dim j As Long
    For j = 2 To LastRowDestWorksheet
        If DestWorksheet.Cells(j, "F").Value = Resource And DestWorksheet.Cells(j, "V").Value = ano And DestWorksheet.Cells(j, "T").Value = mes And _
            DestWorksheet.Cells(j, "W").Value = Rodovia_Concat Then
            'O registro ainda pertence ao mesmo grupo
            Count = Count + 1
            DestWorksheet.Cells(j, "X").Value = Count
        Else
            'O registro pertence a um novo grupo
            Count = 1 'Reinicializa Count
            DestWorksheet.Cells(j, "X").Value = 1
            'Reinicializa variáveis
            Resource = DestWorksheet.Cells(j, "F").Value
            ano = DestWorksheet.Cells(j, "V").Value
            mes = DestWorksheet.Cells(j, "T").Value
            Rodovia_Concat = DestWorksheet.Cells(j, "W").Value
        End If
    Next j
    
    
    
    ' Cria planilha para apresentação de resumo dos expurgos
    Dim DadosExpWorksheet As Worksheet
    Set DadosExpWorksheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    DadosExpWorksheet.Name = "Dados expurgos"
    'Apresentação de resumo dos expurgos
    Percent_Expurgo = (QtdeExpurgo / (LastRowObjWorksheet - 1)) 'Subtrai por 1 por conta do cabeçalho
    DadosExpWorksheet.Range("A1").Value = "Concessionária"
    DadosExpWorksheet.Range("A2").Value = NomeDaNovaPlanilha
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
    DadosExpWorksheet.Range("K1").Value = "AC rep."
    DadosExpWorksheet.Range("K2").Value = AmbC_Repetida
    DadosExpWorksheet.Range("L1").Value = "AD rep."
    DadosExpWorksheet.Range("L2").Value = AmbD_Repetida
    DadosExpWorksheet.Range("M1").Value = "GL rep."
    DadosExpWorksheet.Range("M2").Value = GL_Repetido
    DadosExpWorksheet.Range("N1").Value = "GP rep."
    DadosExpWorksheet.Range("N2").Value = GP_Repetido
    'DadosExpWorksheet.Range("O1").Value = "Tempo Reduzido"
    'DadosExpWorksheet.Range("O2").Value = TempoReduzido


    'Coloca na planilha "1.Instruções" o nome da concessionária atual tratada
    'ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value = NomeDaNovaPlanilha

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
