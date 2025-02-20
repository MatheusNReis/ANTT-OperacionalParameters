Sub AplicaCódigo_ConverteDataHora()


'Atendimento_APH_e_Socorro Mecânico.xlsm

'In the service control spreadsheet, this code adds and creates the column of services codes and
'convert, when necessary, columns data into valid Date/Time for occurrence, triggering and
'Arrival.

'Na planilha de controle de atendimentos da concessionária, acrescenta a coluna Código e
'converte, quando necessário, dados das colunas em Data/Hora válidas para ocorrência, acionamento e
'chegada.

'Created by Matheus Nunes Reis on 04/12/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    'Armazenamento do Caminho do arquivos da pasta
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(ThisWorkbook.Sheets("1.Instruções").Cells(1, "B").Value) 'Necessário informar o caminho da pasta que contém planilhas

    Dim file As Object
    

    'Cria um array chamado "Caminho" para armazenar o caminho somente de arquivo Controle de Atendimentos .xlsx ou .xls existente na pasta
    Dim Caminho(0) As String
    

    For Each file In folder.Files
    
        If ((Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("CONTROLE DE ATENDIMENTOS")) = "CONTROLE DE ATENDIMENTOS") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("CONTROLE DE ATENDIMENTOS")) = "CONTROLE DE ATENDIMENTOS")) Then
            
            'Salva o caminho da planilha no array e o nome do arquivo na variável
            NomeArquivo = file.Name
            Caminho(0) = file.Path
        End If
       
    Next file



'''----------

    'Primeiro o usuário confirma se vai iniciar Ajuste da planilha Controle de Atendimentos
    Dim resposta As VbMsgBoxResult
    'Dim i As Long
    
    resposta = MsgBox("Ajustar planilha " & NomeArquivo & "?", vbYesNo + vbQuestion, "Confirmação de ajuste")
    
    If resposta = vbYes Then
    
        'Application.DisplayAlerts = False 'Desativa os alertas

        
        'For i = ThisWorkbook.Sheets.Count To 1 Step -1
            'If (ThisWorkbook.Sheets(i).Name <> "1.Instruções") And (ThisWorkbook.Sheets(i).Name <> "2.Compilado Método1") Then
                'ThisWorkbook.Sheets(i).Delete
            'End If
        'Next i
    
        'Application.DisplayAlerts = True 'Reativa os alertas
    Else
        Exit Sub
        
    End If
    
'''----------



    Dim ObjWorkbook As Object, ObjWorksheet As Object
    Dim DestWorkbook As Workbook, DestWorksheet As Worksheet
  
    
    'Define a primeira planilha da pasta de trabalho de origem
    Set ObjWorkbook = Workbooks.Open(Caminho(0))
    Set ObjWorksheet = ObjWorkbook.Sheets(1)

    Dim LastRowObjWorksheet As Long
    LastRowObjWorksheet = ObjWorksheet.Cells(ObjWorksheet.Rows.Count, "A").End(xlUp).Row


    
    Dim i As Long
    Dim Codigo As String, Ocorrencia As String, Atendimento As String
    Dim InicioDaOcorrencia As String, Acionamento As String, Chegada As String
    
    
    'Posição a ser inserida a coluna Código
    Codigo = "A"
    
    'Posição das colunas já considerando a existência da nova coluna Código
    Ocorrencia = "C"
    Atendimento = "D"
    InicioDaOcorrencia = "J"
    Acionamento = "K"
    Chegada = "L"
    
    If ObjWorksheet.Cells(1, Codigo).Value <> "Código" Then
        'Se não existir a coluna Código na posição a ser inserida, a nova coluna Código será criada
    
        'Cria coluna Código
        Range(Codigo & ":" & Codigo).EntireColumn.Insert Shift:=xlToRight
    

        'Preenche coluna Código
        ObjWorksheet.Cells(1, "A").Value = "Código"
        Range("B1").Copy
        Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        
        For i = 2 To LastRowObjWorksheet
            ObjWorksheet.Cells(i, "A").Value = Year(ObjWorksheet.Cells(i, Atendimento).Value) & "-" & Month(ObjWorksheet.Cells(i, Atendimento).Value) & "-" & _
                                                Day(ObjWorksheet.Cells(i, Atendimento).Value) & "-" & ObjWorksheet.Cells(i, Ocorrencia).Value
        Next i
        
    End If



    'Conversão de Data/Hora em número serial
    '---
    
    ' Coluna Início da Ocorrência
    For i = 2 To LastRowObjWorksheet
        
        If IsDate(ObjWorksheet.Cells(i, InicioDaOcorrencia)) Then
            ' Se for data/hora ou string no formato de data/hora reconhecível
            'Extrai parte Data e parte Hora
            ParteData = DateValue(ObjWorksheet.Cells(i, InicioDaOcorrencia).Value)
            ParteHora = TimeValue(ObjWorksheet.Cells(i, InicioDaOcorrencia).Value)
        
            'Combina as partes Data e Hora
            timeSerialValue = DateSerial(Year(ParteData), Month(ParteData), Day(ParteData)) + TimeSerial(Hour(ParteHora), Minute(ParteHora), Second(ParteHora))
            
            ObjWorksheet.Cells(i, InicioDaOcorrencia).Value = timeSerialValue
            
        
        ElseIf ObjWorksheet.Cells(i, InicioDaOcorrencia) = 0 Or _
                Not IsNumeric(ObjWorksheet.Cells(i, InicioDaOcorrencia)) Then 'Esta condição deve estar antes da verificação isnumeric()
            ' Se a célula estiver em branco ou não conter número
            
            MsgBox "Problema em " & InicioDaOcorrencia & i & "." & vbNewLine & "O cursor será movido para a célula."
            Range(InicioDaOcorrencia & i).Select
            Exit Sub
        
        'ElseIf IsNumeric(ObjWorksheet.Cells(i, IniciodaOcorrencia)) Then 'verificação isnumeric
            'GoTo ProximaIteracao
        
        'Else
            'MsgBox "Problema no item da célula " & IniciodaOcorrencia & i & "."
            'MsgBox "Problema em " & IniciodaOcorrencia & i & "." & vbNewLine & "O cursor será movido para a célula."
            'Exit Sub
        
        End If
        
'ProximaIteracao:

    Next i
    
    
    
    ' Coluna Acionamento
    For i = 2 To LastRowObjWorksheet
        
        If IsDate(ObjWorksheet.Cells(i, Acionamento)) Then
            ' Se for data/hora ou string no formato de data/hora reconhecível
            'Extrai parte Data e parte Hora
            ParteData = DateValue(ObjWorksheet.Cells(i, Acionamento).Value)
            ParteHora = TimeValue(ObjWorksheet.Cells(i, Acionamento).Value)
        
            'Combina as partes Data e Hora
            timeSerialValue = DateSerial(Year(ParteData), Month(ParteData), Day(ParteData)) + TimeSerial(Hour(ParteHora), Minute(ParteHora), Second(ParteHora))
            
            ObjWorksheet.Cells(i, Acionamento).Value = timeSerialValue
            
        
        ElseIf ObjWorksheet.Cells(i, Acionamento) = 0 Or _
                Not IsNumeric(ObjWorksheet.Cells(i, Acionamento)) Then 'Esta condição deve estar antes da verificação isnumeric()
            ' Se a célula estiver em branco ou não conter número
            
            MsgBox "Problema em " & Acionamento & i & "." & vbNewLine & "O cursor será movido para a célula."
            Range(Acionamento & i).Select
            Exit Sub
        
        End If

    Next i
    
    
    
    ' Coluna Chegada
    For i = 2 To LastRowObjWorksheet
        
        If IsDate(ObjWorksheet.Cells(i, Chegada)) Then
            ' Se for data/hora ou string no formato de data/hora reconhecível
            'Extrai parte Data e parte Hora
            ParteData = DateValue(ObjWorksheet.Cells(i, Chegada).Value)
            ParteHora = TimeValue(ObjWorksheet.Cells(i, Chegada).Value)
        
            'Combina as partes Data e Hora
            timeSerialValue = DateSerial(Year(ParteData), Month(ParteData), Day(ParteData)) + TimeSerial(Hour(ParteHora), Minute(ParteHora), Second(ParteHora))
            
            ObjWorksheet.Cells(i, Chegada).Value = timeSerialValue
            
        
        ElseIf ObjWorksheet.Cells(i, Chegada) = 0 Or _
                Not IsNumeric(ObjWorksheet.Cells(i, Chegada)) Then 'Esta condição deve estar antes da verificação isnumeric()
            ' Se a célula estiver em branco ou não conter número
            
            MsgBox "Problema em " & Chegada & i & "." & vbNewLine & "O cursor será movido para a célula."
            Range(Chegada & i).Select
            Exit Sub
        
        End If

    Next i
    
    '---
    
    
    MsgBox "Fim do Processo"


End Sub
