Sub TratarRecursos_TipoDeVeiculos()

'Atendimento_APH_e_Socorro Mecânico.xlsm

'Processes resources and vehicles types data, relating operacional resources spreadsheet to the spreadsheet
'of services performed, excludes unnecessary/incorrect information, organizes and presents operacional
'checks files in the folder, identifies the correct ones, prepares data source and operational resources spreadsheet
'and prepares a new spreadsheet to receive data processed by method 1 or method 2

'Trata dados de resurso e de tipos de veículos, relacionando planilha de resursos operacionais à planilha de
'atendimentos realizados, exlui informações desnecessárias/incorretas, organiza e apresenta dados operacionais
'conforme regra do primeiro recurso por atendimento
'Verifica arquivos da pasta, identifica os corretos, prepara planilha de origem de dados e de rescursos operacionais
'e prepara nova planilha para receber dados tratados pelo método 1 ou método2

'Version: 1.0

'Created by Matheus Nunes Reis on 15/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
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


    Dim objWorkbook As Object, objWorksheet As Object
    Dim DestWorkbook As Workbook, DestWorksheet As Worksheet
  
    Dim LastRowObjWorksheet As Long
    
    ' Define a primeira planilha da pasta de trabalho de origem
    Set objWorkbook = Workbooks.Open(Caminho(1))
    Set objWorksheet = objWorkbook.Sheets(1)
    
    'Abrir planilha "Recursos Operacionais"
    Dim ROpWorkbook As Object, ROpWorksheet As Object
    Set ROpWorkbook = Workbooks.Open(Caminho2)
    Set ROpWorksheet = ROpWorkbook.Sheets(1)
    
    ' Acrescenta nova planilha na pasta atual de trabalho com nome da concessionária na pasta atual de trabalho recém aberta em objWorkbook
    Dim NomeDaNovaPlanilha As String
    Dim PosicaoInicial As Integer
    Dim PosicaoFinal As Integer
    PosicaoInicial = InStr(Caminho(1), "- ") + Len("- ")
    PosicaoFinal = InStr(Caminho(1), ".xl") - 1 '.xl seria para identificar .xlsx ou .xls
    NomeDaNovaPlanilha = Mid(Caminho(1), PosicaoInicial, PosicaoFinal - PosicaoInicial + 1)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = NomeDaNovaPlanilha
    End With
    
    ' Define a planilha de destino (Última planilha criada)
    Set DestWorksheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Última linha com dados na planilha de origem
    LastRowObjWorksheet = objWorksheet.Cells(objWorksheet.Rows.Count, "B").End(xlUp).Row

    Dim lastRowROPWorksheet As Long
    lastRowROPWorksheet = ROpWorksheet.Cells(ROpWorksheet.Rows.Count, "A").End(xlUp).Row


    ' Reorganização da planilha Recursos Operacionais "RoPWorksheet" para que tenha apenas dados da concessionária desejada
    ROpWorksheet.Range("A:L").Sort Key1:=ROpWorksheet.Range("A1"), Order1:=xlAscending, Header:=xlYes
    Dim firstLineNomeConces As Long
    Dim lastLineNomeConces As Long
    firstLineNomeConces = Application.Match(NomeDaNovaPlanilha, ROpWorksheet.Range("A1:A" & lastRowROPWorksheet), 0)
    'lastLineNomeConces = Application.Match(NomeDaNovaPlanilha, ROpWorksheet.Range("A1:A" & lastRowROPWorksheet), 1)
    'Cria lastLineNomeConces
    For i = (firstLineNomeConces + 1) To lastRowROPWorksheet
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
    ROpWorksheet.Range("A" & lastLineNomeConces + 1 & ":L" & lastRowROPWorksheet).ClearContents
    

    ' Substituir os Recuros (AC01, AD02, GL15...) pelos tipo de veículo correspondente (Ambulância C, Guincho Leve...) da planilha Recursos Operacionais
    'Criação de cabeçalhos (1ª linha)
    DestWorksheet.Range(DestWorksheet.Cells(1, "A"), DestWorksheet.Cells(1, "L")).Value = objWorksheet.Range(objWorksheet.Cells(1, "A"), objWorksheet.Cells(1, "L")).Value

    'Dim lastRowROPWorksheet As Long 'Necessário na função Compara_Serviço_Recurso()
    lastRowROPWorksheet = ROpWorksheet.Cells(ROpWorksheet.Rows.Count, "A").End(xlUp).Row 'Necessário na função Compara_Serviço_Recurso()
    Dim k As Long
    k = 2
    Inconsistência_serviço_recurso = 0
    For i = 2 To LastRowObjWorksheet
        ValorRecurso = Compara_Conces_Recurso_Serviço(ROpWorksheet, lastRowROPWorksheet, objWorksheet.Cells(i, "B").Value, _
                                                objWorksheet.Cells(i, "F").Value, objWorksheet.Cells(i, "E").Value)
        If ValorRecurso <> "" Then
            'Copiar valor de A:L para a planilha de destino usando Range
            DestWorksheet.Range(DestWorksheet.Cells(k, "A"), DestWorksheet.Cells(k, "L")).Value = objWorksheet.Range(objWorksheet.Cells(i, "A"), objWorksheet.Cells(i, "L")).Value
            'Substituição dos valores da coluna F, códigos de recurso pelo tipo de veículo correspondente
            DestWorksheet.Cells(k, "F").Value = ValorRecurso
            k = k + 1
        Else
            Inconsistência_serviço_recurso = Inconsistência_serviço_recurso + 1
        End If
    Next i

    DestWorksheet.Range("P1").Value = "Nº Atendimentos sem expurgo"
    DestWorksheet.Range("P2").Value = LastRowObjWorksheet - 1 'subtrai por 1 desconsiderando cabeçalho
    DestWorksheet.Range("P7").Value = "Inconsistência serviço-recurso(Ex: ambulância em serviço mecânico)"
    DestWorksheet.Range("P8").Value = Inconsistência_serviço_recurso

    'COloca o nome da concessionária atual tratada na planilha "1.Instruções"
    ThisWorkbook.Sheets("1.Instruções").Cells(3, "F").Value = NomeDaNovaPlanilha

    objWorkbook.Close savechanges:=False
    ROpWorkbook.Close savechanges:=False

    Set DestWorksheet = Nothing
    Set DestWorkbook = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set ROpWorksheet = Nothing
    Set ROpWorkbook = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    ThisWorkbook.Save
    
    MsgBox "Tratamento de veículos concluído!"

End Sub

