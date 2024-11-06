Sub GerarTabelasPelasPlanilhasdaPasta() 'No Editor VBA entre em “Ferramentas” -> “Referências” e ative a biblioteca Microsoft Scripting Runtime
                                        'Necessário informar o caminho da pasta com planilhas
    
'Atendimento_APH_e_Socorro Mecânico.xlsm

'Performs recognition of the data source spreadsheet, operational resource spreadsheets and operational parameters,
'relates them to each concessionaire file contained in the same folder and prepares them for application of the initial data
'processing of Module1

'Realiza reconhecimento da planilha de origem dos dados, das planilhas de recursos operacionais e parâmetros operacionais,
'relacina-as para cada arquivo de concessionária contida numa mesma pasta e as prepara para aplicação do tratamento inicial
'de dados do Módulo1

'Version: 1.0

'Created by Matheus Nunes Reis on 07/08/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-OperacionalParameters/refs/heads/main/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    'ExcluirPlanilhasExcetoAsPrimeiras 'Algoritmo que exclui todas as planilhas exceto as primeiras

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(ThisWorkbook.Sheets("1.Instruções").Cells(1, "B").Value) ''''''''Necessário informar o caminho da pasta que contém planilhas

    Dim file As Object
    Dim i As Integer
    i = 1

    ' Cria um array dinâmico chamado "Caminho" para armazenar os caminhos somente de arquivos .xlsx ou .xls existentes na pasta
    Dim Caminho() As String
    Dim Caminho2 As String
    Dim Caminho3 As String
    ReDim Caminho(1 To 1)

    For Each file In folder.Files
        If ((Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Recursos Operacionais")) <> "Recursos Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Recursos Operacionais")) <> "Recursos Operacionais")) And _
            ((Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Parâmetros Operacionais")) <> "Parâmetros Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Parâmetros Operacionais")) <> "Parâmetros Operacionais")) Then
            ' Redimensiona o array para acomodar mais um caminho de pasta
            ReDim Preserve Caminho(1 To i)
            ' Salva o caminho da planilha no array
            Caminho(i) = file.Path
            i = i + 1
        End If
        ' Se o nome do arquivo contém "Recursos Operacionais" seu caminho é guardado na variável Caminho2
        If (Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Recursos Operacionais")) = "Recursos Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Recursos Operacionais")) = "Recursos Operacionais") Then 'InStr(1, fileName, "Recursos Operacionais", vbTextCompare) > 0
            Caminho2 = file.Path 'Caminho do arquivo Recursos Operacionais
        End If
        If (Right(file.Name, 4) = "xlsx" And Left(file.Name, Len("Parâmetros Operacionais")) = "Parâmetros Operacionais") Or _
            (Right(file.Name, 3) = "xls" And Left(file.Name, Len("Parâmetros Operacionais")) = "Parâmetros Operacionais") Then
            Caminho3 = file.Path 'Caminho do arquivo Parâmetros Operacionais
        End If
    Next file
    
    
    'Abrir planilha "Recursos Operacionais"
    Dim ROpWorkbook As Object, ROpWorksheet As Object
    Set ROpWorkbook = Workbooks.Open(Caminho2)
    Set ROpWorksheet = ROpWorkbook.Sheets(1)

    'Abrir planilha "Parâmetros Operacionais"
    Dim POpWorkbook As Object, POpWorksheet As Object
    Set POpWorkbook = Workbooks.Open(Caminho3)
    Set POpWorksheet = POpWorkbook.Sheets(1)

    ' Percorre o vetor Caminho e cria novas planilhas para cada concessionária
    For i = 1 To UBound(Caminho)
       'MsgBox "Caminho(" & i & ") = " & Caminho(i)  'Aparecer os nomes dos links armazenados na variável
       CopiarDadosAtendimentosAPHeSocorroMecânico_dePastaLocal Caminho(i), ROpWorksheet, POpWorksheet 'Inicia a rotina de nome CopiarDadosAtendimento tendo a variável "Caminho(i)" como argumento
    Next i

    ' Fecha a planilha Recursos Operacionais
    ROpWorkbook.Close savechanges:=False
    POpWorkbook.Close savechanges:=False

    Set ROpWorksheet = Nothing
    Set ROpWorkbook = Nothing
    Set POpWorksheet = Nothing
    Set POpWorkbook = Nothing
    Set folder = Nothing
    Set fso = Nothing
    'ThisWorkbook.Save
    MsgBox "Processamento concluído!"
    

End Sub


