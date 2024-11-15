Sub ProtegeBotãoEspecífico_ou_todosBotões()
'Bloqueia botões, objetos e células
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Desbloqueia todas as células
    ws.Cells.Locked = False

    ' Bloqueia botão específico
    'ws.Shapes("Button 2").Locked = True

    ' Protege a planilha - necessário para aplicar bloqueio de botão
    ws.Protect Password:="InserirSenha", UserInterfaceOnly:=True
End Sub
