Attribute VB_Name = "modTestes"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal milissegundos As Long)
Private arquivoLog As Integer
Private falhas As Long
Private total As Long

Public Sub ExecutarTestes()
    Dim parser As New CJsonProdutos
    Dim produtos As Collection
    On Error GoTo Fatal
    arquivoLog = FreeFile
    Open App.Path & "\resultado-testes.txt" For Output As #arquivoLog
    Print #arquivoLog, "Testes VB6 nativos - " & CStr(Now)
    TestarJson "lista vazia", "[]", True, 0
    TestarJson "produtos", "[{""nome"":""Caderno"",""preco"":24.9},{""nome"":""Caneta"",""preco"":3.5}]", True, 2
    TestarJson "campos extras", "[{""nome"":""A"",""preco"":0,""extra"":[null,true,false,{}]}]", True, 1
    TestarJson "Unicode e expoente", "[{""nome"":""Caf\u00e9"",""preco"":1.25e2}]", True, 1
    TestarJson "raiz objeto", "{}", False, 0
    TestarJson "item null", "[null]", False, 0
    TestarJson "nome vazio", "[{""nome"":"" "",""preco"":1}]", False, 0
    TestarJson "nome ausente", "[{""preco"":1}]", False, 0
    TestarJson "preco ausente", "[{""nome"":""A""}]", False, 0
    TestarJson "preco negativo", "[{""nome"":""A"",""preco"":-1}]", False, 0
    TestarJson "preco string", "[{""nome"":""A"",""preco"":""1""}]", False, 0
    TestarJson "preco booleano", "[{""nome"":""A"",""preco"":true}]", False, 0
    TestarJson "overflow", "[{""nome"":""A"",""preco"":1e100}]", False, 0
    TestarJson "virgula final", "[{""nome"":""A"",""preco"":1},]", False, 0
    TestarJson "conteudo final", "[]x", False, 0
    TestarJson "zero inicial", "[{""nome"":""A"",""preco"":01}]", False, 0
    TestarJson "escape invalido", "[{""nome"":""\q"",""preco"":1}]", False, 0
    TestarJson "duplicado", "[{""nome"":""A"",""preco"":1,""preco"":2}]", False, 0
    Set produtos = parser.Ler("[{""nome"":""Caf\u00e9"",""preco"":1.25e2}]")
    Registrar "valor Decimal e Unicode", produtos(1).Preco = CDec(125) And produtos(1).Nome = "Café"
    Registrar "formatacao BR", PrecoBR(CDec(1234.5)) = "R$ 1.234,50"
    TestarApi "", 3, False
    TestarApi "?cenario=vazio", 0, False
    TestarApi "?cenario=erro", 0, True
    TestarApi "?cenario=json", 0, True
    TestarApi "?cenario=preco", 0, True
    TestarApi "?cenario=lento", 0, True
    Print #arquivoLog, "TOTAL=" & total & " FALHAS=" & falhas
    Close #arquivoLog
    Exit Sub
Fatal:
    If arquivoLog <> 0 Then
        Print #arquivoLog, "ERRO FATAL: " & CStr(Err.Number) & " " & Err.Description
        Close #arquivoLog
    End If
End Sub

Private Sub Registrar(ByVal nome As String, ByVal passou As Boolean)
    total = total + 1
    If passou Then
        Print #arquivoLog, "PASS | " & nome
    Else
        falhas = falhas + 1
        Print #arquivoLog, "FAIL | " & nome
    End If
End Sub

Private Sub TestarJson(ByVal nome As String, ByVal json As String, ByVal valido As Boolean, ByVal quantidade As Long)
    Dim parser As New CJsonProdutos, itens As Collection
    On Error GoTo Rejeitado
    Set itens = parser.Ler(json)
    Registrar nome, valido And itens.Count = quantidade
    Exit Sub
Rejeitado:
    Registrar nome, Not valido
End Sub

Private Sub TestarApi(ByVal sufixo As String, ByVal quantidade As Long, ByVal esperaErro As Boolean)
    Dim api As New CProdutosApi, itens As Collection
    Dim inicio As Date
    Dim codigo As Long, mensagem As String, esperado As Boolean
    On Error GoTo Falhou
    api.Iniciar UrlConfigurada() & sufixo
    inicio = Now
    Do While Not api.Concluida
        DoEvents
        Sleep 10
        If DateDiff("s", inicio, Now) > 20 Then Err.Raise vbObjectError + 2199, , "Prazo do teste excedido."
    Loop
    Set itens = api.Resultado()
    Registrar "API " & sufixo, Not esperaErro And itens.Count = quantidade
    api.Cancelar
    Exit Sub
Falhou:
    codigo = Err.Number
    mensagem = Err.Description
    Select Case sufixo
        Case "?cenario=erro"
            esperado = codigo = vbObjectError + 2103 And InStr(mensagem, "503") > 0
        Case "?cenario=json", "?cenario=preco"
            esperado = codigo = vbObjectError + 2100
        Case "?cenario=lento"
            esperado = codigo = &H80072EE2 And DateDiff("s", inicio, Now) <= 5
            mensagem = mensagem & " (" & CStr(DateDiff("s", inicio, Now)) & " s)"
    End Select
    Registrar "API " & sufixo & " | " & mensagem, esperaErro And esperado
    api.Cancelar
End Sub
