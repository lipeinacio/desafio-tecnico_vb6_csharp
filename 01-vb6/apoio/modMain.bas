Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Main()
    If LCase$(Trim$(Command$)) = "/test" Then
        ExecutarTestes
    Else
        frmDesafio.Show
    End If
End Sub

Public Function UrlConfigurada() As String
    Dim arquivo As Integer, endereco As String
    UrlConfigurada = "http://10.0.2.2:18081/produtos"
    On Error GoTo Fim
    arquivo = FreeFile
    Open App.Path & "\url.txt" For Input As #arquivo
    Line Input #arquivo, endereco
    Close #arquivo
    arquivo = 0
    If Len(Trim$(endereco)) > 0 Then UrlConfigurada = Trim$(endereco)
Fim:
    On Error Resume Next
    If arquivo <> 0 Then Close #arquivo
End Function

Public Function PrecoBR(ByVal valor As Variant) As String
    Dim s As String, inteiro As String, grupo As String
    s = Format$(valor, "0.00")
    inteiro = Left$(s, Len(s) - 3)
    Do While Len(inteiro) > 3
        grupo = "." & Right$(inteiro, 3) & grupo
        inteiro = Left$(inteiro, Len(inteiro) - 3)
    Loop
    PrecoBR = "R$ " & inteiro & grupo & "," & Right$(s, 2)
End Function

Public Function RelogioMs() As Double
    Dim tick As Long
    tick = GetTickCount()
    RelogioMs = CDbl(tick)
    If tick < 0 Then RelogioMs = RelogioMs + 4294967296#
End Function

Public Sub ValidarEndereco(ByVal endereco As String)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^https?://[^\s/?#]+([/?#][^\s]*)?$"
    regex.IgnoreCase = True
    If Not regex.Test(endereco) Then Err.Raise vbObjectError + 2101, , "Informe um endereço HTTP ou HTTPS válido."
End Sub

Public Function MensagemFalha(ByVal numero As Long, ByVal descricao As String) As String
    Select Case numero
        Case &H80072EE2
            MensagemFalha = "A API demorou para responder."
        Case vbObjectError + 2100 To vbObjectError + 2103
            MensagemFalha = descricao
        Case 429
            MensagemFalha = "Componente indisponível. Confira a instalação do MSXML 6 no Windows."
        Case Else
            MensagemFalha = "Não foi possível consultar a API. Confira a URL, a rede e se a API está ligada."
    End Select
End Function
