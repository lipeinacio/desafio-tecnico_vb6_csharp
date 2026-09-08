VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CProdutosApi"
Option Explicit
Private Const PRAZO_TOTAL_MS As Double = 3000
Private http As Object
Private inicioMs As Double

Public Sub Iniciar(ByVal endereco As String)
    Cancelar
    ValidarEndereco endereco
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    inicioMs = RelogioMs()
    http.setTimeouts 3000, 3000, 3000, 3000
    http.Open "GET", endereco, True
    http.setRequestHeader "Accept", "application/json"
    http.send
End Sub

Public Function Concluida() As Boolean
    Dim decorrido As Double
    If http Is Nothing Then Err.Raise vbObjectError + 2102, , "Nenhuma consulta iniciada."
    decorrido = RelogioMs() - inicioMs
    If decorrido < 0 Then decorrido = decorrido + 4294967296#
    ' Limite total independente dos timeouts internos do MSXML assincrono.
    If decorrido >= PRAZO_TOTAL_MS Then
        Cancelar
        Err.Raise &H80072EE2, "CProdutosApi", "A API demorou para responder."
    End If
    ' Zero segundos: não bloqueia a interface enquanto a rede trabalha.
    Concluida = http.waitForResponse(0)
End Function


Public Function Resultado() As Collection
    Dim parser As New CJsonProdutos
    Dim status As Long
    status = CLng(http.status)
    If status <> 200 Then Err.Raise vbObjectError + 2103, , "A API respondeu HTTP " & CStr(status) & "."
    Set Resultado = parser.Ler(CStr(http.responseText))
    Set http = Nothing
End Function

Public Sub Cancelar()
    On Error Resume Next
    If Not http Is Nothing Then http.abort
    Set http = Nothing
End Sub

Private Sub Class_Terminate()
    Cancelar
End Sub
