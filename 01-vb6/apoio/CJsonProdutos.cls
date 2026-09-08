VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CJsonProdutos"
Option Explicit
Private texto As String
Private pos As Long

Public Function Ler(ByVal json As String) As Collection
    Dim itens As New Collection
    Dim produto As CProduto
    texto = json
    pos = 1
    Exigir "["
    If Not Consumir("]") Then
        Do
            Set produto = LerProduto()
            itens.Add produto
            If Consumir("]") Then Exit Do
            Exigir ","
        Loop
    End If
    Espacos
    If pos <= Len(texto) Then Falhar "Conteúdo após a lista JSON."
    Set Ler = itens
End Function

Private Function LerProduto() As CProduto
    Dim p As New CProduto
    Dim chave As String, numero As String
    Dim temNome As Boolean, temPreco As Boolean
    Exigir "{"
    If Not Consumir("}") Then
        Do
            chave = LerString()
            Exigir ":"
            Select Case chave
                Case "nome"
                    If temNome Then Falhar "Campo nome duplicado."
                    p.Nome = Trim$(LerString())
                    temNome = True
                Case "preco"
                    If temPreco Then Falhar "Campo preco duplicado."
                    numero = LerNumero()
                    p.Preco = DecimalJson(numero)
                    temPreco = True
                Case Else
                    IgnorarValor 0
            End Select
            If Consumir("}") Then Exit Do
            Exigir ","
        Loop
    End If
    If Not temNome Then Falhar "Existe um produto sem nome válido."
    If Len(Trim$(Replace(Replace(Replace(p.Nome, vbTab, " "), vbCr, " "), vbLf, " "))) = 0 Then Falhar "Existe um produto sem nome válido."
    If Not temPreco Then Falhar "Existe um produto sem preço válido."
    If p.Preco < CDec(0) Then Falhar "Existe um produto sem preço válido."
    Set LerProduto = p
End Function

Private Function Atual() As String
    Atual = Mid$(texto, pos, 1)
End Function

Private Sub Espacos()
    Do While pos <= Len(texto)
        Select Case Atual()
            Case " ", vbTab, vbCr, vbLf: pos = pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Function Consumir(ByVal token As String) As Boolean
    Espacos
    If Mid$(texto, pos, Len(token)) = token Then
        pos = pos + Len(token)
        Consumir = True
    End If
End Function

Private Sub Exigir(ByVal token As String)
    If Not Consumir(token) Then Falhar "JSON inválido: esperado '" & token & "'."
End Sub

Private Sub Falhar(ByVal mensagem As String)
    Err.Raise vbObjectError + 2100, "CJsonProdutos", mensagem & " Posição " & CStr(pos) & "."
End Sub

Private Function Digito(ByVal c As String) As Boolean
    If Len(c) = 1 Then Digito = (c >= "0" And c <= "9")
End Function

Private Function LerNumero() As String
    Dim inicio As Long
    Espacos
    inicio = pos
    If Atual() = "-" Then pos = pos + 1
    If Atual() = "0" Then
        pos = pos + 1
        If Digito(Atual()) Then Falhar "Zero inicial inválido."
    Else
        If Not Digito(Atual()) Then Falhar "Preço ou número JSON inválido."
        Do While Digito(Atual()): pos = pos + 1: Loop
    End If
    If Atual() = "." Then
        pos = pos + 1
        If Not Digito(Atual()) Then Falhar "Fração inválida."
        Do While Digito(Atual()): pos = pos + 1: Loop
    End If
    If LCase$(Atual()) = "e" Then
        pos = pos + 1
        If Atual() = "+" Or Atual() = "-" Then pos = pos + 1
        If Not Digito(Atual()) Then Falhar "Expoente inválido."
        Do While Digito(Atual()): pos = pos + 1: Loop
    End If
    LerNumero = Mid$(texto, inicio, pos - inicio)
End Function

Private Function DecimalJson(ByVal token As String) As Variant
    Dim e As Long, expoente As Long, ponto As Long, casas As Long
    Dim digitos As String, sinal As String, separador As String
    On Error GoTo Invalido
    e = InStr(1, token, "e", vbTextCompare)
    If e > 0 Then
        expoente = CLng(Mid$(token, e + 1))
        token = Left$(token, e - 1)
    End If
    If Left$(token, 1) = "-" Then
        sinal = "-"
        token = Mid$(token, 2)
    End If
    ponto = InStr(token, ".")
    If ponto > 0 Then casas = Len(token) - ponto
    digitos = Replace(token, ".", "")
    Do While Len(digitos) > 1 And Left$(digitos, 1) = "0"
        digitos = Mid$(digitos, 2)
    Loop
    If digitos = "0" Then
        DecimalJson = CDec(0)
        Exit Function
    End If
    casas = casas - expoente
    Do While casas > 0 And Right$(digitos, 1) = "0"
        digitos = Left$(digitos, Len(digitos) - 1)
        casas = casas - 1
    Loop
    ' Não arredondar silenciosamente números além da precisão Decimal.
    If casas > 28 Or casas < -28 Then GoTo Invalido
    If casas < 0 Then
        digitos = digitos & String$(-casas, "0")
        casas = 0
    End If
    If Len(digitos) > 29 Then GoTo Invalido
    separador = Mid$(CStr(1.5), 2, 1)
    If casas > 0 Then
        If Len(digitos) <= casas Then digitos = String$(casas - Len(digitos) + 1, "0") & digitos
        ponto = Len(digitos) - casas
        digitos = Left$(digitos, ponto) & separador & Mid$(digitos, ponto + 1)
    End If
    DecimalJson = CDec(sinal & digitos)
    Exit Function
Invalido:
    On Error GoTo 0
    Falhar "Preço fora da precisão Decimal suportada."
End Function

Private Function LerHex() As Long
    Dim i As Long, c As String, n As Long
    For i = 1 To 4
        c = UCase$(Atual())
        If Len(c) = 0 Then Falhar "Escape Unicode incompleto."
        n = InStr(1, "0123456789ABCDEF", c, vbBinaryCompare) - 1
        If n < 0 Then Falhar "Escape Unicode inválido."
        LerHex = LerHex * 16 + n
        pos = pos + 1
    Next
End Function

Private Function UnicodeChar(ByVal codigo As Long) As String
    If codigo > 32767 Then codigo = codigo - 65536
    UnicodeChar = ChrW$(codigo)
End Function

Private Function LerString() As String
    Dim c As String, resultado As String, codigo As Long, baixo As Long
    Exigir Chr$(34)
    Do While pos <= Len(texto)
        c = Atual()
        pos = pos + 1
        If c = Chr$(34) Then
            LerString = resultado
            Exit Function
        End If
        If c = "\" Then
            c = Atual()
            pos = pos + 1
            Select Case c
                Case Chr$(34), "\", "/": resultado = resultado & c
                Case "b": resultado = resultado & Chr$(8)
                Case "f": resultado = resultado & Chr$(12)
                Case "n": resultado = resultado & vbLf
                Case "r": resultado = resultado & vbCr
                Case "t": resultado = resultado & vbTab
                Case "u"
                    codigo = LerHex()
                    If codigo >= 55296 And codigo <= 56319 Then
                        If Mid$(texto, pos, 2) <> "\u" Then Falhar "Par Unicode incompleto."
                        pos = pos + 2
                        baixo = LerHex()
                        If baixo < 56320 Or baixo > 57343 Then Falhar "Par Unicode inválido."
                        resultado = resultado & UnicodeChar(codigo) & UnicodeChar(baixo)
                    Else
                        If codigo >= 56320 And codigo <= 57343 Then Falhar "Unicode isolado inválido."
                        resultado = resultado & UnicodeChar(codigo)
                    End If
                Case Else: Falhar "Escape JSON inválido."
            End Select
        Else
            codigo = AscW(c)
            If codigo >= 0 And codigo < 32 Then Falhar "Caractere de controle em string JSON."
            resultado = resultado & c
        End If
    Loop
    Falhar "String JSON sem fechamento."
End Function

Private Sub IgnorarValor(ByVal nivel As Long)
    Dim descarte As String
    If nivel > 64 Then Falhar "JSON excede 64 níveis."
    Espacos
    Select Case Atual()
        Case Chr$(34): descarte = LerString()
        Case "{"
            Exigir "{"
            If Consumir("}") Then Exit Sub
            Do
                descarte = LerString()
                Exigir ":"
                IgnorarValor nivel + 1
                If Consumir("}") Then Exit Do
                Exigir ","
            Loop
        Case "["
            Exigir "["
            If Consumir("]") Then Exit Sub
            Do
                IgnorarValor nivel + 1
                If Consumir("]") Then Exit Do
                Exigir ","
            Loop
        Case "t": Exigir "true"
        Case "f": Exigir "false"
        Case "n": Exigir "null"
        Case Else: descarte = LerNumero()
    End Select
End Sub
