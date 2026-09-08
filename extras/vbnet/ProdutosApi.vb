Option Strict On
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.Json
Public Class ProdutosApi
        Public Function Buscar(endereco As String) As List(Of Produto)
            Dim url As Uri = Nothing
            If Not Uri.TryCreate(endereco, UriKind.Absolute, url) OrElse
                (url.Scheme <> "http" AndAlso url.Scheme <> "https") Then
                Throw New ArgumentException("Informe um endereço HTTP ou HTTPS válido.")
            End If

            Dim http As Object = Nothing
            Try
                ' Cria o componente MSXML do Windows.
                Dim tipo = Type.GetTypeFromProgID("MSXML2.ServerXMLHTTP.6.0", True)
                http = Activator.CreateInstance(tipo)

                ' Limites por etapa de rede, em milissegundos.
                CallByName(http, "setTimeouts", CallType.Method, 3000, 3000, 3000, 3000)
                CallByName(http, "open", CallType.Method, "GET", url.AbsoluteUri, False)
                CallByName(http, "setRequestHeader", CallType.Method, "Accept", "application/json")
                CallByName(http, "send", CallType.Method)

                Dim status = CInt(CallByName(http, "status", CallType.Get))
                If status <> 200 Then
                    Throw New InvalidOperationException($"A API respondeu HTTP {status}.")
                End If

                Dim json = CStr(CallByName(http, "responseText", CallType.Get))
                Return LerProdutos(json)
            Catch ex As COMException When ex.HResult = &H80072EE2
                Throw New TimeoutException("A API demorou para responder.", ex)
            Catch ex As COMException
                Throw New InvalidOperationException(
                    "Não foi possível consultar a API. Confira a URL e o container ps-estudo-api.", ex)
            Finally
                If http IsNot Nothing AndAlso Marshal.IsComObject(http) Then
                    Marshal.ReleaseComObject(http)
                End If
            End Try
        End Function

        Private Function LerProdutos(json As String) As List(Of Produto)
            Try
                Using documento = JsonDocument.Parse(json)
                    If documento.RootElement.ValueKind <> JsonValueKind.Array Then
                        Throw New InvalidDataException("Era esperada uma lista JSON.")
                    End If

                    Dim produtos As New List(Of Produto)
                    For Each item In documento.RootElement.EnumerateArray()
                        Dim nome As JsonElement
                        Dim preco As JsonElement
                        Dim valor As Decimal
                        If item.ValueKind <> JsonValueKind.Object OrElse
                            Not item.TryGetProperty("nome", nome) OrElse
                            nome.ValueKind <> JsonValueKind.String OrElse
                            String.IsNullOrWhiteSpace(nome.GetString()) Then
                            Throw New InvalidDataException("Existe um produto sem nome válido.")
                        End If
                        If Not item.TryGetProperty("preco", preco) OrElse
                            preco.ValueKind <> JsonValueKind.Number OrElse
                            Not preco.TryGetDecimal(valor) OrElse valor < 0 Then
                            Throw New InvalidDataException("Existe um produto sem preço válido.")
                        End If
                        produtos.Add(New Produto With {.Nome = nome.GetString().Trim(), .Preco = valor})
                    Next
                    Return produtos
                End Using
            Catch ex As JsonException
                Throw New InvalidDataException("A API retornou JSON inválido.", ex)
            End Try
        End Function
    End Class