Option Strict On
Imports System.Globalization

Public Class Desafio
    Private Async Sub btnCarregar_Click(sender As Object, e As EventArgs) Handles btnCarregar.Click
        btnCarregar.Enabled = False
        txtUrl.Enabled = False
        lblStatus.Text = "Consultando..."

        Try
            Dim endereco = txtUrl.Text.Trim()
            Dim produtos As List(Of Produto) = Nothing
            Dim erro As String = Nothing

            ' Consulta fora da thread da tela; captura a falha na própria tarefa.
            Await Task.Run(
                Sub()
                    Try
                        Dim api As New ProdutosApi()
                        produtos = api.Buscar(endereco)
                    Catch ex As Exception
                        erro = ex.Message
                    End Try
                End Sub)

            If IsDisposed OrElse Disposing Then Return

            If erro IsNot Nothing Then
                lblStatus.Text = erro & If(lvProdutos.Items.Count > 0,
                    " A lista anterior foi mantida.", "")
                Return
            End If

            lvProdutos.BeginUpdate()
            Try
                lvProdutos.Items.Clear()
                For Each produto In produtos
                    Dim linha As New ListViewItem(produto.Nome)
                    linha.SubItems.Add(produto.Preco.ToString(
                        "C2", CultureInfo.GetCultureInfo("pt-BR")))
                    lvProdutos.Items.Add(linha)
                Next
            Finally
                lvProdutos.EndUpdate()
            End Try

            lblStatus.Text = If(produtos.Count = 0,
                "Nenhum produto encontrado.",
                $"{produtos.Count} produtos carregados.")
        Catch ex As Exception
            If Not IsDisposed AndAlso Not Disposing Then
                lblStatus.Text = "Falha ao atualizar a tela: " & ex.Message
            End If
        Finally
            If Not IsDisposed AndAlso Not Disposing Then
                btnCarregar.Enabled = True
                txtUrl.Enabled = True
            End If
        End Try
    End Sub
End Class
