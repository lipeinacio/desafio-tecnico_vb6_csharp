VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesafio
   BorderStyle     =   1
   Caption         =   "Desafio"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0
   ScaleHeight     =   6300
   ScaleWidth      =   6885
   StartUpPosition =   2
   BeginProperty Font
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0
      Italic          =   0
      Strikethrough   =   0
   EndProperty
   Begin VB.TextBox txtUrl
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Text            =   "http://10.0.2.2:18081/produtos"
      Top             =   570
      Width           =   2580
   End
   Begin VB.CommandButton btnCarregar
      Caption         =   "Carregar produtos"
      Height          =   345
      Left            =   2850
      TabIndex        =   1
      Top             =   570
      Width           =   3105
   End
   Begin MSComctlLib.ListView lvProdutos
      Height          =   3810
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6720
      View            =   3
      LabelEdit       =   1
      FullRowSelect   =   -1
      BorderStyle     =   0
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblStatus
      Caption         =   "Pronto para consultar"
      Height          =   585
      Left            =   180
      TabIndex        =   3
      Top             =   5565
      Width           =   6525
      WordWrap        =   -1
   End
   Begin VB.Timer tmrConsulta
      Enabled         =   0
      Interval        =   100
      Left            =   6120
      Top             =   570
   End
End
Attribute VB_Name = "frmDesafio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private api As CProdutosApi

Private Sub Form_Load()
    txtUrl.Text = UrlConfigurada()
    lvProdutos.ColumnHeaders.Add , , "Nome", 3000
    lvProdutos.ColumnHeaders.Add , , "Preço", 2700
    lblStatus.Caption = "Pronto para consultar"
End Sub

' 1. O botão inicia a consulta sem bloquear a tela.
Private Sub btnCarregar_Click()
    If Not btnCarregar.Enabled Then Exit Sub
    On Error GoTo Falha
    btnCarregar.Enabled = False
    txtUrl.Enabled = False
    lblStatus.Caption = "Consultando..."
    Set api = New CProdutosApi
    api.Iniciar Trim$(txtUrl.Text)
    tmrConsulta.Enabled = True
    Exit Sub
Falha:
    ExibirFalha Err.Number, Err.Description
End Sub

' 2. O Timer verifica se a resposta chegou.
Private Sub tmrConsulta_Timer()
    On Error GoTo Falha
    If api Is Nothing Then Exit Sub
    If Not api.Concluida Then Exit Sub
    tmrConsulta.Enabled = False
    ExibirProdutos api.Resultado()
    Finalizar
    Exit Sub
Falha:
    ExibirFalha Err.Number, Err.Description
End Sub

' 3. A lista só muda depois que a resposta foi validada.
Private Sub ExibirProdutos(ByVal produtos As Collection)
    Dim produto As CProduto
    Dim linha As MSComctlLib.ListItem
    lvProdutos.ListItems.Clear
    For Each produto In produtos
        Set linha = lvProdutos.ListItems.Add(, , produto.Nome)
        linha.SubItems(1) = PrecoBR(produto.Preco)
    Next
    If produtos.Count = 0 Then
        lblStatus.Caption = "Nenhum produto encontrado."
    Else
        lblStatus.Caption = CStr(produtos.Count) & " produtos carregados."
    End If
End Sub

Private Sub ExibirFalha(ByVal numero As Long, ByVal descricao As String)
    lblStatus.Caption = MensagemFalha(numero, descricao)
    If lvProdutos.ListItems.Count > 0 Then
        lblStatus.Caption = lblStatus.Caption & " A lista anterior foi mantida."
    End If
    Finalizar
End Sub

Private Sub Finalizar()
    tmrConsulta.Enabled = False
    If Not api Is Nothing Then api.Cancelar
    Set api = Nothing
    btnCarregar.Enabled = True
    txtUrl.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Finalizar
End Sub
