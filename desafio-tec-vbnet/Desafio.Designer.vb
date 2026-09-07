<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Desafio
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        txtUrl = New TextBox()
        btnCarregar = New Button()
        lvProdutos = New ListView()
        Coluna = New ColumnHeader()
        ColumnHeader1 = New ColumnHeader()
        lblStatus = New Label()
        SuspendLayout()
        '
        ' txtUrl
        '
        txtUrl.AccessibleName = "txtUrl"
        txtUrl.Location = New Point(12, 38)
        txtUrl.Name = "txtUrl"
        txtUrl.Size = New Size(172, 23)
        txtUrl.TabIndex = 0
        txtUrl.Text = "http://127.0.0.1:18081/produtos"
        '
        ' btnCarregar
        '
        btnCarregar.AccessibleName = "btnCarregar"
        btnCarregar.Location = New Point(190, 38)
        btnCarregar.Name = "btnCarregar"
        btnCarregar.Size = New Size(207, 23)
        btnCarregar.TabIndex = 1
        btnCarregar.Text = "Carregar produtos"
        btnCarregar.UseVisualStyleBackColor = True
        '
        ' lvProdutos
        '
        lvProdutos.AccessibleName = "lvProdutos"
        lvProdutos.BorderStyle = BorderStyle.None
        lvProdutos.Columns.AddRange(New ColumnHeader() {Coluna, ColumnHeader1})
        lvProdutos.FullRowSelect = True
        lvProdutos.Location = New Point(12, 108)
        lvProdutos.Name = "lvProdutos"
        lvProdutos.Size = New Size(435, 254)
        lvProdutos.TabIndex = 0
        lvProdutos.UseCompatibleStateImageBehavior = False
        lvProdutos.View = View.Details
        '
        ' Coluna
        '
        Coluna.Text = "Nome"
        Coluna.Width = 200
        '
        ' ColumnHeader1
        '
        ColumnHeader1.Text = "Preço"
        ColumnHeader1.Width = 180
        '
        ' lblStatus
        '
        lblStatus.AccessibleName = "lblStatus"
        lblStatus.Location = New Point(22, 371)
        lblStatus.Name = "lblStatus"
        lblStatus.Size = New Size(192, 21)
        lblStatus.TabIndex = 2
        lblStatus.Text = "Pronto para consultar"
        '
        ' Desafio
        '
        AccessibleName = "lvProdutos"
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(459, 401)
        Controls.Add(lblStatus)
        Controls.Add(lvProdutos)
        Controls.Add(btnCarregar)
        Controls.Add(txtUrl)
        Name = "Desafio"
        Text = "Desafio"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtUrl As TextBox
    Friend WithEvents btnCarregar As Button
    Friend WithEvents lvProdutos As ListView
    Friend WithEvents lblStatus As Label
    Friend WithEvents Coluna As ColumnHeader
    Friend WithEvents ColumnHeader1 As ColumnHeader

End Class
