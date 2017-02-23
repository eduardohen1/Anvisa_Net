Public Class FrmPrincipal
    'Variaveis gerais
    Public vConnection As OleDb.OleDbConnection

    'Controle de Telas:
    Public vFrmValidacao As Boolean

    'Funções para chamar telas:
    Private Sub chamaFrmValidacao()
        Dim vChildForm As New FrmValidacao
        If Not vFrmValidacao Then
            vChildForm.MdiParent = Me
            vChildForm.StartPosition = FormStartPosition.CenterScreen
            vChildForm.Show()
            vChildForm.WindowState = FormWindowState.Normal
            vFrmValidacao = True
        Else
            MsgBox("Atenção!" & vbCrLf & "Janela já está aberta. Verifique...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Validação ANVISA")
        End If

    End Sub

    Private Sub textoStatus(ByVal vTexto As String)
        statusText.Text = vTexto
    End Sub

    Private Sub ToolStripButton1_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.MouseEnter
        textoStatus("Validar arquivo")
    End Sub

    Private Sub ToolStripButton1_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.MouseLeave
        textoStatus("")
    End Sub

    Private Sub ToolStripButton2_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.MouseEnter
        textoStatus("Movimentações de estoque")
    End Sub

    Private Sub ToolStripButton2_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.MouseLeave
        textoStatus("")
    End Sub

    Private Sub ToolStripButton3_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.MouseEnter
        textoStatus("Estoque inicial")
    End Sub

    Private Sub ToolStripButton3_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.MouseLeave
        textoStatus("")
    End Sub

    Private Sub ToolStripButton4_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.MouseEnter
        textoStatus("Sair do sistema")
    End Sub

    Private Sub ToolStripButton4_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.MouseLeave
        textoStatus("")
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub FrmPrincipal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        chamaFrmValidacao()
    End Sub
End Class
