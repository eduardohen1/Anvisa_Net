<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmValidacao
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmValidacao))
        Me.cmdValidar = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdAbrir = New System.Windows.Forms.Button
        Me.txtNomeArquivo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdCancelar = New System.Windows.Forms.Button
        Me.dlgAbrir = New System.Windows.Forms.OpenFileDialog
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdValidar
        '
        Me.cmdValidar.Image = CType(resources.GetObject("cmdValidar.Image"), System.Drawing.Image)
        Me.cmdValidar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdValidar.Location = New System.Drawing.Point(196, 93)
        Me.cmdValidar.Name = "cmdValidar"
        Me.cmdValidar.Size = New System.Drawing.Size(126, 43)
        Me.cmdValidar.TabIndex = 5
        Me.cmdValidar.Text = "&Validar"
        Me.cmdValidar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdValidar.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdAbrir)
        Me.GroupBox1.Controls.Add(Me.txtNomeArquivo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(442, 75)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'cmdAbrir
        '
        Me.cmdAbrir.Image = CType(resources.GetObject("cmdAbrir.Image"), System.Drawing.Image)
        Me.cmdAbrir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAbrir.Location = New System.Drawing.Point(333, 16)
        Me.cmdAbrir.Name = "cmdAbrir"
        Me.cmdAbrir.Size = New System.Drawing.Size(98, 39)
        Me.cmdAbrir.TabIndex = 2
        Me.cmdAbrir.Text = "&Abrir"
        Me.cmdAbrir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAbrir.UseVisualStyleBackColor = True
        '
        'txtNomeArquivo
        '
        Me.txtNomeArquivo.Location = New System.Drawing.Point(9, 35)
        Me.txtNomeArquivo.Name = "txtNomeArquivo"
        Me.txtNomeArquivo.Size = New System.Drawing.Size(318, 20)
        Me.txtNomeArquivo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Nome do arquivo:"
        '
        'cmdCancelar
        '
        Me.cmdCancelar.Image = CType(resources.GetObject("cmdCancelar.Image"), System.Drawing.Image)
        Me.cmdCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdCancelar.Location = New System.Drawing.Point(328, 93)
        Me.cmdCancelar.Name = "cmdCancelar"
        Me.cmdCancelar.Size = New System.Drawing.Size(126, 43)
        Me.cmdCancelar.TabIndex = 6
        Me.cmdCancelar.Text = "&Cancelar"
        Me.cmdCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancelar.UseVisualStyleBackColor = True
        '
        'dlgAbrir
        '
        Me.dlgAbrir.FileName = "OpenFileDialog1"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 147)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(336, 77)
        Me.TextBox1.TabIndex = 8
        '
        'FrmValidacao
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(576, 236)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmdValidar)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdCancelar)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmValidacao"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Validação"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdValidar As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAbrir As System.Windows.Forms.Button
    Friend WithEvents txtNomeArquivo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdCancelar As System.Windows.Forms.Button
    Friend WithEvents dlgAbrir As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
