Imports System.Xml.Linq
Imports System.Xml

Public Class FrmValidacao
    Private Sub finalizarTela()
        FrmPrincipal.vFrmValidacao = False
        Me.Dispose()
        Me.Close()
    End Sub

    Function criarDocumentoXML(ByVal vArquivo As String) As XDocument

        Dim vvDocumento As XDocument
        vvDocumento = XDocument.Load(vArquivo)

        Return vvDocumento '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>
        '<clientes>
        '   <cliente codigo="1">
        '      <nome>Macoratti</nome>
        '     <email>macoratti@yahoo.com</email>
        '</cliente>
        '<cliente codigo="2">
        '   <nome>Jessica</nome>
        '  <email>jessica@bol.com.br</email>
        '</cliente>
        '</clientes>

    End Function

    Private Sub carregarXML_Cabecalho(ByVal vArquivo As String)

        Dim vDS As New DataSet
        Dim vI As Integer = 0 'variavel para controle das tabelas
        Dim vJ As Integer = 0 'variavel para controle de campos da tabela
        Dim vK As Integer = 0 'qte de linhas por tabela
        Dim vL As Integer = 0 'valores por campos de tabela

        Dim vNomeTabela As String = ""
        Dim vNomeCampo As String = ""
        Dim vValorCampo As String = ""
        Dim vParametro As String = ""
        Dim vSQL As String = ""

        Try
            vDS.ReadXml(vArquivo)

            While vI < vDS.Tables.Count() 'Nome da tabela.

                'zerando valores
                vJ = 0
                'Nome da tabela
                vNomeTabela = vDS.Tables(vI).ToString 'dados

                'percorrendo nomes dos campos e criando parametros
                While vJ < vDS.Tables(vI).Columns.Count 'nome campos

                    Dim vX As String = vDS.Tables(vI).Columns(vJ).ColumnName
                    TextBox1.Text = TextBox1.Text & vX & vbCrLf
                    'Select Case vDS.Tables(vI).Columns(vJ).ColumnName
                    '    Case Else
                    'TextBox1.Text = TextBox1.Text & vDS.Tables(vI).Rows(0).Item(vJ) & vbCrLf
                    'End Select

                    vJ = vJ + 1
                End While

                vI = vI + 1
            End While

        Catch ex As Exception
            MsgBox("Erro ao recuperar os parâmetros." & vbCrLf & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Controle de Dados")
        End Try

    End Sub

    Private Sub validarArquivo(ByVal vArquivo As String)

        Dim vDS As New DataSet
        Dim vI As Integer
        Try
            carregarXML_Cabecalho(vArquivo)
            'vDS.ReadXml(vArquivo)
            '
            'While vI < vDS.Tables.Count()
            'Dim vNomeTabela As String = vDS.Tables(vI).ToString
            'vI = vI + 1
            'End While
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cmdAbrir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbrir.Click
        dlgAbrir.InitialDirectory = Environment.SpecialFolder.MyPictures
        dlgAbrir.Filter = "Arquivo xml|*.xml"
        dlgAbrir.FileName = ""

        Try
            If dlgAbrir.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtNomeArquivo.Text = dlgAbrir.FileName
            End If

        Catch ex As Exception
            MsgBox("Erro ao buscar arquivo xml." & vbCrLf & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.ApplicationModal, "Validação ANVISA")
            txtNomeArquivo.Text = ""
        End Try
    End Sub

    Private Sub cmdValidar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidar.Click
        If Len(txtNomeArquivo.Text) > 0 Then
            validarArquivo(txtNomeArquivo.Text)
        Else
            MsgBox("Digite ou encontre o arquivo XML para validação." & vbCrLf & "Operação cancelada.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Validação ANVISA")
        End If
    End Sub

    Private Sub cmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelar.Click
        finalizarTela()
    End Sub

    Private Sub FrmValidacao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class