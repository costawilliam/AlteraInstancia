Imports System.Diagnostics
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Threading
Public Class FormPrincipal
    Public TipoBanco As String
    Public dir As String
    Public stringConexao As String
    Public quantidade As Integer
    Public querySelect As String = "Select COUNT(*) From configs Where nome = 'NomeInstanciaServicos'"
    Public queryVersao As String = "select valor from configs where nome = 'VersaoBanco'"
    Public versao As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Visible = False
        rbAccess.Checked = True
        GroupBox2.Visible = False
        TipoBanco = "Access"
        dir = Directory.GetCurrentDirectory
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        Dim Executar As Boolean = False

        If TesteConexao() = True Then
            VerificaVersao(queryVersao)
            If (versao < 78) Then
                MsgBox("Não é possível utilizar serviços com instância nomeada nesta versão do Ponto Secullum 4. Atualize o sistema e o banco de dados.")
                Exit Sub
            Else
                If txtNomeInstancia.Text.Trim = "" Then
                    MsgBox("Preencha o nome da instância corretamente!")
                    txtNomeInstancia.Focus()
                Else
                    If rbAccess.Checked = True Then
                        If Not IO.File.Exists(dir + "\psec.dat") Or Not IO.File.Exists(dir + "\remger.exe") Or Not IO.File.Exists(dir + "\remserv.exe") Or Not IO.File.Exists(dir + "\instger.exe") Or Not IO.File.Exists(dir + "\instger.exe") Then
                            MsgBox("Um ou mais arquivos necessários não foi localizado, certifique se que os seguintes arquivos  existem na pasta do sistema:" + vbCrLf + "   - psec.dat" + vbCrLf + "   - remger.exe" + vbCrLf + "   - remserv.exe" + vbCrLf + "   - insger.exe" + vbCrLf + "   - insterv.exe")
                            Executar = False
                        Else
                            TipoBanco = "Access"
                            Executar = True
                        End If
                    Else
                        If Not IO.File.Exists(dir + "\remger.exe") Or Not IO.File.Exists(dir + "\remserv.exe") Or Not IO.File.Exists(dir + "\instger.exe") Or Not IO.File.Exists(dir + "\instger.exe") Then
                            MsgBox("Um ou mais arquivos necessários não foi localizado, certifique se que os seguintes arquivos  existem na pasta do sistema:" + vbCrLf + "   - remger.exe" + vbCrLf + "   - remserv.exe" + vbCrLf + "   - insger.exe" + vbCrLf + "   - insterv.exe")
                            Executar = False
                        Else
                            TipoBanco = "Sql"
                            Executar = True
                        End If
                    End If
                End If
            End If

            If Executar = True Then
                btnAlterar.Enabled = False

                Seleciona(querySelect)

                Dim gt As New GerenciarThreads(dir, TipoBanco, quantidade, txtNomeInstancia.Text, stringConexao)
                gt.Executar()
                btnAlterar.Enabled = True
            Else

            End If

        Else
            MsgBox("Não foi possível conectar com o banco de dados!")
            Exit Sub
        End If
    End Sub

    Public Sub Seleciona(query As String)
        Dim tabela As New DataTable

        If TipoBanco = "Access" Then

            If Not IO.File.Exists(dir + "\psec.dat") Then
                MsgBox("Arquivo 'psec.dat' não encontrado!")
            Else
                Try
                    stringConexao = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & Directory.GetCurrentDirectory & "\psec.dat"
                    Using conexao As New OleDbConnection
                        conexao.ConnectionString = stringConexao
                        conexao.Open()

                        Using comando As New OleDbCommand
                            comando.Connection = conexao
                            comando.CommandText = query

                            Using leitor = comando.ExecuteReader
                                tabela.Load(leitor)
                            End Using
                        End Using
                        DataGridView1.Visible = True
                        DataGridView1.DataSource = tabela
                        quantidade = DataGridView1.CurrentRow.Cells(0).Value
                        DataGridView1.Visible = False
                    End Using
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
            End If
        ElseIf TipoBanco = "Sql" Then
            stringConexao = "Data Source=" & txtServidor.Text & ";Initial Catalog=" & txtNomeBanco.Text & ";Persist Security Info=True;User ID=" & txtUsuario.Text & ";Password=" & txtSenha.Text
            Try
                Using conexao As New SqlConnection
                    conexao.ConnectionString = stringConexao
                    conexao.Open()
                    Using comando As New SqlCommand
                        comando.Connection = conexao
                        comando.CommandText = query
                        Using leitor = comando.ExecuteReader
                            tabela.Load(leitor)
                        End Using
                    End Using
                    DataGridView1.Visible = True
                    DataGridView1.DataSource = tabela
                    quantidade = DataGridView1.CurrentRow.Cells(0).Value
                    DataGridView1.Visible = False
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Public Sub VerificaVersao(query As String)
        Dim tabela As New DataTable
        Dim retornoConsulta As String = ""
        Dim retornoConsultaSplit As String() = Nothing

        If TipoBanco = "Access" Then

            If Not IO.File.Exists(dir + "\psec.dat") Then
                MsgBox("Arquivo 'psec.dat' não encontrado!")
            Else
                Try
                    stringConexao = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & Directory.GetCurrentDirectory & "\psec.dat"
                    Using conexao As New OleDbConnection
                        conexao.ConnectionString = stringConexao
                        conexao.Open()

                        Using comando As New OleDbCommand
                            comando.Connection = conexao
                            comando.CommandText = query

                            Using leitor = comando.ExecuteReader
                                tabela.Load(leitor)
                            End Using
                        End Using
                        DataGridView1.Visible = True
                        DataGridView1.DataSource = tabela
                        retornoConsulta = DataGridView1.CurrentRow.Cells(0).Value
                        retornoConsultaSplit = retornoConsulta.Split(".")
                        versao = retornoConsultaSplit(0)
                        DataGridView1.Visible = False
                    End Using
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        ElseIf TipoBanco = "Sql" Then
            stringConexao = "Data Source=" & txtServidor.Text & ";Initial Catalog=" & txtNomeBanco.Text & ";Persist Security Info=True;User ID=" & txtUsuario.Text & ";Password=" & txtSenha.Text
            Try
                Using conexao As New SqlConnection
                    conexao.ConnectionString = stringConexao
                    conexao.Open()
                    Using comando As New SqlCommand
                        comando.Connection = conexao
                        comando.CommandText = query
                        Using leitor = comando.ExecuteReader
                            tabela.Load(leitor)
                        End Using
                    End Using
                    DataGridView1.Visible = True
                    DataGridView1.DataSource = tabela
                    retornoConsulta = DataGridView1.CurrentRow.Cells(0).Value
                    retornoConsultaSplit = retornoConsulta.Split(".")
                    versao = retornoConsultaSplit(0)
                    DataGridView1.Visible = False
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Public Function TesteConexao() As Boolean
        Dim retorno As Boolean = True
        If rbAccess.Checked = True Then
            Try
                stringConexao = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & Directory.GetCurrentDirectory & "\psec.dat"

                Using conexao As New OleDbConnection
                    conexao.ConnectionString = stringConexao
                    conexao.Open()
                End Using

            Catch ex As Exception
                retorno = False
            End Try
        ElseIf rbSQL.Checked = True Then
            Try
                stringConexao = "Data Source=" & txtServidor.Text & ";Initial Catalog=" & txtNomeBanco.Text & ";Persist Security Info=True;User ID=" & txtUsuario.Text & ";Password=" & txtSenha.Text

                Using conexao As New SqlConnection
                    conexao.ConnectionString = stringConexao
                    conexao.Open()
                End Using

            Catch ex As Exception
                retorno = False
            End Try
        Else
            retorno = False
        End If

        Return retorno
    End Function

    Private Sub rbSQL_CheckedChanged(sender As Object, e As EventArgs) Handles rbSQL.CheckedChanged
        If rbSQL.Checked = True Then
            GroupBox2.Visible = True
            TipoBanco = "Sql"
        End If
    End Sub

    Private Sub rbAccess_CheckedChanged(sender As Object, e As EventArgs) Handles rbAccess.CheckedChanged
        If rbAccess.Checked = True Then
            GroupBox2.Visible = False
            TipoBanco = "Access"
        End If
    End Sub

    Private CaracteresNaoAceitos As String = "ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç~^´`¨¢%$#@!&*()_-+=§¹²³£¢¬ªº°\|<,>.:;?/°}]º[{ª§¬'"""

    Private Sub txtNomeInstancia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNomeInstancia.KeyPress
        If CaracteresNaoAceitos.IndexOf(e.KeyChar) > 0 Then
            e.Handled = True
        End If
    End Sub

End Class
