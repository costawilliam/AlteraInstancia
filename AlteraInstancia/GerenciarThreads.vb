Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Threading

Public Class GerenciarThreads

    Private manualResetEvent As ManualResetEvent()
    Private thrRemove, thrInstala, thrExecuta As Thread
    Private diretorio As String = ""
    Private TipoBanco As String = ""
    Private queryInsert As String = ""
    Private queryUpdate As String = ""
    Private Inst As String = ""
    Private stringConexao As String = ""
    Private quant As Integer

    Public Sub RetornaComando(Inst As String)
        queryInsert = "insert into configs values ('NomeInstanciaServicos', '" + Inst + "')"
        queryUpdate = "update configs set valor = '" + Inst + "' where nome = 'NomeInstanciaServicos'"
    End Sub

    Public Sub New(ByVal dir As String, tipoDB As String, Quantidade As Integer, instancia As String, strConex As String)
        diretorio = dir
        TipoBanco = tipoDB
        quant = Quantidade
        Inst = instancia
        stringConexao = strConex
    End Sub

    Private Sub RemoveServicos()

        Dim pInfo As New ProcessStartInfo()
        pInfo.FileName = diretorio + "\remger.exe"
        Dim p As Process = Process.Start(pInfo)
        p.WaitForExit()

        Dim pInfo2 As New ProcessStartInfo()
        pInfo2.FileName = diretorio + "\remserv.exe"
        Dim p2 As Process = Process.Start(pInfo2)
        p2.WaitForExit()

        Thread.Sleep(3000)
    End Sub

    Private Sub ExecutaProcesso()
        RetornaComando(Inst)
        If (TipoBanco = "Access") Then

            'stringConexao = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & diretorio & "\psec.dat"
            'Seleciona(query)

            If quant = 0 Then
                'insert
                ExecutaComandoAccess(queryInsert, stringConexao)
            Else
                'update
                ExecutaComandoAccess(queryUpdate, stringConexao)
            End If

        ElseIf (TipoBanco = "Sql") Then
            'stringConexao = "Data Source=" & Form1.txtServidor.Text & ";Initial Catalog=" & Form1.txtNomeBanco.Text & ";Persist Security Info=True;User ID=" & Form1.txtUsuario.Text & ";Password=" & Form1.txtSenha.Text
            'Seleciona(query)

            If quant = 0 Then
                'insert
                ExecutaComandoSQL(queryInsert, stringConexao)
            Else
                'update
                ExecutaComandoSQL(queryUpdate, stringConexao)
            End If
        Else
            MessageBox.Show("Tipo de banco de dados não localizado!" + vbCrLf + "Verifique se o arquivo bc_config.ini está configurado corretamente, sem linhas de comentário.",
                                "Não foi possível identificar o tipo do banco de dados", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Thread.Sleep(2000)

    End Sub

    Private Sub InstalaServicos()

        Dim pInfo3 As New ProcessStartInfo()
        pInfo3.FileName = diretorio + "\instger.exe"
        Dim p3 As Process = Process.Start(pInfo3)
        p3.WaitForExit()

        Dim pInfo4 As New ProcessStartInfo()
        pInfo4.FileName = diretorio + "\instserv.exe"
        Dim p4 As Process = Process.Start(pInfo4)
        p4.WaitForExit()

    End Sub

    Public Sub ExecutaComandoSQL(query As String, stringconex As String)
        Try
            Using conexao As New SqlConnection
                conexao.ConnectionString = stringconex
                conexao.Open()

                Using comando As New SqlCommand
                    comando.Connection = conexao
                    comando.CommandText = query
                    comando.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As Exception
            MsgBox("ExecutaComandoSQl:" & ex.Message)
        End Try
    End Sub

    Public Sub ExecutaComandoAccess(query As String, stringconex As String)
        Try
            Using conexao As New OleDbConnection
                conexao.ConnectionString = stringconex
                conexao.Open()
                Using comando As New OleDbCommand
                    comando.Connection = conexao
                    comando.CommandText = query
                    comando.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MsgBox("ExecutaComandoAccess:" + ex.Message)
        End Try
    End Sub

    Public Sub Executar()
        Dim thrRemove, thrExecuta, thrInstala As Thread

        thrRemove = New Thread(AddressOf RemoveServicos)
        thrRemove.Name = "Thread Remove Serviços"
        thrRemove.Start()
        thrRemove.Join()

        thrExecuta = New Thread(AddressOf ExecutaProcesso)
        thrExecuta.Name = "Thread Executa Script"
        thrExecuta.Start()
        thrExecuta.Join()

        thrInstala = New Thread(AddressOf InstalaServicos)
        thrInstala.Name = "Thread Instala Serviços"
        thrInstala.Start()
        thrInstala.Join()
    End Sub

End Class