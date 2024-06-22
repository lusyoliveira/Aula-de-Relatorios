Imports System.IO
Module conexao
    Dim x As Integer
    Private aConexao As New ADODB.Connection
    Private aconexao2 As New ADODB.Connection

    Dim acess_server, acess_db, acess_usr, acess_pwd


    Public Function OpenRecordset(ByVal sqlquery As String)

again:

        If Directory.Exists("c:\temp") = False Then
            Directory.CreateDirectory("c:\temp")
            Dim arq As StreamWriter = New StreamWriter("c:\temp\config_softclinica.ini")
            arq.WriteLine("(local)")
            arq.Close()
            acess_server = "(local)"
        Else

            If File.Exists("c:\temp\config_softclinica.ini") = False Then
                Dim arq2 As StreamWriter = New StreamWriter("c:\temp\config_softclinica.ini")
                arq2.WriteLine("(local)")
                arq2.Close()
            End If
            Dim arq As StreamReader = New StreamReader("c:\temp\config_softclinica.ini")
            acess_server = arq.ReadLine()
            arq.Close()
        End If
        If File.Exists("c:\temp\banco_softclinica.ini") = False Then
            Dim arq2 As StreamWriter = New StreamWriter("c:\temp\banco_softclinica.ini")
            arq2.WriteLine("dbSoftClinica")
            arq2.Close()
        End If
        Dim arq3 As StreamReader = New StreamReader("c:\temp\banco_softclinica.ini")
        acess_db = arq3.ReadLine()
        arq3.Close()

        '   acess_db = "dbSoftClinica"

        Dim aResultado As New ADODB.Recordset
        '        acess_server = "servidor"

        'acess_db = "dbSoftClinica_jundiai"
        acess_usr = "usraureo"
        acess_pwd = "aureo"

        Try


            If aConexao.State = 0 Then

                aConexao.ConnectionString = "driver={sql server};" & _
                                                "server=" + acess_server + ";" & _
                                                "Database=" + acess_db + ";" & _
                                                "PWD=" + acess_pwd + ";" & _
                                                "UID=" + acess_usr + ";"
                aConexao.ConnectionTimeout = 30
                aConexao.CommandTimeout = 180
                aConexao.Open()
            End If
        Catch ex As Exception
            If MsgBox("Atenção: O servidor não foi encontrado! Será realizada uma nova tentativa de localização ! Por favor, clique em OK !", MsgBoxStyle.Information + MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                Dim arq2 As StreamWriter = New StreamWriter("c:\temp\config_softclinica.ini")
                arq2.WriteLine("servidor")
                arq2.Close()
                GoTo again
            Else

                End
            End If
        End Try


        aResultado.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        aResultado.Open(sqlquery, aConexao, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        OpenRecordset = aResultado
        aResultado = Nothing

    End Function
    Public Function abrir(ByVal sql As String)

        Dim aResultado As New ADODB.Recordset
        Dim tbconfig As ADODB.Recordset
        tbconfig = OpenRecordset("Select temporario from tbconfig")
        If tbconfig.RecordCount = 0 Then
            Return Nothing
            Exit Function
        End If
        acess_db = tbconfig.Fields("temporario").Value.ToString
        If File.Exists(acess_db) = False Then
            Return Nothing
            Exit Function
        End If

        If sql.ToUpper = "FECHAR" Then
            aconexao2.Close()
            Return False
            Exit Function
        End If

        If aconexao2.State = 0 Then
            aconexao2.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & acess_db
            aconexao2.Open()
        End If

        'aResultado2.Open(sql, aconexao2, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        aResultado.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        aResultado.Open(sql, aconexao2, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        abrir = aResultado
        aResultado = Nothing

    End Function
   
End Module
