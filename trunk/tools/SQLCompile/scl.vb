Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Imports System.Collections.Specialized

Public Class scl
    Private sResult As String = ""
    Private bErrorLatch As Boolean
    Private sConnectString As String
    Private sDataBase As String

    Public Event Notify(ByVal Msg As String, ByVal Type As String)

    Public Sub Process(ByVal sList As String, ByVal sData As String)
        Dim line As String
        Dim s As String
        Dim sFile As String
        Dim sType As String
        Dim sPWD As String
        Dim bWrite As Boolean

        sPWD = Environment.CurrentDirectory
        Try
            If sData <> "" Then
                sDataBase = sData
                sConnectString = GetConnectString(sDataBase)
                WriteLine("Database context is " & sDataBase & ".", "I")
            Else
                s = "default"
                Dim instance As New System.Configuration.AppSettingsReader
                instance.GetValue(s, s.GetType).ToString()
                sDataBase = instance.GetValue(s, s.GetType).ToString()
                If sDataBase <> "" Then
                    sConnectString = GetConnectString(sDataBase)
                    WriteLine("Database context defaulted to " & sDataBase & ".", "I")
                End If
            End If
            s = Path.GetDirectoryName(sList)
            If s <> "" Then
                sFile = Path.GetFileName(sList)
                Environment.CurrentDirectory = s
            Else
                sFile = sList
            End If

            Dim file As New System.IO.StreamReader(sFile)
            ' Read and display the lines from the file until the end 
            ' of the file is reached.
            Do
                line = file.ReadLine()
                bWrite = True
                If line.TrimEnd <> "" Then
                    Select Case Mid(line, 1, 1)
                        Case "#"    ' ignore comments

                        Case "{"    ' configuration parameter

                            If LCase(Mid(line, 1, 10)) = "{database=" Then
                                sDataBase = Mid(line, 11, 100)
                                sDataBase = sDataBase.Replace("}", "")
                                sConnectString = GetConnectString(sDataBase)
                                If Not CheckAccess() Then
                                    Exit Sub
                                End If
                                WriteLine("Database context changed to " & sDataBase & ".", "I")
                            ElseIf LCase(Mid(line, 1, 5)) = "{sql=" Then
                                s = Mid(line, 6, 1500)
                                s = s.Replace("}", "")
                                sResult = ""
                                If CompileSQL(s) Then
                                    WriteLine("SQL:" & s & " - okay", "I")
                                Else
                                    WriteLine("SQL:" & s & " - errors", "E")
                                    WriteLine(sResult, "E")
                                End If
                            ElseIf LCase(Mid(line, 1, 5)) = "{scl=" Then
                                s = Mid(line, 6, 1500)
                                s = s.Replace("}", "")
                                WriteLine("spawning scl file: " & s, "I")
                                Dim c As New scl
                                AddHandler c.Notify, AddressOf WriteLine
                                c.Process(s, sDataBase)
                                WriteLine("returning from: " & s, "I")
                            End If

                            'version; backup db; etc

                        Case Else
                            If sConnectString = Nothing Then
                                WriteLine("The Database is not defined in 'scl' file.", "E")
                                WriteLine("Aborting compilation", "E")
                                Exit Sub
                            End If

                            If Dir(line) = "" Then
                                line &= " - file not found..."
                                sType = "E"
                            Else
                                If CompileOne(line) Then
                                    line &= " - okay"
                                    sType = "I"
                                Else
                                    line &= " - errors"
                                    sType = "E"
                                End If
                            End If
                            s = Path.Combine(Environment.CurrentDirectory, line)
                            WriteLine(s, sType)
                    End Select
                End If
            Loop Until file.EndOfStream
            file.Close()
        Catch ex As Exception
            WriteLine("CompileList error [" & sList & "]:", "E")
            WriteLine(ex.ToString, "E")
        End Try
        Environment.CurrentDirectory = sPWD
    End Sub

    Private Function CompileOne(ByVal sFile As String) As Boolean
        Dim i As Integer = 0
        Dim s As String = ""
        Dim st As String
        Dim line As String

        Try
            Dim file As New System.IO.StreamReader(sFile)
            Do
                line = file.ReadLine()
                If Len(line) > 2 Then
                    If LCase(Mid(line, 1, 3)) = "go " Then
                        st = LTrim(Mid(line, 3))
                        If st = "" Or Mid(st, 1, 2) = "--" Then    ' trim the space from go statements
                            line = "go"                     ' so the split in CleanSQL will work
                        End If                              ' correctly.
                    End If
                End If
                If line = "go" Then
                    If i = 0 Then
                        line = vbCrLf & "go"
                    End If
                    i = 0
                Else
                    i += 1
                End If
                s &= line & vbCrLf
            Loop Until file.EndOfStream
            file.Close()
        Catch ex As Exception
            WriteLine("CompileOne error [" & sFile & "]:", "E")
            WriteLine(ex.ToString, "E")
            Return False
        End Try

        sResult = ""
        If s <> "" Then
            bErrorLatch = True
            If Not CompileSQL(s) Then
                bErrorLatch = False
            End If
        Else
            bErrorLatch = False
        End If

        If bErrorLatch Then
            DeleteFile(sFile)
            Return True
        Else
            WriteLogFile(sFile)
            Return False
        End If
    End Function

    Private Sub WriteLogFile(ByVal sFile As String)
        Dim s As String
        s = sFile & "." & System.Environment.MachineName & "." & sDataBase & ".log"
        s = s.Replace("\", "~")
        s = s.Replace("/", "~")
        Dim file As New System.IO.StreamWriter(s)
        file.Write(sResult)
        file.Close()
    End Sub

    Private Sub DeleteFile(ByVal sFile As String)
        Dim s As String = sFile & "." & System.Environment.MachineName & "." & sDataBase & ".log"
        s = s.Replace("\", "~")
        s = s.Replace("/", "~")
        Try
            If Dir(s) <> "" Then
                Kill(s)
            End If
        Catch ex As Exception
            If ex.Source <> "" Then
                WriteLine("DeleteFile error [" & s & "]:", "E")
                WriteLine(ex.ToString, "E")
            End If
        End Try
    End Sub

    Private Function CompileSQL(ByVal sText As String) As Boolean
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim sCommands() As String
        Dim b As Boolean = True

        Try
            CompileSQL = False
            psConn = New SqlConnection(sConnectString)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()
            psAdapt = New SqlDataAdapter("", psConn)
            sCommands = Split(sText, vbCrLf & "go" & vbCrLf, , CompareMethod.Text)

            For Each s As String In sCommands
                Try
                    If Trim(s) <> "" Then
                        psAdapt.SelectCommand.CommandText = s
                        psAdapt.SelectCommand.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    If ex.InnerException Is Nothing Then
                        sResult &= ex.Message & vbCrLf
                    Else
                        Dim ex2 As Exception = ex.InnerException
                        Do While Not ex2 Is Nothing
                            sResult &= ex2.Message & vbCrLf
                            ex2 = ex2.InnerException
                        Loop
                    End If
                    b = False
                End Try
            Next

            psConn.Close()
            CompileSQL = b

        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                sResult &= ex.Message & vbCrLf
            Else
                Dim ex2 As Exception = ex.InnerException
                Do While Not ex2 Is Nothing
                    sResult &= ex2.Message & vbCrLf
                    ex2 = ex2.InnerException
                Loop
            End If
        End Try
    End Function

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)

        For Each ex As SqlError In e.Errors
            sResult &= ex.Message & vbCrLf
        Next
    End Sub

    Private Function CheckAccess() As Boolean
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim DS As DataSet
        Dim s As String = ""
        Dim b As Boolean = False

        Try
            psConn = New SqlConnection(sConnectString)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMsg
            psConn.Open()
            psAdapt = New SqlDataAdapter("", psConn)
            psAdapt.SelectCommand.CommandText = "select is_member('db_owner')"
            DS = New DataSet
            psAdapt.Fill(DS, "data")

            If Not DS.Tables("data") Is Nothing Then
                If DS.Tables("data").Rows.Count > 0 Then
                    If DS.Tables("data").Rows(0).Item(0).ToString = "1" Then
                        b = True
                    End If
                End If
            End If
            psConn.Close()
            If Not b Then
                WriteLine("Error, the user is not dbo in target database!", "E")
            End If
            Return b

        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                s &= ex.Message & vbCrLf
            Else
                Dim ex2 As Exception = ex.InnerException
                Do While Not ex2 Is Nothing
                    s &= ex2.Message & vbCrLf
                    ex2 = ex2.InnerException
                Loop
            End If
            WriteLine("CheckAccess error:", "E")
            WriteLine(s, "E")
            Return False
        End Try
    End Function

    Private Sub psConn_InfoMsg(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Dim s As String = ""
        For Each ex As SqlError In e.Errors
            s &= ex.Message & vbCrLf
            If ex.Number <> 0 Then
                bErrorLatch = False
            End If
        Next
        WriteLine("CheckAccess:", "I")
        WriteLine(s, "I")
    End Sub

    Public Function GetConnectString(ByVal sName As String) As String
        Dim settings As System.Configuration.ConnectionStringSettingsCollection = _
            ConfigurationManager.ConnectionStrings

        Try
            If Not settings Is Nothing Then
                Return settings.Item(sName).ConnectionString
            End If

        Catch ex As Exception
            WriteLine("GetConnectString error:", "E")
            WriteLine(ex.ToString, "E")
        End Try

        Return ""
    End Function

    Private Sub WriteLine(ByVal sMsg As String, ByVal sType As String)
        RaiseEvent Notify(sMsg, sType)
    End Sub
End Class
