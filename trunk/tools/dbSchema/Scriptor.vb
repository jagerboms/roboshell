Option Explicit On
Option Strict On

'dbSchema is free software issued as open source;
' you can redistribute it and/or modify it under the terms of the
' GNU General Public License version 2 as published by the Free Software Foundation.
'dbSchema is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with dbSchema;
' if not, go to the web site (http://www.gnu.org/licenses/gpl-2.0.html)
' or write to:
'   The Free Software Foundation, Inc.,
'   59 Temple Place,
'   Suite 330,
'   Boston, MA 02111-1307 USA. 

Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Specialized

Module Scriptor
    Dim Connect As String
    Dim fixdef As Boolean = False
    Dim UniCode As Boolean = False

    Dim Server As String
    Dim UserID As String
    Dim Password As String
    Dim Network As String = ""
    Dim sType As String = ""
    Dim ConsName As Boolean = True
    Dim sObject As String = ""
    Dim mode As Boolean

    Sub main()
        Dim i As Integer
        Dim fv As System.Diagnostics.FileVersionInfo
        Dim Database As String = ""

        Try
            fv = System.Diagnostics.FileVersionInfo.GetVersionInfo( _
                        System.Reflection.Assembly.GetExecutingAssembly.Location)

            Console.WriteLine("dbSchema (version " & fv.FileVersion & ")")
            Console.WriteLine("Copyright 2009 Russell Hansen, Tolbeam Pty Limited")
            Console.WriteLine("")
            Console.WriteLine("dbSchema comes with ABSOLUTELY NO WARRANTY;")
            Console.WriteLine("for details see the -l option.")
            Console.WriteLine("This is free software, and you are welcome")
            Console.WriteLine("to redistribute it under certain conditions")
            Console.WriteLine("described in the GNU General Public License")
            Console.WriteLine("version 2.")

            If GetSwitch("-?") Then
                Console.WriteLine("")
                Console.WriteLine("usage: dbSchema [-sServer] [-dDatabase] [-uUserID [-pPassword]] [-tType] [-oObject] [-f] [-c] [-?] [-l]")
                Console.WriteLine(" where:")
                Console.WriteLine("   -sServer is the name of the SQL server to access.")
                Console.WriteLine("     provided the local machine is used.")
                Console.WriteLine("")
                Console.WriteLine("   -dDatabase is the name of the database to access.")
                Console.WriteLine("     If not provided either the master database or for job types")
                Console.WriteLine("     the msdb database is used.")
                Console.WriteLine("     Use an asterisk * to extract from all the databases on the")
                Console.WriteLine("     Server. The data is extracted into a directory with the database")
                Console.WriteLine("     name. If the directory does not exist it will be created, otherwise")
                Console.WriteLine("     the contents are moved to a backup directory.")
                Console.WriteLine("")
                Console.WriteLine("   -uUserID is the name of the user for database access. If not")
                Console.WriteLine("     provided a Trusted Connection is made.")
                Console.WriteLine("")
                Console.WriteLine("   -pPassword is the user password for database access. This parameter")
                Console.WriteLine("     is ignored except when a UserID is provided.")
                Console.WriteLine("")
                Console.WriteLine("   -tType is the type of object to retrieve. If not provided all")
                Console.WriteLine("     types except jobs and data are returned. Can be one of:")
                Console.WriteLine("      P - stored procedure        U - user table")
                Console.WriteLine("      F - user defined function   V - view")
                Console.WriteLine("      J - job                     D - data")
                Console.WriteLine("")
                Console.WriteLine("   -oObject is the like object name to retrieve. If not provided")
                Console.WriteLine("     all objects are retrieved. This performs a database 'like'")
                Console.WriteLine("     operation so wildcard in the name are supported.")
                Console.WriteLine("     When the type is 'D' the object parameter contains the table")
                Console.WriteLine("     the data to be scripted.")
                Console.WriteLine("")
                Console.WriteLine("   -f full text switch. If provided the scripts are include")
                Console.WriteLine("     existance checks and table components like indexes are")
                Console.WriteLine("     created in separate files.")
                Console.WriteLine("")
                Console.WriteLine("   -c ignore constraint name switch. If provided")
                Console.WriteLine("     names are not included in the generated scripts.")
                Console.WriteLine("")
                Console.WriteLine("   -w where clause filter for data scripting. eg. -w""Status<>'dl'""")
                Console.WriteLine("")
                Console.WriteLine("   -? displays the usage details on the console.")
                Console.WriteLine("")
                Console.WriteLine("   -l displays licence details on the console.")
                Return
            End If

            If GetSwitch("-l") Then
                Console.WriteLine("")
                Console.WriteLine("dbSchema is free software issued as open source;")
                Console.WriteLine("you can redistribute it and/or modify it under the terms")
                Console.WriteLine("of the GNU General Public License version 2 as published")
                Console.WriteLine("by the Free Software Foundation.")
                Console.WriteLine("dbSchema is distributed in the hope that it will be useful,")
                Console.WriteLine("but WITHOUT ANY WARRANTY; without even the implied warranty")
                Console.WriteLine("of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.")
                Console.WriteLine("See the GNU General Public License for more details.")
                Console.WriteLine("You should have received a copy of the GNU General Public")
                Console.WriteLine("License along with dbSchema; if not, go to the web site:")
                Console.WriteLine("")
                Console.WriteLine("   http://www.gnu.org/licenses/gpl-2.0.html")
                Console.WriteLine("")
                Console.WriteLine("or write to:")
                Console.WriteLine("")
                Console.WriteLine("   The Free Software Foundation, Inc.,")
                Console.WriteLine("   59 Temple Place,")
                Console.WriteLine("   Suite 330,")
                Console.WriteLine("   Boston, MA 02111-1307 USA.")
                Return
            End If

            Server = GetCommandParameter("-s")
            If Server = "" Then Server = System.Environment.MachineName
            Database = GetCommandParameter("-d")
            UserID = GetCommandParameter("-u")
            Password = GetCommandParameter("-p")
            Network = GetCommandParameter("-n")
            mode = GetSwitch("-f")
            fixdef = GetSwitch("-z")
            UniCode = GetSwitch("-y")
            sType = GetCommandParameter("-t")
            sObject = GetCommandParameter("-o")
            If GetSwitch("-c") Then
                ConsName = False
            End If

            If Mid(LCase(sType), 1, 1) = "j" Then
                ProcessJobs(Database)
            ElseIf Mid(LCase(sType), 1, 1) = "d" Then
                ProcessData(Database, sObject)
            ElseIf Database = "*" Then
                ProcessAllDBs()
            Else
                If Database = "" Then Database = "master"
                ProcessDB(Database)
            End If

        Catch ex As Exception
            i = 1
            Console.WriteLine(ex.ToString)
        End Try
        If i <> 0 Then
            Console.WriteLine("press enter")
            Console.Read()
        End If
    End Sub

    Sub ProcessJobs(ByVal Database As String)
        Dim s As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        Try                                 ' Read the config XML into a DataSet
            If Database = "" Or Database = "*" Then
                Database = "msdb"
            End If
            Connect = GetConnectString(Database)
            psConn = New SqlConnection(Connect)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()

            s = "select job_id"
            s &= " from dbo.sysjobs"
            If sObject <> "" Then
                s &= " where name like '" & sObject & "'"
            End If
            s &= " order by name"

            psAdapt = New SqlDataAdapter(s, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.Text
            psAdapt.Fill(Details)
            psConn.Close()

            For Each dr In Details.Tables(0).Rows
                s = GetString(dr.Item("job_id"))
                GetJob(s, Connect, mode)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Function GetJob(ByVal sJobID As String, ByVal sConnect As String, ByVal mode As Boolean) As Integer
        Dim js As New Job(sJobID, sConnect)
        Dim sOut As String = ""
        Dim s As String

        s = js.JobName
        If s = "" Then Return -1

        s = Replace(s, ":", "_")
        s = Replace(s, "\", "_")
        s = Replace(s, "/", "_")
        If mode Then
            sOut = js.FullText
        Else
            sOut = js.CommonText
        End If
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf

        PutFile("job." & s & ".sql", sOut)
        Return 0
    End Function

    Sub ProcessData(ByVal Database As String, ByVal Table As String)
        Dim sOut As String = ""
        Dim s As String

        Try
            Connect = GetConnectString(Database)
            s = GetCommandParameter("-w")
            Dim tdefn As New TableColumns(Table, Connect, True)
            sOut = tdefn.DataScript(s)
            sOut &= vbCrLf & "go" & vbCrLf
            PutFile("data." & Table & ".sql", sOut)

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Sub ProcessAllDBs()
        Dim s As String
        Dim sBack As String
        Dim sDir As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow
        Dim sPWD As String

        sPWD = Environment.CurrentDirectory
        Try                                 ' Read the config XML into a DataSet
            Connect = GetConnectString("master")
            psConn = New SqlConnection(Connect)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()

            s = "select name from sysdatabases where name not in ('master','tempdb','model','msdb')"
            psAdapt = New SqlDataAdapter(s, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.Text
            psAdapt.Fill(Details)
            psConn.Close()

            For Each dr In Details.Tables(0).Rows
                s = GetString(dr.Item("name"))

                If System.IO.Directory.Exists(s) Then
                    Environment.CurrentDirectory = s
                    sBack = "Backup" & Format(Now(), "yyyyMMddHHmmss")
                    MkDir(sBack)
                    sDir = Dir("*")
                    Do While sDir <> ""
                        System.IO.File.Move(sDir, System.IO.Path.Combine(sBack, sDir))
                        sDir = Dir()
                    Loop
                Else
                    MkDir(s)
                    Environment.CurrentDirectory = s
                End If
                ProcessDB(s)
                Environment.CurrentDirectory = sPWD
            Next

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Sub ProcessDB(ByVal Database As String)
        Dim s As String
        Dim st As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        Try                                 ' Read the config XML into a DataSet
            Connect = GetConnectString(Database)
            psConn = New SqlConnection(Connect)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()

            s = "select type, name "
            s &= "from dbo.sysobjects "
            Select Case sType
                Case "P", "U", "V"
                    s &= "where type = '" & sType & "' "
                Case "F"
                    s &= "where type in ('FN', 'TF') "
                Case Else
                    s &= "where type in ('P', 'U', 'V', 'FN', 'TF') "
            End Select
            If sObject <> "" Then
                s &= "and name like '" & sObject & "' "
            End If
            s &= "and uid = 1 "
            s &= "and name not like 'dt_%' "
            s &= "and name not in ('syssegments','sysconstraints',"
            s &= "'sp_alterdiagram','sp_creatediagram',"
            s &= "'sp_dropdiagram','sp_helpdiagramdefinition','sp_helpdiagrams',"
            s &= "'sp_renamediagram','sp_upgraddiagrams','sysdiagrams','fn_diagramobjects') "
            s &= "order by type, name"

            psAdapt = New SqlDataAdapter(s, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.Text
            psAdapt.Fill(Details)
            psConn.Close()

            For Each dr In Details.Tables(0).Rows
                s = GetString(dr.Item("name"))
                st = GetString(dr.Item("type"))
                If GetString(dr.Item("type")) = "U" Then
                    If mode Then
                        GetTableFull(s, Connect, ConsName)
                    Else
                        GetTable(s, Connect, ConsName)
                    End If
                Else
                    GetText(s, st)
                End If
            Next

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Function GetTable(ByVal sTable As String, ByVal sConnect As String, ByVal ConsName As Boolean) As Integer
        Dim ts As New TableColumns(sTable, sConnect, fixdef)
        Dim sOut As String
        Dim s As String

        If Not ConsName Then ts.ScriptConstraints = False
        sOut = ts.TableText
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf

        For Each s In ts.IKeys
            If s <> "" Then
                sOut &= ts.IndexShort(s)
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            End If
        Next

        For Each s In ts.FKeys
            If s <> "" Then
                sOut &= ts.FKeyShort(s)
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            End If
        Next

        PutFile("table." & sTable & ".sql", sOut)
        Return 0
    End Function

    Private Function GetTableFull(ByVal sTable As String, ByVal sConnect As String, ByVal ConsName As Boolean) As Integer
        Dim ts As New TableColumns(sTable, sConnect, fixdef)
        Dim sOut As String
        Dim s As String

        If Not ConsName Then ts.ScriptConstraints = False
        sOut = ts.FullTableText
        sOut &= "go" & vbCrLf
        PutFile("table." & sTable & ".sql", sOut)

        For Each s In ts.IKeys
            If s <> "" Then
                sOut = ts.IndexText(s)
                sOut &= "go" & vbCrLf
                PutFile("index." & sTable & "." & s & ".sql", sOut)
            End If
        Next

        For Each s In ts.FKeys
            If s <> "" Then
                sOut = ts.FKeyText(s)
                sOut &= "go" & vbCrLf
                PutFile("fkey." & sTable & "." & s & ".sql", sOut)
            End If
        Next

        Return 0
    End Function

    Private Function GetText(ByVal Name As String, ByVal Type As String) As Integer
        Dim sText As String
        Dim Pre As String = ""

        sText = GetdbText(Name, Type)

        Select Case Type
            Case "P"
                Pre = "proc."
            Case "U"
                Pre = "table."
            Case "V"
                Pre = "view."
            Case "FN", "TF"
                Pre = "udf."
        End Select

        PutFile(Pre & Name & ".sql", sText)
        Return 0
    End Function

#Region "common functions"
    Private Function GetdbText(ByVal Name As String, ByVal Type As String) As String
        Dim s As String
        Dim sText As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow
        Dim b As Boolean = True

        psConn = New SqlConnection(Connect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        s = "select text from syscomments "
        s &= "where id = object_id('" & Name & "') order by number, colid"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text

        'psAdapt = New SqlDataAdapter("dbo.sp_helptext", psConn)
        'psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
        'psAdapt.SelectCommand.Parameters.Add("@objname", SqlDbType.NVarChar, 776).Value = """" & Name & """"
        psAdapt.Fill(Details)
        psConn.Close()

        sText = ""
        If Type <> "P" And Type <> "FN" And Type <> "TF" Then
            b = False
        End If

        For Each dr In Details.Tables(0).Rows        ' Columns
            s = CType(dr.Item("Text"), String)
            If Len(s) < 4000 Then s &= vbCrLf
            sText &= s ' & vbCrLf
        Next
        sText = sText.Replace(vbCrLf, Chr(13))
        sText = sText.Replace(Chr(10), Chr(13))
        sText = sText.Replace(Chr(13), vbCrLf)
        Do While 1 = 1
            Select Case Mid(sText, 1, 1)
                Case " ", Chr(9), Chr(10), Chr(13)
                    sText = Mid(sText, 2)
                Case Else
                    Exit Do
            End Select
        Loop

        Do While 1 = 1
            Select Case Right(sText, 1)
                Case " ", Chr(9), Chr(10), Chr(13)
                    sText = Mid(sText, 1, Len(sText) - 1)
                Case Else
                    Exit Do
            End Select
        Loop
        sText &= vbCrLf

        Return sText
    End Function

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Console.WriteLine(e.Message)
    End Sub

    Function PutFile(ByVal sName As String, ByVal sContent As String) As Boolean
        Dim file As System.IO.StreamWriter

        If UniCode Then
            file = New System.IO.StreamWriter(sName, False, System.Text.Encoding.Unicode)
        Else
            file = New System.IO.StreamWriter(sName)
        End If

        file.Write(sContent)
        file.Close()
        Console.WriteLine(sName)
        Return True
    End Function

    Private Function GetSwitch(ByRef sSwitch As String) As Boolean
        Dim sCommand As String

        sCommand = Microsoft.VisualBasic.Command()
        If InStr(1, sCommand, sSwitch, CompareMethod.Text) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetCommandParameter(ByRef sSwitch As String) As String
        Dim sCommand As String
        Dim sParameter As String
        Dim i As Integer

        sCommand = Microsoft.VisualBasic.Command()
        i = InStr(1, sCommand, sSwitch, CompareMethod.Text)
        sParameter = ""
        If i > 0 Then
            sParameter = Mid(sCommand, i + 2)
            If Mid(sParameter, 1, 1) = """" Then
                sParameter = Mid(sParameter, 2)
                i = InStr(1, sParameter, """", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            Else
                i = InStr(1, sParameter, " ", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            End If
        End If
        GetCommandParameter = sParameter
    End Function

    Public Function GetString(ByVal objValue As Object) As String
        If IsDBNull(objValue) Then
            Return ""
        ElseIf objValue Is Nothing Then
            Return ""
        Else
            Try
                Return CType(objValue, String).TrimEnd
            Catch ex As Exception
                Return objValue.ToString
            End Try
        End If
    End Function

    Private Function GetTriggers(ByVal sTableName As String) As DataTable
        Dim s As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim TableDetails As New DataSet

        psConn = New SqlConnection(Connect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        s = "select o.name TriggerName "
        s &= "from dbo.sysobjects o "
        s &= "where o.type = 'TR' "
        s &= "and o.parent_obj = object_id('" & sTableName & "')"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(TableDetails)
        psConn.Close()

        Return TableDetails.Tables(0)
    End Function

    Private Function GetConnectString(ByVal sDatabase As String) As String
        Dim s As String

        s = "server=" & Server & ";database=" & sDatabase
        If Network <> "" Then s &= ";Network=" & Network
        If UserID <> "" Then
            s &= ";User ID=" & UserID
            If Password <> "" Then
                s &= ";pwd=" & Password
            End If
        Else
            s &= ";Integrated Security=SSPI"
        End If
        Return s
    End Function
#End Region
End Module
