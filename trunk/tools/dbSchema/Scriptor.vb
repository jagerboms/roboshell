Option Explicit On
Option Strict On

#Region "copyright Russell Hansen, Tolbeam Pty Limited"
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
#End Region

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
    Dim LogFile As String = ""
    Dim verbose As Boolean = False

    Sub main()
        Dim fv As System.Diagnostics.FileVersionInfo
        Dim Database As String = ""
        Dim s As String

        Try
            LogFile = GetCommandParameter("-g")
            verbose = GetSwitch("-v")
            If LogFile <> "" Then
                LogFile = System.IO.Path.Combine(Environment.CurrentDirectory, LogFile)
            End If
            fv = System.Diagnostics.FileVersionInfo.GetVersionInfo( _
                        System.Reflection.Assembly.GetExecutingAssembly.Location)

            SendMessage("dbSchema (version " & fv.FileVersion & ")", "H")
            SendMessage("Copyright 2009 Russell Hansen, Tolbeam Pty Limited", "T")
            SendMessage("", "T")
            SendMessage("dbSchema comes with ABSOLUTELY NO WARRANTY;", "N")
            SendMessage("for details see the -l option.", "N")
            SendMessage("This is free software, and you are welcome", "N")
            SendMessage("to redistribute it under certain conditions", "N")
            SendMessage("described in the GNU General Public License", "N")
            SendMessage("version 2.", "N")
            SendMessage("", "N")

            If GetSwitch("-?") Then
                SendMessage("usage: dbSchema [-sServer] [-dDatabase] [-uUserID [-pPassword]] [-tType]", "T")
                SendMessage("                [-oObject] [-f] [-c] [-?] [-l] [-gLogFile] [-v]", "T")
                SendMessage(" where:", "T")
                SendMessage("   -sServer is the name of the SQL server to access.", "T")
                SendMessage("     provided the local machine is used.", "T")
                SendMessage("", "T")
                SendMessage("   -dDatabase is the name of the database to access.", "T")
                SendMessage("     If not provided either the master database or for job types", "T")
                SendMessage("     the msdb database is used.", "T")
                SendMessage("     Use an asterisk * to extract from all the databases on the", "T")
                SendMessage("     Server (except master, model and tempdb). Only job scripts are", "T")
                SendMessage("     extracted from the msdb database (i.e. no table, procedures etc.).", "T")
                SendMessage("     The data is extracted into a directory with the database", "T")
                SendMessage("     name. If the directory does not exist it will be created, otherwise", "T")
                SendMessage("     the contents are moved to a backup directory.", "T")
                SendMessage("", "T")
                SendMessage("   -uUserID is the name of the user for database access. If not", "T")
                SendMessage("     provided a Trusted Connection is made.", "T")
                SendMessage("", "T")
                SendMessage("   -pPassword is the user password for database access. This parameter", "T")
                SendMessage("     is ignored except when a UserID is provided.", "T")
                SendMessage("", "T")
                SendMessage("   -tType is the type of object to retrieve. If not provided all", "T")
                SendMessage("     types except jobs and data are returned. Can be one of:", "T")
                SendMessage("      P - stored procedure        U - user table", "T")
                SendMessage("      F - user defined function   V - view", "T")
                SendMessage("      J - job                     D - data", "T")
                SendMessage("      S - script permissions", "T")
                SendMessage("", "T")
                SendMessage("   -oObject is the like object name to retrieve. If not provided", "T")
                SendMessage("     all objects are retrieved. This performs a database 'like'", "T")
                SendMessage("     operation so wildcard in the name are supported.", "T")
                SendMessage("     When the type is 'D' the object parameter contains the table", "T")
                SendMessage("     the data to be scripted.", "T")
                SendMessage("     When the type is 'S' the object parameter contains the user", "T")
                SendMessage("     the permissions are to be scripted.", "T")
                SendMessage("", "T")
                SendMessage("   -f full text switch. If provided the scripts are include", "T")
                SendMessage("     existance checks and table components like indexes are", "T")
                SendMessage("     created in separate files.", "T")
                SendMessage("", "T")
                SendMessage("   -c ignore constraint name switch. If provided", "T")
                SendMessage("     names are not included in the generated scripts.", "T")
                SendMessage("", "T")
                SendMessage("   -w where clause filter for data scripting. eg. -w""Status<>'dl'""", "T")
                SendMessage("", "T")
                SendMessage("   -? displays the usage details on the console.", "T")
                SendMessage("", "T")
                SendMessage("   -l displays licence details on the console.", "T")
                SendMessage("", "T")
                SendMessage("   -gLogFile defines the file where screen output is saved.", "T")
                SendMessage("", "T")
                SendMessage("   -v verbose output switch. If set extended output is produced.", "T")
                SendMessage("", "T")
                Return
            End If

            If GetSwitch("-l") Then
                SendMessage("dbSchema is free software issued as open source;", "T")
                SendMessage("you can redistribute it and/or modify it under the terms", "T")
                SendMessage("of the GNU General Public License version 2 as published", "T")
                SendMessage("by the Free Software Foundation.", "T")
                SendMessage("dbSchema is distributed in the hope that it will be useful,", "T")
                SendMessage("but WITHOUT ANY WARRANTY; without even the implied warranty", "T")
                SendMessage("of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.", "T")
                SendMessage("See the GNU General Public License for more details.", "T")
                SendMessage("You should have received a copy of the GNU General Public", "T")
                SendMessage("License along with dbSchema; if not, go to the web site:", "T")
                SendMessage("", "T")
                SendMessage("   http://www.gnu.org/licenses/gpl-2.0.html", "T")
                SendMessage("", "T")
                SendMessage("or write to:", "T")
                SendMessage("", "T")
                SendMessage("   The Free Software Foundation, Inc.,", "T")
                SendMessage("   59 Temple Place,", "T")
                SendMessage("   Suite 330,", "T")
                SendMessage("   Boston, MA 02111-1307 USA.", "T")
                SendMessage("", "T")
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

            If LogFile <> "" Then
                SendMessage("Machine Name : " & Environment.MachineName, "T")
                SendMessage("Directory    : " & Environment.CurrentDirectory, "T")
                s = Environment.CommandLine
                If Password <> "" Then
                    s = Replace(s, " -p" & Password & " ", " -p?????????? ")
                End If
                SendMessage("Command Line : " & s, "T")
            End If

            If Mid(LCase(sType), 1, 1) = "j" Then
                ProcessJobs(Database)
            ElseIf Mid(LCase(sType), 1, 1) = "d" Then
                ProcessData(Database, sObject)
            ElseIf Mid(LCase(sType), 1, 1) = "s" Then
                ProcessPermissions(sObject)
            ElseIf Database = "*" Then
                ProcessAllDBs()
            Else
                If Database = "" Then Database = "master"
                ProcessDB(Database)
            End If

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
        SendMessage("", "T")
        SendMessage("", "N")
    End Sub

    Sub ProcessJobs(ByVal Database As String)
        Dim s As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        Try                                 ' Read the config XML into a DataSet
            SendMessage("", "N")
            If Database = "" Or Database = "*" Then
                Connect = GetConnectString("msdb")
                SendMessage("Retrieving jobs from 'msdb'.", "T")
            Else
                Connect = GetConnectString(Database)
                SendMessage("Retrieving jobs from '" & Database & "'.", "T")
            End If
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
            SendMessage(ex.ToString, "E")
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

    Sub ProcessPermissions(ByVal User As String)
        Dim s As String
        Dim sOut As String = ""
        Dim sUser As String = ""
        Dim b As Boolean
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        Try
            Connect = GetConnectString("master")
            psConn = New SqlConnection(Connect)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()

            s = "select name from master.dbo.sysdatabases where name not in ('master','tempdb','model')"
            s &= " select l.sid,u.name from sys.syslogins l left join sys.sysusers u on u.sid = l.sid where l.name = '" & User & "'"
            psAdapt = New SqlDataAdapter(s, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.Text
            psAdapt.Fill(Details)
            psConn.Close()

            If Details.Tables.Count > 1 Then
                If Details.Tables(1).Rows.Count > 0 Then
                    dr = Details.Tables(1).Rows(0)
                    sUser = GetString(dr("name"))
                End If
                b = False
            Else
                b = True
            End If

            sOut = "use master" & vbCrLf
            If b Or sUser = "" Then
                sOut &= "create user " & User & " for login " & User & vbCrLf
                sUser = User
            End If

            sOut &= "grant select on dbo.sysdatabases to " & sUser & vbCrLf
            sOut &= "grant select on sys.servers to " & sUser & vbCrLf
            sOut &= "grant select on INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS to " & sUser & vbCrLf
            sOut &= "grant select on INFORMATION_SCHEMA.KEY_COLUMN_USAGE to " & sUser & vbCrLf
            sOut &= "grant select on INFORMATION_SCHEMA.COLUMNS to " & sUser & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "use msdb" & vbCrLf
            If Not b Then
                sUser = GetUID("msdb", User)
            End If
            If b Or sUser = "" Then
                sOut &= "create user " & User & " for login " & User & vbCrLf
                sUser = User
            End If
            sOut &= "grant select on dbo.sysjobs to " & sUser & vbCrLf
            sOut &= "grant select on dbo.syscategories to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysoperators to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysjobservers to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysjobsteps to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysproxies to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysjobschedules to " & sUser & vbCrLf
            sOut &= "grant select on dbo.sysschedules to " & sUser & vbCrLf
            sOut &= "go" & vbCrLf

            For Each dr In Details.Tables(0).Rows
                s = GetString(dr.Item("name"))

                sOut &= vbCrLf
                sOut &= "use " & s & vbCrLf
                If Not b Then
                    sUser = GetUID(s, User)
                End If
                If b Or sUser = "" Then
                    sOut &= "create user " & User & " for login " & User & vbCrLf
                    sUser = User
                End If
                sOut &= "grant select on dbo.syscolumns to " & sUser & vbCrLf
                sOut &= "grant select on dbo.sysobjects to " & sUser & vbCrLf
                sOut &= "grant select on dbo.syscomments to " & sUser & vbCrLf
                sOut &= "grant select on sys.indexes to " & sUser & vbCrLf
                sOut &= "grant select on sys.index_columns to " & sUser & vbCrLf
                sOut &= "grant select on sys.columns to " & sUser & vbCrLf
                sOut &= "grant select on sys.sql_modules to " & sUser & vbCrLf
                sOut &= "go" & vbCrLf
            Next
            PutFile("script.dbSchema-Permission.sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Private Function GetUID(ByVal sDatabase As String, ByVal User As String) As String
        Dim sUser As String = ""
        Dim s As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet

        Connect = GetConnectString(sDatabase)
        psConn = New SqlConnection(Connect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        s = " select u.name from sys.syslogins l join sys.sysusers u on u.sid = l.sid where l.name = '" & User & "'"
        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(Details)
        psConn.Close()

        If Details.Tables.Count > 0 Then
            If Details.Tables(0).Rows.Count > 0 Then
                sUser = GetString(Details.Tables(0).Rows(0)("name"))
            End If
        End If

        Return sUser
    End Function

    Sub ProcessData(ByVal Database As String, ByVal Table As String)
        Dim sOut As String = ""
        Dim s As String

        Try
            Connect = GetConnectString(Database)
            s = GetCommandParameter("-w")
            Dim tdefn As New TableColumns(Table, Connect, True)
            sOut = tdefn.DataScript(s)
            PutFile("data." & Table & ".sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Sub ProcessAllDBs()
        Dim s As String
        Dim dbVersion As Integer
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

            s = "select name,cmptlevel from master.dbo.sysdatabases where name not in ('master','tempdb','model')"
            psAdapt = New SqlDataAdapter(s, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.Text
            psAdapt.Fill(Details)
            psConn.Close()

            For Each dr In Details.Tables(0).Rows
                s = GetString(dr.Item("name"))
                dbVersion = CInt(dr("cmptlevel"))

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
                If s = "msdb" And dbVersion > 80 Then
                    ProcessJobs(s)
                Else
                    ProcessDB(s)
                End If
                Environment.CurrentDirectory = sPWD
            Next

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Sub ProcessDB(ByVal Database As String)
        Dim s As String
        Dim st As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        Try
            SendMessage("", "N")
            SendMessage("Retrieving schema for '" & Database & "'.", "T")
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
            SendMessage(ex.ToString, "E")
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
        Dim sHead As String
        Dim Settings As String = ""
        Dim Pre As String = ""

        sText = GetdbText(Name, Type)

        sHead = "if object_id('dbo." & Name & "') is not null" & vbCrLf
        sHead &= "begin" & vbCrLf
        sHead &= "    drop "

        Select Case Type
            Case "P"
                Pre = "proc."
                sHead &= "procedure"
                If mode Then
                    Settings = GetSetings(Name)
                End If
            Case "V"
                Pre = "view."
                sHead &= "view"
            Case "FN", "TF"
                Pre = "udf."
                sHead &= "function"
                If mode Then
                    Settings = GetSetings(Name)
                End If
        End Select

        sHead &= " dbo." & Name & vbCrLf
        sHead &= "end" & vbCrLf
        sHead &= "go" & vbCrLf
        sHead &= Settings

        If mode Then
            sText = sHead & sText
            sText &= "go" & vbCrLf
        End If

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

    Private Function GetSetings(ByVal Name As String) As String
        Dim s As String
        Dim sText As String = ""
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        psConn = New SqlConnection(Connect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        s = "select uses_ansi_nulls,uses_quoted_identifier from sys.sql_modules "
        s &= "where object_id=object_id('" & Name & "')"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text

        psAdapt.Fill(Details)
        psConn.Close()

        If Details.Tables.Count > 0 Then
            If Details.Tables(0).Rows.Count > 0 Then
                dr = Details.Tables(0).Rows(0)

                If CInt(dr("uses_ansi_nulls")) = 0 Then
                    sText &= "set ansi_nulls off" & vbCrLf
                End If
                If CInt(dr("uses_quoted_identifier")) = 0 Then
                    sText &= "set quoted_identifier off" & vbCrLf
                End If

                If sText <> "" Then
                    sText &= "go" & vbCrLf
                End If
            End If
        End If

        Return sText
    End Function

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        SendMessage(e.Message, "N")
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
        SendMessage(sName, "N")
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

    Private Sub SendMessage(ByVal sMessage As String, ByVal sType As String)
        If Not verbose And sType = "N" Then Return
        Console.WriteLine(sMessage)
        If LogFile <> "" Then
            Dim file As New System.IO.StreamWriter(LogFile, True)
            If sType = "H" Then
                file.WriteLine("Run Time: " & Now())
            End If
            file.WriteLine(sMessage)
            file.Close()
        End If
    End Sub
#End Region
End Module
