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
    Dim UniCode As Boolean = False

    Dim sType As String = ""
    Dim IncludePerm As Boolean = False
    Dim JobSQL As Boolean = False
    Dim sObject As String = ""
    Dim sSchema As String = ""
    Dim mode As String = "S"
    Dim bXML As Boolean = False
    Dim LogFile As String = ""
    Dim verbose As Boolean = False
    Dim opt As New ScriptOptions

    Dim sqllib As New sql

    Public Sub main()
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

            If GetSwitch("-?") Or Trim(Microsoft.VisualBasic.Command()) = "" Then
                SendMessage("usage: dbSchema [-sServer] [-dDatabase] [-uUserID [-pPassword]] [-iTimeOut]", "T")
                SendMessage("                [-tType] [-oObject] [-f] [-c] [-m] [-?] [-l] [-gLogFile] [-v]", "T")
                SendMessage(" where:", "T")
                SendMessage("   -sServer is the name of the SQL server to access.", "T")
                SendMessage("     provided the local machine is used.", "T")
                SendMessage("", "T")
                SendMessage("   -dDatabase is the name of the database to access.", "T")
                SendMessage("     If not provided either the master database or for job types", "T")
                SendMessage("     the msdb database is used.", "T")
                SendMessage("     Use an asterisk * to extract from all the databases on the", "T")
                SendMessage("     Server (except master, model, tempdb and distribution). Only job", "T")
                SendMessage("     scripts are extracted from the msdb database (i.e. no table,", "T")
                SendMessage("     procedures etc.). The data is extracted into a directory with", "T")
                SendMessage("     the database name. If the directory does not exist it will be", "T")
                SendMessage("     created, otherwise the contents are moved to a backup directory.", "T")
                SendMessage("", "T")
                SendMessage("   -uUserID is the name of the user for database access. If not", "T")
                SendMessage("     provided a Trusted Connection is made.", "T")
                SendMessage("", "T")
                SendMessage("   -pPassword is the user password for database access. This parameter", "T")
                SendMessage("     is ignored except when a UserID is provided.", "T")
                SendMessage("", "T")
                SendMessage("   -iTimeOut is the timeout in seconds used when accessing the database.", "T")
                SendMessage("", "T")
                SendMessage("   -tType is the type of object to retrieve. If not provided all", "T")
                SendMessage("     types except jobs, script permissions and data are returned.", "T")
                SendMessage("     Can be one of:", "T")
                SendMessage("      P - stored procedure", "T")
                SendMessage("      U - user table", "T")
                SendMessage("      F - user defined function", "T")
                SendMessage("      V - view", "T")
                SendMessage("      T - trigger", "T")
                SendMessage("      J - job", "T")
                SendMessage("      D - data", "T")
                SendMessage("      S - script permissions", "T")
                SendMessage("     Stored procedures, tables, functions, views and triggers types", "T")
                SendMessage("     can be combined.", "T")
                SendMessage("", "T")
                SendMessage("   -oObject is the like object name to retrieve. If not provided", "T")
                SendMessage("     all objects are retrieved. This performs a database 'like'", "T")
                SendMessage("     operation so wildcard in the name are supported.", "T")
                SendMessage("     When the type is 'D' the object parameter contains the table", "T")
                SendMessage("     the data to be scripted.", "T")
                SendMessage("     When the type is 'S' the object parameter contains the user", "T")
                SendMessage("     the permissions are to be scripted.", "T")
                SendMessage("", "T")
                SendMessage("   -hSchema is schema to retrieve. If not provided", "T")
                SendMessage("     objects from all schemas are retrieved.", "T")
                SendMessage("", "T")
                SendMessage("   -fType determines the type of script to be generated. Can be one of:", "T")
                SendMessage("      F - full includes existance checks and separate component files", "T")
                SendMessage("      X - table components in XML existance checks for code files", "T")
                SendMessage("      I - intermediate has no existance checks but separate component files", "T")
                SendMessage("      S - summary has no existance checks and all components are in a single", "T")
                SendMessage("          file.", "T")
                SendMessage("", "T")
                SendMessage("   -c ignore constraint name switch. If provided, constraint names are not", "T")
                SendMessage("      included in the generated scripts.", "T")
                SendMessage("", "T")
                SendMessage("   -a switch to include column collation in table scripts. If not provided,", "T")
                SendMessage("      column collations are not included in the generated scripts.", "T")
                SendMessage("", "T")
                SendMessage("   -m include permissions switch. If provided, permissions are included along", "T")
                SendMessage("      with the job creation scripts.", "T")
                SendMessage("", "T")
                SendMessage("   -j extract job scrips switch. If provided, job step sql scripts are extracted", "T")
                SendMessage("      otherwise job creation scripts are generated.", "T")
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

            s = GetCommandParameter("-s")
            If s = "" Then s = System.Environment.MachineName
            sqllib.Server = s
            Database = GetCommandParameter("-d")
            sqllib.UserID = GetCommandParameter("-u")
            sqllib.Password = GetCommandParameter("-p")
            s = GetCommandParameter("-i")
            If IsNumeric(s) Then
                sqllib.TimeOut = CInt(s)
            End If
            sqllib.Network = GetCommandParameter("-n")
            s = GetCommandParameter("-f")
            Select Case UCase(Mid(s, 1, 1))
                Case "X"
                    mode = "F"
                    bXML = True
                Case "F"
                    mode = "F"
                Case "I"
                    mode = "I"
            End Select
            opt.DefaultFix = GetSwitch("-z")
            UniCode = GetSwitch("-y")
            sType = GetCommandParameter("-t")
            sObject = GetCommandParameter("-o")
            sSchema = GetCommandParameter("-h")
            If GetSwitch("-a") Then
                opt.CollationShow = True
            End If
            If GetSwitch("-c") Then
                opt.DefaultShowName = False
            End If

            If GetSwitch("-m") Then
                IncludePerm = True
            End If

            If GetSwitch("-j") Then
                JobSQL = True
            End If

            If LogFile <> "" Then
                SendMessage("Machine Name : " & Environment.MachineName, "T")
                SendMessage("Directory    : " & Environment.CurrentDirectory, "T")
                s = Environment.CommandLine
                If sqllib.Password <> "" Then
                    s = Replace(s, sqllib.Password, "??????????")
                End If
                SendMessage("Command Line : " & s, "T")
            End If

            If Mid(LCase(sType), 1, 1) = "j" Then
                ProcessJobs(Database)
            ElseIf Mid(LCase(sType), 1, 1) = "d" Then
                ProcessData(Database, sObject, "dbo")
            ElseIf Mid(LCase(sType), 1, 1) = "s" Then
                ProcessPermissions(sObject)
            ElseIf Database = "*" Then
                ProcessAllDB()
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

    Private Function ProcessJobs(ByVal Database As String) As Boolean
        Dim s As String
        Dim dt As DataTable
        Dim dr As DataRow

        Try                                 ' Read the config XML into a DataSet
            SendMessage("", "N")
            If Database = "" Or Database = "*" Then
                sqllib.Database = "msdb"
                SendMessage("Retrieving jobs from 'msdb'.", "T")
            Else
                sqllib.Database = Database
                SendMessage("Retrieving jobs from '" & Database & "'.", "T")
            End If

            dt = sqllib.JobList(sObject)
            For Each dr In dt.Rows
                s = sqllib.GetString(dr.Item("job_id"))
                GetJob(s, mode)
            Next

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
            Return False
        End Try

        Return True
    End Function

    Private Sub ProcessPermissions(ByVal User As String)
        Dim sDB As String
        Dim s As String
        Dim sOut As String = ""
        Dim sUser As String
        Dim dd As DataTable
        Dim dr As DataRow

        Try
            sqllib.Database = "master"
            dd = sqllib.DatabaseList()
            sUser = sqllib.UserCreate(User)

            sOut = "use master" & vbCrLf
            sOut &= sUser & vbCrLf
            sOut &= "go" & vbCrLf
            s = sqllib.UserGrant("master", User)
            If s <> "" Then
                sOut &= s & vbCrLf
                sOut &= "go" & vbCrLf
            End If

            For Each dr In dd.Rows
                sDB = sqllib.GetString(dr.Item("name"))
                sOut &= vbCrLf
                sOut &= "use " & sDB & vbCrLf
                sOut &= sUser & vbCrLf
                sOut &= "go" & vbCrLf
                s = sqllib.UserGrant(sDB, User)
                If s <> "" Then
                    sOut &= s & vbCrLf
                    sOut &= "go" & vbCrLf
                End If
            Next
            WriteFile("script", "", "dbSchema-Permission", "", "sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Private Sub ProcessData(ByVal Database As String, ByVal Table As String, ByVal Schema As String)
        Dim sOut As String = ""
        Dim s As String

        Try
            sqllib.Database = Database
            s = GetCommandParameter("-w")
            Dim cData As New Data(Table, Schema, sqllib)
            sOut = cData.DataScript(s)

            WriteFile("data", Schema, Table, "", "sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Private Sub ProcessAllDBs()
        Dim s As String
        Dim sBack As String
        Dim sDir As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sPWD As String

        sPWD = Environment.CurrentDirectory
        Try                                 ' Read the config XML into a DataSet
            sqllib.Database = "master"
            dt = sqllib.DatabaseList()

            For Each dr In dt.Rows
                s = sqllib.GetString(dr.Item("name"))

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
                If s = "msdb" Then
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

    Private Sub ProcessAllDB()
        Dim s As String
        Dim OK As Boolean
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sPWD As String

        sPWD = Environment.CurrentDirectory
        Try                                 ' Read the config XML into a DataSet
            sqllib.Database = "master"
            dt = sqllib.DatabaseList()

            For Each dr In dt.Rows
                s = sqllib.GetString(dr.Item("name"))

                If System.IO.Directory.Exists(s) Then
                    Environment.CurrentDirectory = s
                    System.IO.File.Delete("*.sql")
                Else
                    MkDir(s)
                    Environment.CurrentDirectory = s
                End If

                If s = "msdb" Then
                    OK = ProcessJobs(s)
                Else
                    OK = ProcessDB(s)
                End If

                If Not OK Then
                    SendMessage(Environment.CurrentDirectory, "I")
                    Environment.CurrentDirectory = sPWD
                    SendMessage("Removing " & s, "I")
                    RmDir(s)
                Else
                    Environment.CurrentDirectory = sPWD
                End If
            Next

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
        Environment.CurrentDirectory = sPWD
    End Sub

    Private Function ProcessDB(ByVal Database As String) As Boolean
        Dim s As String
        Dim st As String
        Dim ss As String
        Dim dt As DataTable
        Dim dr As DataRow

        Try
            SendMessage("", "N")
            SendMessage("Retrieving schema for '" & Database & "'.", "T")
            sqllib.Database = Database
            dt = sqllib.DatabaseObject(sObject, sSchema, sType)

            For Each dr In dt.Rows
                s = sqllib.GetString(dr.Item("name"))
                st = sqllib.GetString(dr.Item("type"))
                ss = sqllib.GetString(dr.Item("sch"))
                If sqllib.GetString(dr.Item("type")) = "U" Then
                    If bXML Then
                        GetTableXML(s, ss)
                    Else
                        Select Case mode
                            Case "F"
                                GetTableFull(s, ss)
                            Case "I"
                                GetTableIntermediate(s, ss)
                            Case Else
                                GetTable(s, ss)
                        End Select
                    End If
                Else
                    GetText(s, ss, st, sqllib.GetString(dr.Item("parent")))
                End If
            Next

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
            Return False
        End Try

        Return True
    End Function

#Region "common functions"
    Private Function GetTable(ByVal sTable As String, ByVal Schema As String) As Integer
        Dim ts As New TableDefn(sTable, Schema, sqllib)
        Dim sOut As String
        Dim s As String

        If ts.State = 3 Then
            SendMessage("Table " & Schema & "." & sTable & " was not found and not scripted.", "T")
            Return 0
        End If

        sOut = ts.TableText(opt)
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf

        For Each ic As TableIndex In ts.IKeys
            If Not ic.PrimaryKey Then
                sOut &= ic.IndexShort
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            End If
        Next

        For Each fk As ForeignKey In ts.FKeys
            sOut &= fk.ForeignKeyShort
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        Next

        If IncludePerm Then
            s = ts.Permissions.Text
            If s <> "" Then
                sOut &= s
                sOut &= "go" & vbCrLf
            End If
        End If

        WriteFile("table", Schema, sTable, "", "sql", sOut)
        Return 0
    End Function

    Private Function GetTableIntermediate(ByVal sTable As String, _
                    ByVal Schema As String) As Integer
        Dim ts As New TableDefn(sTable, Schema, sqllib)
        Dim sOut As String
        Dim s As String

        sOut = ts.TableText(opt)
        sOut &= "go" & vbCrLf
        WriteFile("table", Schema, sTable, "", "sql", sOut)

        For Each ic As TableIndex In ts.IKeys
            If Not ic.PrimaryKey Then
                sOut = ic.IndexShort
                sOut &= "go" & vbCrLf
                WriteFile("index", Schema, sTable, ic.Name, "sql", sOut)
            End If
        Next

        For Each fk As ForeignKey In ts.FKeys
            sOut = fk.ForeignKeyShort
            sOut &= "go" & vbCrLf
            s = fk.LinkedTable & "." & fk.Name
            If fk.LinkedSchema <> "dbo" Then
                s = fk.LinkedSchema & "." & s
            End If
            WriteFile("fkey", Schema, sTable, s, "sql", sOut)
        Next

        If IncludePerm Then
            s = ts.Permissions.Text
            If s <> "" Then
                sOut = s
                sOut &= "go" & vbCrLf
                WriteFile("perm", Schema, sTable, "", "sql", sOut)
            End If
        End If

        Return 0
    End Function

    Private Function GetTableFull(ByVal sTable As String, ByVal Schema As String) As Integer
        Dim ts As New TableDefn(sTable, Schema, sqllib)
        Dim sOut As String
        Dim s As String

        sOut = ts.FullTableText(opt)
        sOut &= "go" & vbCrLf
        WriteFile("table", Schema, sTable, "", "sql", sOut)

        For Each ic As TableIndex In ts.IKeys
            If Not ic.PrimaryKey Then
                sOut = ic.IndexText
                sOut &= "go" & vbCrLf
                WriteFile("index", Schema, sTable, ic.Name, "sql", sOut)
            End If
        Next

        For Each fk As ForeignKey In ts.FKeys
            sOut = fk.ForeignKeyText
            sOut &= "go" & vbCrLf
            s = fk.LinkedTable & "." & fk.Name
            If fk.LinkedSchema <> "dbo" Then
                s = fk.LinkedSchema & "." & s
            End If
            WriteFile("fkey", Schema, sTable, s, "sql", sOut)
        Next

        If IncludePerm Then
            s = ts.Permissions.Text
            If s <> "" Then
                sOut = s
                sOut &= "go" & vbCrLf
                WriteFile("perm", Schema, sTable, "", "sql", sOut)
            End If
        End If

        Return 0
    End Function

    Private Function GetTableXML(ByVal sTable As String, ByVal Schema As String) As Integer
        Dim ts As New TableDefn(sTable, Schema, sqllib)
        Dim sOut As String
        Dim s As String

        sOut = ts.XML(opt)
        WriteFile("table", Schema, sTable, "", "tdef", sOut)

        If IncludePerm Then
            s = ts.Permissions.Text
            If s <> "" Then
                sOut = s
                sOut &= "go" & vbCrLf
                WriteFile("perm", Schema, sTable, "", "sql", sOut)
            End If
        End If

        Return 0
    End Function

    Private Function GetText(ByVal Name As String, ByVal Schema As String, ByVal Type As String, ByVal Parent As String) As Integer
        Dim sPerm As String = ""
        Dim sText As String
        Dim sHead As String
        Dim Settings As String = ""
        Dim Pre As String = ""
        Dim sName As String
        Dim qName As String
        Dim qSchema As String

        qName = sqllib.QuoteIdentifier(Name)
        qSchema = sqllib.QuoteIdentifier(Schema)
        sText = GetdbText(qName, qSchema, Type)
        sName = sqllib.getName(sText)

        If qName <> sName And qSchema & "." & qName <> sName Then
            Select Case Type
                Case "P"
                    Pre = "Procedure"
                Case "V"
                    Pre = "View"
                Case "FN", "TF"
                    Pre = "Function"
                Case "TR"
                    Pre = "Trigger"
                Case Else
                    Pre = "Object"
            End Select

            SendMessage(Pre & " " & Schema & "." & Name & " was renamed " & sName & " not scripted.", "T")
            Return 0
        End If

        sHead = "if object_id('" & qSchema & "." & qName & "') is not null" & vbCrLf
        sHead &= "begin" & vbCrLf
        sHead &= "    drop "

        Select Case Type
            Case "P"
                Pre = "proc"
                sHead &= "procedure"
                If IncludePerm Then
                    sPerm = ProcPermissions(qName, qSchema)
                    If sPerm <> "" Then
                        Select Case mode
                            Case "F", "I"
                                sPerm &= "go" & vbCrLf
                                WriteFile("perm", Schema, Name, "", "sql", sPerm)

                            Case Else
                                sText &= "go" & vbCrLf
                                sText &= vbCrLf
                                sText &= sPerm
                        End Select
                    End If
                End If
                If mode = "F" Then
                    Settings = GetSetings(qName, qSchema)
                End If

            Case "V"
                Pre = "view"
                sHead &= "view"

            Case "FN" 'scalar returning function
                Pre = "udf"
                sHead &= "function"
                If IncludePerm Then
                    sPerm = ProcPermissions(qName, qSchema)
                    If sPerm <> "" Then
                        Select Case mode
                            Case "F", "I"
                                sPerm &= "go" & vbCrLf
                                WriteFile("perm", Schema, Name, "", "sql", sPerm)

                            Case Else
                                sText &= "go" & vbCrLf
                                sText &= vbCrLf
                                sText &= sPerm
                        End Select
                    End If
                End If
                If mode = "F" Then
                    Settings = GetSetings(qName, qSchema)
                End If

            Case "TF" 'table returning function
                Pre = "udf"
                sHead &= "function"
                If IncludePerm Then
                    sPerm = TFNPermissions(qName, qSchema)
                    If sPerm <> "" Then
                        Select Case mode
                            Case "F", "I"
                                sPerm &= "go" & vbCrLf
                                WriteFile("perm", Schema, Name, "", "sql", sPerm)

                            Case Else
                                sText &= "go" & vbCrLf
                                sText &= vbCrLf
                                sText &= sPerm
                        End Select
                    End If
                End If
                If mode = "F" Then
                    Settings = GetSetings(qName, qSchema)
                End If

            Case "TR"
                Pre = "trigger." & Parent & "."
                sHead &= "trigger"
                If mode = "F" Then
                    Settings = GetSetings(qName, qSchema)
                End If
        End Select

        sHead &= " " & qSchema & "." & qName & vbCrLf
        sHead &= "end" & vbCrLf
        sHead &= "go" & vbCrLf
        sHead &= Settings

        If mode = "F" Then
            sText = sHead & sText
            sText &= "go" & vbCrLf
        End If

        WriteFile(Pre, Schema, Name, "", "sql", sText)
        Return 0
    End Function

    Private Function GetdbText(ByVal Name As String, ByVal Schema As String, ByVal Type As String) As String
        Dim s As String
        Dim o As Object
        Dim sText As String
        Dim dr As DataRow
        Dim dt As New DataTable

        dt = sqllib.ObjectText(Name, Schema)

        sText = ""
        For Each dr In dt.Rows        ' Columns
            o = dr.Item("Text")
            If IsDBNull(o) Then
                s = ""
            ElseIf o Is Nothing Then
                s = ""
            Else
                s = CType(o, String)
            End If
            If Len(s) < 4000 Then s &= vbCrLf
            sText &= s
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

    Private Function GetSetings(ByVal Name As String, ByVal Schema As String) As String
        Dim sText As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        dt = sqllib.ObjectSettings(Name, Schema)

        If Not dt Is Nothing Then
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                If sqllib.GetInteger(dr("nulls"), -1) = 0 Then
                    sText &= "set ansi_nulls off" & vbCrLf
                End If
                If sqllib.GetInteger(dr("quoted"), -1) = 0 Then
                    sText &= "set quoted_identifier off" & vbCrLf
                End If

                If sText <> "" Then
                    sText &= "go" & vbCrLf
                End If
            End If
        End If

        Return sText
    End Function

    Private Function ProcPermissions(ByVal Name As String, ByVal Schema As String) As String
        Dim sText As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        dt = sqllib.ProcPermissions(Name, Schema)

        If Not dt Is Nothing Then
            For Each dr In dt.Rows
                Select Case sqllib.GetString(dr("state"))
                    Case "G"
                        sText &= "grant execute on " & Schema & "." & Name & " to " _
                              & sqllib.GetString(dr("grantee")) & vbCrLf
                    Case "W"
                        sText &= "grant execute on " & Schema & "." & Name & " to " _
                              & sqllib.GetString(dr("grantee")) & " with grant option" & vbCrLf
                    Case "D"
                        sText &= "deny execute on " & Schema & "." & Name & " to " _
                              & sqllib.GetString(dr("grantee")) & vbCrLf
                End Select
            Next
        End If
        Return sText
    End Function

    Private Function TFNPermissions(ByVal Name As String, ByVal Schema As String) As String
        Dim dt As DataTable
        Dim dc As DataTable = Nothing
        Dim dr As DataRow
        Dim i As Integer
        Dim j As Integer
        Dim sOut As String = ""
        Dim s As String
        Dim sC As String

        dt = sqllib.TablePermissions(Name, Schema)

        If dt Is Nothing Then
            Return ""
        End If

        If dt.Rows.Count = 0 Then
            Return ""
        End If

        For Each r As DataRow In dt.Rows
            s = LCase(sqllib.GetString(r.Item("permission_name")))
            j = sqllib.GetInteger(r.Item("columns"), 0)
            If j > 1 Then
                If dc Is Nothing Then
                    dc = sqllib.FunctionColumns(Name, Schema)
                End If
                sC = ""
                s &= " ("
                i = 1
                For Each dr In dc.Rows
                    If (j And CInt(2 ^ i)) <> 0 Then
                        s &= sC & sqllib.GetString(dr("name"))
                        sC = ", "
                    End If
                    i += 1
                Next
                s &= ")"
            End If
            s &= " on " & Schema & "." & sqllib.QuoteIdentifier(Name)
            s &= " to " & sqllib.GetString(r.Item("grantee"))
            Select Case sqllib.GetString(r.Item("state"))
                Case "GRANT_WITH_GRANT_OPTION"
                    sOut &= "grant " & s & " with grant option" & vbCrLf

                Case "GRANT"
                    sOut &= "grant " & s & vbCrLf

                Case "DENY"
                    sOut &= "deny " & s & vbCrLf

            End Select
        Next

        Return sOut
    End Function

    Private Function GetJob(ByVal sJobID As String, ByVal mode As String) As Integer
        Dim js As New Job(sJobID, sqllib)
        Dim sOut As String = ""
        Dim s As String

        s = js.JobName
        If s = "" Then Return -1

        s = Replace(s, ":", "_")
        s = Replace(s, "\", "_")
        s = Replace(s, "/", "_")
        If JobSQL Then
            sOut = js.StepSQL
            If sOut <> "" Then
                WriteFile("jobsql", "", s, "", "sql", sOut)
            End If
        Else
            If mode = "F" Then
                sOut = js.FullText
            Else
                sOut = js.CommonText
            End If
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf

            WriteFile("job", "", s, "", "sql", sOut)
        End If
        Return 0
    End Function

    Private Function WriteFile(ByVal Pre As String, ByVal Schema As String, _
                          ByVal ObjectName As String, ByVal Post As String, _
                          ByVal Ext As String, ByVal Content As String) As Boolean
        Dim i As Integer
        Dim s As String = ""
        Dim ss As String = ""
        Dim sd As String = ""

        If Pre <> "" Then
            s = Pre
            sd = "."
        End If
        If Schema <> "" And Schema <> "dbo" Then
            s &= sd & Schema
            sd = "."
        End If
        If ObjectName <> "" Then
            s &= sd & ObjectName
            sd = "."
        End If
        If Post <> "" Then
            s &= sd & Post
            sd = "."
        End If
        If Ext = "" Then
            s &= sd & "sql"
        Else
            s &= sd & Ext
        End If

        If s = "" Then
            SendMessage("No FileName provided.", "E")
            Return False
        End If

        Dim file As System.IO.StreamWriter

        For i = 1 To Len(s)
            ss = LCase(Mid(s, i, 1))
            If InStr(System.IO.Path.GetInvalidPathChars(), ss, CompareMethod.Text) <> 0 Then
                s = Replace(s, ss, "#0x" & Hex(Asc(ss)) & "#")
            End If
        Next

        If UniCode Then
            file = New System.IO.StreamWriter(s, False, System.Text.Encoding.Unicode)
        Else
            file = New System.IO.StreamWriter(s)
        End If

        file.Write(Content)
        file.Close()
        SendMessage(s, "N")
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
            sParameter = LTrim(Mid(sCommand, i + 2))
            If Mid(sParameter, 1, 1) = "-" Then
                sParameter = ""
            ElseIf Mid(sParameter, 1, 1) = """" Then
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
        Return sParameter
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
