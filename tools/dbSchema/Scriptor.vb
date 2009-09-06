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
    Dim fixdef As Boolean = False
    Dim UniCode As Boolean = False

    Dim sType As String = ""
    Dim ConsName As Boolean = True
    Dim sObject As String = ""
    Dim mode As String = "S"
    Dim LogFile As String = ""
    Dim verbose As Boolean = False

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
                SendMessage("   -fType determines the type of script to be generated. Can bw one of:", "T")
                SendMessage("      F - full includes existance checks and separate component files", "T")
                SendMessage("      I - intermediate has no existance checks but separate component files", "T")
                SendMessage("      S - summary has no existance checks and all components are in a single file.", "T")
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

            s = GetCommandParameter("-s")
            If s = "" Then s = System.Environment.MachineName
            sqllib.Server = s
            Database = GetCommandParameter("-d")
            sqllib.UserID = GetCommandParameter("-u")
            sqllib.Password = GetCommandParameter("-p")
            sqllib.Network = GetCommandParameter("-n")
            s = GetCommandParameter("-f")
            Select Case UCase(Mid(s, 1, 1))
                Case "F"
                    mode = "F"
                Case "I"
                    mode = "I"
            End Select
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
                If sqllib.Password <> "" Then
                    s = Replace(s, " -p" & sqllib.Password & " ", " -p?????????? ")
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

    Private Sub ProcessJobs(ByVal Database As String)
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
        End Try
    End Sub

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
            PutFile("script.dbSchema-Permission.sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Private Sub ProcessData(ByVal Database As String, ByVal Table As String)
        Dim sOut As String = ""
        Dim s As String

        Try
            sqllib.Database = Database
            s = GetCommandParameter("-w")
            Dim tdefn As New TableColumns(Table, sqllib, True)
            sOut = tdefn.DataScript(s)
            PutFile("data." & Table & ".sql", sOut)

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

    Private Sub ProcessAllDBs()
        Dim s As String
        Dim dbVersion As Integer
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

    Private Sub ProcessDB(ByVal Database As String)
        Dim s As String
        Dim st As String
        Dim dt As DataTable
        Dim dr As DataRow

        Try
            SendMessage("", "N")
            SendMessage("Retrieving schema for '" & Database & "'.", "T")
            sqllib.Database = Database
            dt = sqllib.DatabaseObject(sObject, sType)

            For Each dr In dt.Rows
                s = sqllib.GetString(dr.Item("name"))
                st = sqllib.GetString(dr.Item("type"))
                If sqllib.GetString(dr.Item("type")) = "U" Then
                    Select Case mode
                        Case "F"
                            GetTableFull(s, ConsName)
                        Case "I"
                            GetTableIntermediate(s, ConsName)
                        Case Else
                            GetTable(s, ConsName)
                    End Select
                Else
                    GetText(s, st)
                End If
            Next

        Catch ex As Exception
            SendMessage(ex.ToString, "E")
        End Try
    End Sub

#Region "common functions"
    Private Function GetTable(ByVal sTable As String, ByVal ConsName As Boolean) As Integer
        Dim ts As New TableColumns(sTable, sqllib, fixdef)
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

    Private Function GetTableIntermediate(ByVal sTable As String, ByVal ConsName As Boolean) As Integer
        Dim ts As New TableColumns(sTable, sqllib, fixdef)
        Dim sOut As String
        Dim s As String

        If Not ConsName Then ts.ScriptConstraints = False
        sOut = ts.TableText
        sOut &= "go" & vbCrLf
        PutFile("table." & sTable & ".sql", sOut)

        For Each s In ts.IKeys
            If s <> "" Then
                sOut = ts.IndexShort(s)
                sOut &= "go" & vbCrLf
                PutFile("index." & sTable & "." & s & ".sql", sOut)
            End If
        Next

        For Each s In ts.FKeys
            If s <> "" Then
                sOut = ts.FKeyShort(s)
                sOut &= "go" & vbCrLf
                PutFile("fkey." & sTable & "." & ts.LinkedTable(s) & "." & s & ".sql", sOut)
            End If
        Next

        Return 0
    End Function

    Private Function GetTableFull(ByVal sTable As String, ByVal ConsName As Boolean) As Integer
        Dim ts As New TableColumns(sTable, sqllib, fixdef)
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
                PutFile("fkey." & sTable & "." & ts.LinkedTable(s) & "." & s & ".sql", sOut)
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
                If mode = "F" Then
                    Settings = GetSetings(Name)
                End If
            Case "V"
                Pre = "view."
                sHead &= "view"
            Case "FN", "TF"
                Pre = "udf."
                sHead &= "function"
                If mode = "F" Then
                    Settings = GetSetings(Name)
                End If
        End Select

        sHead &= " dbo." & Name & vbCrLf
        sHead &= "end" & vbCrLf
        sHead &= "go" & vbCrLf
        sHead &= Settings

        If mode = "F" Then
            sText = sHead & sText
            sText &= "go" & vbCrLf
        End If

        PutFile(Pre & Name & ".sql", sText)
        Return 0
    End Function

    Private Function GetdbText(ByVal Name As String, ByVal Type As String) As String
        Dim s As String
        Dim sText As String
        Dim dr As DataRow
        Dim b As Boolean = True
        Dim dt As New DataTable

        dt = sqllib.ObjectText(Name)

        sText = ""
        If Type <> "P" And Type <> "FN" And Type <> "TF" Then
            b = False
        End If

        For Each dr In dt.Rows        ' Columns
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
        Dim sText As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        dt = sqllib.ObjectSettings(Name)

        If Not dt Is Nothing Then
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                If CInt(dr("nulls")) = 0 Then
                    sText &= "set ansi_nulls off" & vbCrLf
                End If
                If CInt(dr("quoted")) = 0 Then
                    sText &= "set quoted_identifier off" & vbCrLf
                End If

                If sText <> "" Then
                    sText &= "go" & vbCrLf
                End If
            End If
        End If

        Return sText
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
        If mode = "F" Then
            sOut = js.FullText
        Else
            sOut = js.CommonText
        End If
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf

        PutFile("job." & s & ".sql", sOut)
        Return 0
    End Function

    Private Function PutFile(ByVal sName As String, ByVal sContent As String) As Boolean
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
