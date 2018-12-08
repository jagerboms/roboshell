Option Explicit On
Option Strict On

Imports System.IO
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Data.SqlClient

Public Class SQLCompliler
    Dim FileName As String
    Dim DataBase As String
    Dim dbSource As String
    Dim CurrentNode As TreeNode
    Dim CurrentDB As String
    Dim RunState As Integer = 0
    Dim hasMissing As Boolean
    Dim badDB As Boolean = False
    Dim Files As SQLFiles
    Dim Cons As Connects
    Dim sResult As String = ""
    Dim sTitle As String = "SQL Compiler"

    Private Sub SQLCompliler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = sTitle
        SetButtons()
        FileName = GetCommandParameter("-f")
        DataBase = GetCommandParameter("-d")
        LoadDatabases()
        LoadFile()
    End Sub

    Private Sub TSOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSOpen.Click
        Dim openFileDialog As New System.Windows.Forms.OpenFileDialog

        openFileDialog.Title = sTitle
        If FileName <> "" Then
            openFileDialog.InitialDirectory = Path.GetDirectoryName(FileName)
        End If
        openFileDialog.Filter = "SQL Compile Lists (*.scl)|*.scl|All files (*.*)|*.*"
        openFileDialog.FilterIndex = 1
        openFileDialog.Multiselect = False
        openFileDialog.RestoreDirectory = True

        If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FileName = openFileDialog.FileName
            LoadFile()
        End If
        SetButtons()
    End Sub

    Private Sub TSRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSRefresh.Click
        LoadFile()
    End Sub

    Private Sub TSStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSStart.Click
        Dim bDBOK As Boolean = False

        If hasMissing Then
            If MsgBox("There are missing files, are you sure you wish to continue?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        RunState = 9
        SetButtons()
        For Each sql As SQLFile In Files
            Select Case sql.FileType
                Case "DB"
                    If sql.State <> "E" Then
                        DataBase = sql.Name
                        bDBOK = True
                    Else
                        bDBOK = False
                    End If
                Case "SQL"
                    If bDBOK And sql.State = "U" Then
                        CompileSQL(sql.File)
                    End If
                Case "FILE"
                    If bDBOK And sql.State = "U" Then
                        sResult = ""
                        If CompileFile(sql.File) Then
                            sql.State = "C"
                        Else
                            sql.State = "E"
                        End If
                        sql.Results = sResult
                        If sql.Node.IsSelected Then
                            View()
                        End If
                    End If
            End Select
            Application.DoEvents()
        Next
        RunState = 2
        SetButtons()
    End Sub

    Private Sub TSLicence_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSLicence.Click
        Dim vw As New view
        Dim s As String = ""

        s = System.Reflection.Assembly.GetExecutingAssembly.Location
        s = Path.GetDirectoryName(s)
        s = Path.Combine(s, "Licence.txt")
        Dim file As New System.IO.StreamReader(s)
        s = file.ReadToEnd
        vw.Text = sTitle & " Licencing"
        vw.Output.Text = s
        vw.Output.BackColor = Drawing.Color.White
        vw.Output.ScrollBars = RichTextBoxScrollBars.Vertical
        vw.Width = 400
        vw.Height = 500
        vw.Show()
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        View()
    End Sub

    Private Sub LoadDatabases()
        Dim s As String
        Dim settings As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        Cons = New Connects
        If Not settings Is Nothing Then
            For Each x As ConnectionStringSettings In settings
                s = x.Name
                If s <> "LocalSqlServer" Then
                    Cons.Add(s, x.ConnectionString, x.ProviderName)
                End If
            Next
        End If
    End Sub

    Private Sub LoadFile()
        Files = New SQLFiles
        Me.TreeView1.Nodes.Clear()
        hasMissing = False
        CurrentDB = ""
        badDB = False
        DataBase = GetCommandParameter("-d")
        If DataBase <> "" Then
            dbSource = "Parameter"
        Else
            dbSource = "default"
            Dim instance As New System.Configuration.AppSettingsReader
            instance.GetValue(dbSource, dbSource.GetType).ToString()
            DataBase = instance.GetValue(dbSource, dbSource.GetType).ToString()
        End If
        If FileName <> "" Then
            Me.Text = "SQL Compiler - " & FileName
            LoadIt(FileName)
            Me.TreeView1.ExpandAll()
        Else
            Me.Text = "SQL Compiler"
        End If
    End Sub

    Private Sub LoadIt(ByVal sList As String)
        Dim line As String
        Dim s As String
        Dim sFile As String
        Dim sPWD As String
        Dim base As TreeNode = Nothing

        RunState = 0
        sPWD = Environment.CurrentDirectory
        Try
            s = Path.GetDirectoryName(sList)
            If s <> "" Then
                sFile = Path.GetFileName(sList)
                Environment.CurrentDirectory = s
            Else
                sFile = sList
            End If

            's = Path.Combine(s, sFile)
            If Dir(sFile) = "" Then
                AddNode(s, "SCL", s)
            Else
                Dim file As New System.IO.StreamReader(sFile)
                ' Read and display the lines from the file until the end 
                ' of the file is reached.
                Do
                    line = file.ReadLine()

                    If line.TrimEnd <> "" Then
                        s = ""
                        Select Case Mid(line, 1, 1)
                            Case "#"    ' ignore comments

                            Case "{"    ' configuration parameter
                                If LCase(Mid(line, 1, 5)) = "{scl=" Then
                                    s = Mid(line, 6, 1500)
                                    s = s.Replace("}", "")
                                    s = Path.Combine(Environment.CurrentDirectory, s)
                                    LoadIt(s)
                                ElseIf LCase(Mid(line, 1, 5)) = "{sql=" Then
                                    s = Mid(line, 6, 1500)
                                    s = s.Replace("}", "")
                                    AddNode(s, "SQL", sList)
                                ElseIf LCase(Mid(line, 1, 10)) = "{database=" Then
                                    s = Mid(line, 11, 80)
                                    s = s.Replace("}", "")
                                    dbSource = sList
                                    DataBase = s
                                End If

                            Case Else
                                s = Path.Combine(Environment.CurrentDirectory, line)
                                Select Case LCase(Path.GetExtension(s))
                                    Case ".sql"
                                        AddNode(s, "FILE", sList)
                                    Case ".tdef"
                                        AddNode(s, "TDEF", sList)
                                    Case Else
                                        AddNode(s, "FILE", sList)
                                End Select
                        End Select
                    End If
                Loop Until file.EndOfStream
                file.Close()
            End If
            RunState = 1

        Catch ex As Exception
            AddNodeError(sList, "FILE", sList, ex.Message)
        End Try

        Environment.CurrentDirectory = sPWD
        SetButtons()
    End Sub

    Private Sub AddNode(ByVal File As String, ByVal Type As String, ByVal Source As String)
        AddNodeError(File, Type, Source, "")
    End Sub

    Private Sub AddNodeError(ByVal File As String, ByVal Type As String, ByVal Source As String, ByVal sResult As String)
        Dim s As String
        Dim st As String
        Dim sErr As String
        Dim b As Boolean = False
        Dim sql As SQLFile
        Dim co As Connect

        If CurrentDB <> DataBase Then
            CurrentNode = Nothing
        End If

        If CurrentNode Is Nothing Then
            If DataBase = "" Then
                s = "no database"
                st = "U"
                dbSource = "missing"
                sErr = "No database key defined."
                badDB = True
            Else
                s = DataBase
                co = Cons.Item(DataBase)
                If co Is Nothing Then
                    st = "E"
                    sErr = "The provided database key was not found in the configuration file."
                    badDB = True
                Else
                    b = True
                    If co.State = "OK" Then
                        st = "C"
                    Else
                        st = "E"
                        badDB = True
                    End If
                    sErr = co.ErrorText
                End If
            End If
            sql = Files.Add(s, dbSource, b, st, sErr)
            CurrentNode = sql.Node
            sql = Nothing
            Me.TreeView1.Nodes.Add(CurrentNode)
            CurrentDB = DataBase
        End If

        sql = Files.Add(File, Type, Source)
        If sResult <> "" Then
            sql.Results = sResult
        End If
        If Not sql.Exists Then
            hasMissing = True
        End If
        CurrentNode.Nodes.Add(sql.Node)
    End Sub

    Private Sub View()
        Dim vw As New view
        Dim tn As TreeNode
        Dim sql As SQLFile
        Dim c As Connect
        Dim sOut As String
        Dim st As String

        sOut = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fmodern\fprq1\fcharset0 Courier New;}}{\colortbl ;\red0\green0\blue255;\red255\green0\blue0;}\viewkind4\uc1\pard \tx450\cf1\lang1033\f0\fs22 "

        tn = Me.TreeView1.SelectedNode
        If tn Is Nothing Then
            st = FileName
            sOut &= "\b File:\b0\tab " & rtfOut(FileName) & "\par "
            If Dir(FileName) = "" Then
                sOut &= "\b Error:\b0\tab File does not exist!\par "
            End If
        Else
            st = tn.Tag.ToString
            sql = Files.Item(st)
            st = tn.Text
            If sql Is Nothing Then
                sOut &= "\b Object:\b0\tab " & rtfOut(st) & "\par "
                sOut &= "\b Error:\b0\tab Cannot find object!\par "
            Else
                Select Case sql.FileType
                    Case "DB"
                        sOut &= "\b Database:\b0\tab " & rtfOut(sql.Name) & "\par "
                        c = Cons.Item(sql.Name)
                        If Not c Is Nothing Then
                            sOut &= "\b Connect:\b0\tab " & rtfOut(c.ConnectString) & "\par "
                            sOut &= "\b Provider:\b0\tab " & rtfOut(c.Provider) & "\par "
                        End If

                    Case "SCL"
                        sOut &= "\b SCL File:\b0\tab " & rtfOut(sql.Name) & "\par "

                    Case "FILE"
                        sOut &= "\b SQL File:\b0\tab " & rtfOut(sql.Name) & "\par "

                    Case "SQL"
                        sOut &= "\b SQL Text:\b0\tab " & rtfOut(sql.File) & "\par "

                    Case Else
                        sOut &= "\b Unknown Type:\b0\tab " & rtfOut(sql.FileType) & "\par "
                        sOut &= "\b Name:\b0\tab " & rtfOut(sql.File) & "\par "

                End Select
                If sql.Results <> "" Then
                    If sql.State = "E" Then
                        sOut &= "\b Error"
                    Else
                        sOut &= "\b Output"
                    End If
                    sOut &= ":\b0\par " & rtfOut(sql.Results) & " \par "
                End If
            End If
        End If
        vw.Text = sTitle & " " & st
        sOut &= " }"
        Me.Output.Rtf = sOut
        Me.Output.BackColor = Drawing.Color.White
        Me.Show()
    End Sub

    Private Function CompileFile(ByVal sFile As String) As Boolean
        Dim s As String = ""
        Dim bErrorLatch As Boolean

        Try
            Dim file As New System.IO.StreamReader(sFile)
            s = file.ReadToEnd
            file.Close()
        Catch ex As Exception
            SaveOutput("Error compiling " & sFile & vbCrLf & ex.Message, "E")
            Return False
        End Try

        If s <> "" Then
            bErrorLatch = True
            If Not CompileSQL(s) Then
                bErrorLatch = False
            End If
        Else
            SaveOutput(sFile & " contains no text!", "M")
            bErrorLatch = False
        End If

        Return bErrorLatch
    End Function

    Private Function CompileSQL(ByVal sText As String) As Boolean
        Dim sCommands As String = ""
        Dim i As Integer
        Dim j As Integer = 1
        Dim k As Integer = 0
        Dim cc As Integer = 0
        Dim Mode As Integer = 0
        Dim b As Boolean = True
        Dim c As String
        Dim result As IAsyncResult
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim cn As Connect

        Try
            CompileSQL = False

            cn = Cons.Item(DataBase)
            If cn Is Nothing Then
                SaveOutput("CompileSQL: Error retreving connsection string for Database '" & DataBase & "'.", "E")
                Return False
            End If
            psConn = New SqlConnection(cn.ConnectString)
            AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
            psConn.Open()
            psAdapt = New SqlDataAdapter("", psConn)

            sText &= vbCrLf & "go" & vbCrLf
            For i = 1 To Len(sText)
                c = Mid(sText, i, 1)
                If Mode = 0 Then    ' for looking for go
                    If c <> vbCr And c <> vbLf Then
                        k = i
                        If LCase(Mid(sText, i, 2)) = "go" Then
                            Mode = 9
                            i += 2
                            c = Mid(sText, i, 1)
                        Else
                            Mode = 1
                        End If
                    End If
                End If
                Select Case Mode
                    Case 1   ' general text
                        Select Case c
                            Case vbCr, vbLf
                                Mode = 0
                            Case "'"
                                Mode = 2
                            Case """"
                                Mode = 3
                            Case "-"
                                If Mid(sText, i + 1, 1) = "-" Then
                                    Mode = 4
                                    i += 1
                                End If
                            Case "/"
                                If Mid(sText, i + 1, 1) = "*" Then
                                    Mode = 5
                                    i += 1
                                    cc = 1
                                End If
                        End Select

                    Case 2     ' quotes
                        If c = "'" Then
                            Mode = 1
                        End If

                    Case 3     ' double quotes
                        If c = """" Then
                            Mode = 1
                        End If

                    Case 4     ' line comment
                        If c = vbCr Or c = vbLf Then
                            Mode = 0
                        End If

                    Case 5     ' block comments
                        Select Case c
                            Case "/"
                                If Mid(sText, i + 1, 1) = "*" Then
                                    i += 1
                                    cc += 1
                                End If
                            Case "*"
                                If Mid(sText, i + 1, 1) = "/" Then
                                    i += 1
                                    cc -= 1
                                    If cc = 0 Then
                                        Mode = 1
                                    End If
                                End If
                        End Select

                    Case 9     ' go?
                        Select Case c
                            Case " ", vbTab

                            Case vbCr
                                If Mid(sText, i + 1, 1) = vbLf Then
                                    i += 1
                                End If
                                Mode = 99

                            Case vbLf
                                Mode = 99

                            Case "-"
                                If Mid(sText, i + 1, 1) = "-" Then
                                    Mode = 98
                                    i += 1
                                End If

                            Case Else
                                Mode = 1
                        End Select

                    Case 98
                        If c = vbCr Then
                            If Mid(sText, i + 1, 1) = vbLf Then
                                Mode = 99
                                i += 1
                            End If
                        Else
                            If c = vbLf Then
                                Mode = 99
                            End If
                        End If

                End Select
                If Mode = 99 Then
                    If k > j Then
                        sCommands = Mid(sText, j, k - j)

                        psAdapt.SelectCommand.CommandText = sCommands
                        'psAdapt.SelectCommand.ExecuteNonQuery()
                        result = psAdapt.SelectCommand.BeginExecuteNonQuery()
                        While Not result.IsCompleted
                            Threading.Thread.Sleep(100)
                        End While
                        psAdapt.SelectCommand.EndExecuteNonQuery(result)

                    End If
                    Mode = 0
                    j = i + 1
                End If
            Next

            CompileSQL = b

            psConn.Close()

        Catch ex As Exception
            SaveOutput(ex.Message, "E")
            CompileSQL = False
        End Try
    End Function

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        For Each ex As SqlError In e.Errors
            SaveOutput(ex.Message, "M")
        Next
    End Sub

    Private Sub SetButtons()
        Select Case RunState
            Case 0
                Me.TSOpen.Enabled = True
                Me.TSRefresh.Enabled = False
                Me.TSStart.Enabled = False
            Case 1
                Me.TSOpen.Enabled = True
                Me.TSRefresh.Enabled = True
                If badDB Then
                    Me.TSStart.Enabled = False
                ElseIf DataBase = "" Then
                    Me.TSStart.Enabled = False
                Else
                    Me.TSStart.Enabled = True
                End If
            Case 2
                Me.TSOpen.Enabled = True
                Me.TSRefresh.Enabled = True
                Me.TSStart.Enabled = False
            Case 9
                Me.TSOpen.Enabled = False
                Me.TSRefresh.Enabled = False
                Me.TSStart.Enabled = False
        End Select
    End Sub

    Private Function rtfOut(ByVal sIn As String) As String
        Dim s As String
        s = Replace(sIn, "\", "\\")
        s = Replace(s, "{", "\{")
        s = Replace(s, "}", "\}")

        s = Replace(s, "<error>", "\cf2 ")
        s = Replace(s, "</error>", "\cf1 ")

        s = Replace(s, vbCrLf, " \par ")
        Return s
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

    Private Sub SaveOutput(ByVal sMsg As String, ByVal sType As String)
        If sType = "E" Then
            sResult &= "<error>" & sMsg & "</error>" & vbCrLf
        Else
            sResult &= sMsg & vbCrLf
        End If
    End Sub
End Class
