Option Explicit On 
Option Strict On

Imports System
Imports System.IO
Imports System.Diagnostics

Public Class MonitorForm
    Inherits System.Windows.Forms.Form
    Friend oOwner As Monitor
    Friend DblClkKey As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents lstLog As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.lstLog = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'ContextMenu2
        '
        '
        'lstLog
        '
        Me.lstLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstLog.ContextMenu = Me.ContextMenu2
        Me.lstLog.Location = New System.Drawing.Point(0, 24)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.Size = New System.Drawing.Size(424, 264)
        Me.lstLog.TabIndex = 0
        '
        'Monitor
        '
        Me.AutoScaleMode = Windows.Forms.AutoScaleMode.None
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 293)
        Me.ContextMenu = Me.ContextMenu2
        Me.Controls.Add(Me.lstLog)
        Me.KeyPreview = True
        Me.Name = "Monitor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Monitor"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Monitor_Load(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = Publics.ShellIcon
    End Sub

    Private Sub Monitor_KeyDown(ByVal sender As Object, _
                    ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        oOwner.ProcessKey(e.KeyCode, e.Modifiers.ToString)
    End Sub

    Private Sub Monitor_Closing(ByVal sender As Object, _
            ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        oOwner.ProcessClose()
    End Sub

    Private Sub ContextMenu2_Popup(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles ContextMenu2.Popup
        oOwner.DoMenu()
    End Sub
End Class

Public Class MonitorDefn
    Inherits ObjectDefn

    Public Title As String
    Public ServiceParameter As String
    Public ServerParameter As String
    Public TitleParameter() As String
    Public HelpPage As String

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject

        ' Create Actions
        If MyBase.Actions.count = 0 Then
            MyBase.Actions.Add("Refresh", "", True, False, "N", False, _
                                         "refresh.gif", "Refresh information")
            MyBase.Actions.Add("Pause", "", True, False, "N", False, _
                                         "pause.gif", "Pause capture")
            MyBase.Actions.Add("Resume", "", False, False, "N", False, _
                                         "resume.gif", "Resume capture")
            MyBase.Actions.Add("Clear", "", True, False, "N", False, _
                                         "clear.gif", "Clear screen")
            MyBase.Actions.Add("View", "", True, False, "N", False, _
                                         "viewfile.gif", "View details")
            MyBase.Actions.Add("Start", "", False, False, "N", False, _
                                         "start.gif", "Start Service")
            MyBase.Actions.Add("Stop", "", False, False, "N", False, _
                                         "stop.gif", "Stop Service")
        End If

        Return CType(New Monitor(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case Name
            Case "Title"
                Title = GetString(Value)
            Case "ServiceParameter"
                ServiceParameter = GetString(Value)
            Case "ServerParameter"
                ServerParameter = GetString(Value)
            Case "TitleParameters"
                TitleParameter = Split(GetString(Value), "||")
            Case "HelpPage"
                HelpPage = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Monitor object")
        End Select
    End Sub
End Class

Public Class Monitor
    Inherits ShellObject

    Private sDefn As MonitorDefn
    Private fForm As MonitorForm
    Dim ToolBar As System.Windows.Forms.ToolStrip
    Private WithEvents mAction As ShellMenu
    Private FileSyn As FileSync
    Private objService As System.ServiceProcess.ServiceController
    Private FileName As String
    Private sMachine As String
    Private sService As String
    Private StateChange As String = ""
    Private sCaption As String

    Public Sub New(ByVal Defn As MonitorDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Shadows ReadOnly Property parms() As ShellParameters
        Get
            Return MyBase.parms
        End Get
    End Property

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim s As String
        Try
            Me.parms.MergeValues(Parms)
            If fForm Is Nothing Then
                fForm = New MonitorForm
                fForm.oOwner = Me
                fForm.BackColor = Publics.GetBackColour

                If Not Parms Is Nothing Then
                    s = sDefn.ServerParameter
                    If Not Parms.Item(s) Is Nothing Then
                        If Not Parms.Item(s).Value Is Nothing Then
                            If Parms.Item(s).Output Then
                                sMachine = GetString(Parms.Item(s).Value)
                            End If
                        End If
                    End If

                    s = sDefn.ServiceParameter
                    If Not Parms.Item(s) Is Nothing Then
                        If Not Parms.Item(s).Value Is Nothing Then
                            If Parms.Item(s).Output Then
                                sService = GetString(Parms.Item(s).Value)
                            End If
                        End If
                    End If
                End If

                fForm.Name = sDefn.Title
                SetCaption()
                If Not Publics.MDIParent Is Nothing Then
                    fForm.MdiParent = Publics.MDIParent
                End If

                FileName = "\\" & sMachine & "\ServiceLogs\" & sService & "_" & _
                                            DateTime.Today.ToString("yyyyMMdd") & ".log"
                Try
                    FileSyn = New FileSync(fForm.lstLog)
                    FileSyn.FileName = FileName
                    FileSyn.DoSync()
                Catch ex As Exception
                End Try
                objService = New System.ServiceProcess.ServiceController(sService, sMachine)
                Publics.SetFormPosition(CType(fForm, Form), sDefn.Name)
                InitialiseAction()
                SetActions()
                fForm.Show()
            Else
                SetCaption()
            End If
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Private Sub SetCaption()
        Dim s As String
        Dim ss As String
        Dim sTemp As String

        s = sDefn.Title
        If Not sDefn.TitleParameter Is Nothing Then
            For Each ss In sDefn.TitleParameter
                sTemp = GetString(parms.Item(ss).Value)
                If sTemp <> "" Then
                    s &= " - " & sTemp
                End If
            Next
        End If
        If s <> sCaption Then
            fForm.Text = s
            sCaption = s
        End If
    End Sub

    Public Overrides Sub Listener(ByVal Parms As ShellParameters)
    End Sub

    Public Overrides Sub Suspend(ByVal Mode As Boolean)
        If Mode Then
            fForm.Cursor = Cursors.WaitCursor
            fForm.Enabled = False
        Else
            fForm.Enabled = True
            fForm.Cursor = Cursors.Default
        End If
    End Sub

    Public Overrides Sub MsgOut(ByVal Msgs As ShellMessages)
        Dim s As String = ""
        Dim b As Boolean = False

        For Each ms As ShellMessage In Msgs
            s &= ms.Message & vbCrLf
            If ms.Type = "U" And Not b Then
                ProcessAction("Refresh")
                b = True
            End If
        Next
        Publics.MessageOut(s)
    End Sub

    Private Sub InitialiseAction()
        mAction = New ShellMenu(sDefn)
        mAction.fForm = fForm
        mAction.Update(Me.parms)

        For Each a As ActionDefn In sDefn.Actions
            If a.IsDblClick Then
                fForm.DblClkKey = a.Name
            End If
        Next
    End Sub

    Private Sub mAction_Action(ByVal sAction As String) Handles mAction.Action
        If sAction <> "" Then
            Me.ProcessAction(sAction)
        End If
    End Sub

    Friend Sub ProcessAction(ByVal sKey As String)
        For Each a As ActionDefn In sDefn.Actions
            If a.Name = sKey And a.Enabled Then
                Select Case sKey
                    Case "Refresh"
                        FileSyn.DoSync()
                    Case "Pause"
                        FileSyn.Pause = True
                    Case "Resume"
                        FileSyn.Pause = False
                    Case "Clear"
                        fForm.lstLog.Items.Clear()
                    Case "View"
                        System.Diagnostics.Process.Start("notepad.exe", FileName)
                    Case "Start"
                        Try
                            StateChange = "t"
                            objService.Start()
                        Catch ex As Exception
                            Publics.MessageOut(ex.ToString)
                        End Try
                    Case "Stop"
                        Try
                            StateChange = "p"
                            objService.Stop()
                        Catch ex As Exception
                            Publics.MessageOut(ex.ToString)
                        End Try
                    Case Else
                        Dim p As New ShellProcess(a.Process, Me, Me.parms)
                        Exit Sub
                End Select
                SetActions()
            End If
        Next
    End Sub

    Friend Sub ProcessClose()
        SaveFormPosition(CType(fForm, Form), sDefn.Name)
    End Sub

    Friend Sub SetActions()
        Try
            objService.Refresh()
            Select Case StateChange
                Case "t"
                    If objService.Status = ServiceProcess.ServiceControllerStatus.Running Then
                        StateChange = ""
                    Else
                        sDefn.Actions.Item("Stop").Enabled = False
                        sDefn.Actions.Item("Start").Enabled = False
                    End If
                Case "p"
                    If objService.Status = ServiceProcess.ServiceControllerStatus.Stopped Then
                        StateChange = ""
                    Else
                        sDefn.Actions.Item("Stop").Enabled = False
                        sDefn.Actions.Item("Start").Enabled = False
                    End If
            End Select

            If StateChange = "" Then
                Select Case objService.Status
                    Case ServiceProcess.ServiceControllerStatus.Running
                        sDefn.Actions.Item("Stop").Enabled = True
                        sDefn.Actions.Item("Start").Enabled = False
                    Case ServiceProcess.ServiceControllerStatus.Stopped
                        sDefn.Actions.Item("Stop").Enabled = False
                        sDefn.Actions.Item("Start").Enabled = True
                End Select
            Else
                sDefn.Actions.Item("Stop").Enabled = False
                sDefn.Actions.Item("Start").Enabled = False
            End If
        Catch ex As Exception
            sDefn.Actions.Item("Stop").Enabled = False
            sDefn.Actions.Item("Start").Enabled = False
        End Try
        sDefn.Actions.Item("Pause").Enabled = Not FileSyn.Pause
        sDefn.Actions.Item("Resume").Enabled = FileSyn.Pause

        For Each a As ActionDefn In sDefn.Actions
            mAction.Enable(a)
        Next
    End Sub

    Friend Sub ProcessKey(ByVal KeyCode As Integer, ByVal Shift As String)
        If KeyCode = Keys.F1 Then
            If Shift = "None" Then
                Publics.RaiseHelp(False, sDefn.HelpPage)
                Exit Sub
            End If
            If Shift = "Shift" Then
                Publics.RaiseHelp(True, sDefn.HelpPage)
                Exit Sub
            End If
        End If

        For Each a As ActionDefn In sDefn.Actions
            If a.KeyCode = KeyCode And a.Enabled And a.IsKey Then
                If a.Shift = Shift Or a.Shift Is Nothing Then
                    ProcessAction(a.Name)
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Friend Sub DoMenu()
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim mi As MenuItem

        fForm.ContextMenu2.MenuItems.Clear()
        mi = fForm.ContextMenu2.MenuItems.Add("Copy to Clipboard", _
                                        New EventHandler(AddressOf mnuClip_Click))
    End Sub

    Private Sub mnuClip_Click(ByVal sender As System.Object, _
                                            ByVal e As System.EventArgs)
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim Line As String = ""
        fForm.Cursor = Cursors.WaitCursor
        For Each s As String In fForm.lstLog.Items
            Line &= s & vbCrLf
            Application.DoEvents()
        Next
        Clipboard.SetDataObject(New DataObject(DataFormats.Text, Line))
        fForm.Cursor = Cursors.Default
    End Sub
End Class

Public Class FileSync
    Private LogWatcher As New FileSystemWatcher
    Private bPauseUpdates As Boolean
    Private lFilePos As Long
    Private lstBox As System.Windows.Forms.ListBox

    Public FileName As String

    Public Sub New(ByVal lst As System.Windows.Forms.ListBox)
        lstBox = lst
    End Sub

    Public Property Pause() As Boolean
        Get
            Pause = bPauseUpdates
        End Get
        Set(ByVal Value As Boolean)
            bPauseUpdates = Value
            If Not Value Then
                UpdateSync()
            End If
        End Set
    End Property

    Public Sub DoSync()
        lFilePos = 0
        bPauseUpdates = False
        Dim sFile As String = FileName.Substring(FileName.LastIndexOf("\") + 1)
        Dim sPath As String = FileName.Substring(0, FileName.LastIndexOf("\"))

        LogWatcher.Path = sPath
        LogWatcher.NotifyFilter = NotifyFilters.LastWrite
        LogWatcher.Filter = sFile
        AddHandler LogWatcher.Changed, AddressOf OnChanged
        LogWatcher.EnableRaisingEvents = True
        lstBox.Items.Clear()
        UpdateSync()  ' force init of first update
    End Sub

    Private Sub OnChanged(ByVal source As Object, ByVal e As FileSystemEventArgs)
        UpdateSync()  ' force update after receiving event from file system
    End Sub

    Private Sub UpdateSync()
        Static bDoing As Boolean = False

        If bPauseUpdates Then   ' updates are paused so just exit
            Exit Sub
        End If

        If FileName = "" Then
            Exit Sub
        End If

        If bDoing Then
            Exit Sub
        End If
        bDoing = True

        Dim fs As New FileStream(FileName, FileMode.Open, _
                                    FileAccess.Read, FileShare.ReadWrite)

        ' if there has not been any change in filelength exit.
        If (fs.Length <= lFilePos) Then
            lFilePos = fs.Length
            bDoing = False
            Exit Sub
        End If

        ' if this is our first read - just read the last 200k
        If (lFilePos = 0 And fs.Length > 200000) Then
            lFilePos = fs.Length - 200000
        End If

        Dim sLine As String = ""
        Dim r As New StreamReader(fs)
        r.BaseStream.Seek(lFilePos, SeekOrigin.Begin)  ' seek to last known location
        lstBox.BeginUpdate()
        Do While r.Peek() >= 0
            sLine = r.ReadLine()
            lstBox.Items.Add(sLine)

            ' keep the size down
            If (lstBox.Items.Count > 5000) Then
                lstBox.Items.RemoveAt(0)
                lstBox.Items(0) = "** Previous Contents removed. Use OPEN button **"
            End If
        Loop
        lFilePos = fs.Position                         ' record where we are up to
        lstBox.EndUpdate()

        r.Close()
        fs.Close()
        lstBox.SelectedIndex = lstBox.Items.Count - 1
        bDoing = False
    End Sub
End Class

