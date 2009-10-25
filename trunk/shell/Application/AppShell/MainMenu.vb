Option Explicit On 
'Option Strict On

Imports System.IO

Public Class MainMenu
    Inherits System.Windows.Forms.Form

    Dim sDefn As New MenuDefn("MainMenu")

    Public Property Actions() As ActionDefns
        Get
            Return sDefn.Actions
        End Get
        Set(ByVal value As ActionDefns)
            sDefn.Actions = value
        End Set
    End Property

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
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.SuspendLayout()
        '
        'ContextMenu2
        '
        '
        'MainMenu
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ClientSize = New System.Drawing.Size(423, 60)
        Me.ContextMenu = Me.ContextMenu2
        Me.KeyPreview = True
        Me.Name = "MainMenu"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public WithEvents mAction As ShellMenu ' ShellObject

    Private Sub Navigate_Load(ByVal sender As Object, _
                            ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sStartUp As String
        sDefn.Actions = New ActionDefns
        Me.Visible = False

        Application.DoEvents()
        If Not Publics.InitialiseApp(Me) Then
            Me.Close()
            Exit Sub
        End If
        sStartUp = GetCommandParameter("-p")
        If sStartUp = "" Then
            MainMenu()
        Else
            Dim p As New ShellProcess(sStartUp, Me, Nothing)
            Me.Close()
        End If
        Publics.AboutClose()
    End Sub

    Public Sub Suspend(ByVal Mode As Boolean)
    End Sub

    Public Sub Progress(ByVal Progress As Integer)
    End Sub

    Public Sub ProcessError(ByVal es As ShellMessages)
        Dim s As String = ""
        For Each e As Exception In es
            s &= e.Message & vbCrLf
        Next
        Publics.MessageOut(s)
    End Sub

    Public Sub MsgOut(ByVal msgs As ShellMessages)
        Dim s As String = ""
        Dim sType As String = "I"

        For Each ss As ShellMessage In msgs
            s &= ss.Message & vbCrLf
            If ss.Type = "E" Then
                sType = "E"
            End If
        Next
        Publics.MessageOut(s, sType)
    End Sub

    Private Sub MainMenu()
        Me.Cursor = Cursors.WaitCursor
        Me.Icon = Publics.ShellIcon
        Dim s As String = Publics.GetVariable("SystemName")
        If Publics.GetVariable("Environment") <> "" Then
            s = Publics.GetVariable("Environment") & " - " & s
        End If
        Me.Text = s
        If Publics.IsMDI Then
            Me.IsMdiContainer = True
            Publics.MDIParent = CType(Me, Form)
        Else
            'Me.Width = 100
            Me.MaximizeBox = False
        End If
        Me.BackColor = Publics.GetBackColour()
        Publics.SetFormPosition(Me, "MainMenu")
        InitialiseAction()
        Application.DoEvents()
        Publics.inInit = False
        Me.Cursor = Cursors.Default
        If Format(Publics.BusinessDate, "yyyyMMdd") <> Format(Now(), "yyyyMMdd") Then
            Publics.MessageOut("The business date is not the current date." & vbCrLf & vbCrLf _
                & "      Business date: " & Format(Publics.BusinessDate, "d-MMM-yyyy"), _
                "E")
        End If
        Me.Visible = True
    End Sub

    Private Sub InitialiseAction()
        If sDefn.Actions.count = 0 Then
            Publics.MessageOut("No menu definition can not continue!")
            Exit Sub
        End If

        mAction = New ShellMenu(sDefn)
        mAction.fForm = Me
        mAction.Update(Nothing)
    End Sub

    Private Sub Navigate_KeyUp(ByVal sender As Object, _
                ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        If e.KeyCode = Keys.F1 Then
            Publics.RaiseHelp(e.Shift, "MainMenu")
        End If
    End Sub

    Private Sub mAction_Action(ByVal sAction As String) Handles mAction.Action
        If sAction <> "" Then
            For Each a As ActionDefn In sDefn.Actions
                If a.Name = sAction And a.Enabled Then
                    Dim p As New ShellProcess(a.Process, Me, Nothing)
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub Navigate_Closing(ByVal sender As Object, _
            ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim s As String
        If Not Publics.inInit Then
            s = Publics.GetVariable("SystemName")
            If s = "" Then
                s = "Application"
            End If
            If MsgBox("Are you sure you want to exit", MsgBoxStyle.YesNo, s) = _
                                                                MsgBoxResult.No Then
                e.Cancel = True
            Else
                SaveFormPosition(Me, "MainMenu")
            End If
        End If
    End Sub

    Private Sub ContextMenu2_Popup(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles ContextMenu2.Popup

        Dim mi As MenuItem
        Me.ContextMenu2.MenuItems.Clear()
        mi = Me.ContextMenu2.MenuItems.Add("Set multiple form", _
                                    New EventHandler(AddressOf SetMDIOn))
        If Publics.IsMDI Then
            mi.Checked = True
        End If
        mi = Me.ContextMenu2.MenuItems.Add("Set single form", _
                                    New EventHandler(AddressOf SetMDIOff))
        If Not Publics.IsMDI Then
            mi.Checked = True
        End If
    End Sub

    Private Sub SetMDIOn(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer = -1

        SaveProperty("MainMenu", "IsMDI", "u", "Y")
        Publics.MessageOut("Multiple document interface is now on" & vbCrLf & _
               "and will activated next time you start the application")
    End Sub

    Private Sub SetMDIOff(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer = -1
        SaveProperty("MainMenu", "IsMDI", "u", "N")
        Publics.MessageOut("Single document interface is now on" & vbCrLf & _
               "and wilzl activated next time you start the application")
    End Sub
End Class

