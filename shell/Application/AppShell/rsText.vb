Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class rsText
    Private sErrorMsg As String = ""
    Private sUserMsg As String = ""
    Private bRequired As Boolean = False
    Private bEnabled As Boolean = True
    Private bBorder As Boolean = True
    Private borderColour As Color

    Public Sub New()
        InitializeComponent()
        Dim i As Integer = DialogStyle.BorderWidth
        Me.TextBox1.Top = i
        Me.TextBox1.Left = i + 1
        SetColour()
        Me.TextBox1.ForeColor = DialogStyle.NameToColour(DialogStyle.ForeColour)
        Me.Height = 21 + 2 * i
        Me.Width = Me.TextBox1.Width + 2 * i + 1
    End Sub

    Public Overloads Property Height() As Integer
        Get
            Height = MyBase.Height
        End Get
        Set(ByVal h As Integer)
            Dim i As Integer = DialogStyle.BorderWidth
            Dim j As Integer
            If Me.TextBox1.Multiline Then
                Me.TextBox1.Height = h - 2 * i
                MyBase.Height = h
            Else
                j = Me.TextBox1.Height + 2 * i
                If h <= j Then
                    MyBase.Height = j
                Else
                    MyBase.Height = h
                    j = (h - j) / 2 + i
                    Me.TextBox1.Top = j
                End If
            End If
            Me.Invalidate()
        End Set
    End Property

    Public Overloads Property Width() As Integer
        Get
            Width = MyBase.Width
        End Get
        Set(ByVal w As Integer)
            Dim i As Integer = DialogStyle.BorderWidth
            If w < 10 Then
                MyBase.Width = 10
            Else
                Me.TextBox1.Width = w - 2 * i - 1
                MyBase.Width = w
            End If
            Me.Invalidate()
        End Set
    End Property

    Public Property Multiline() As Boolean
        Get
            Multiline = Me.TextBox1.Multiline
        End Get
        Set(ByVal ml As Boolean)
            Dim i As Integer = DialogStyle.BorderWidth
            Me.TextBox1.Multiline = ml
            If ml Then
                Me.TextBox1.Height = MyBase.Height - 2 * i
            End If
        End Set
    End Property

    Public Property Required() As Boolean
        Get
            Required = bRequired
        End Get
        Set(ByVal rq As Boolean)
            bRequired = rq
            SetColour()
        End Set
    End Property

    Public Property Border() As Boolean
        Get
            Border = bBorder
        End Get
        Set(ByVal bd As Boolean)
            bBorder = bd
        End Set
    End Property

    Public Overloads Property Enabled() As Boolean
        Get
            Enabled = bEnabled
        End Get
        Set(ByVal ro As Boolean)
            bEnabled = ro
            Me.TextBox1.ReadOnly = Not ro
        End Set
    End Property

    Public Property Justify() As String
        Get
            Dim s As String
            Select Case Me.TextBox1.TextAlign
                Case HorizontalAlignment.Right
                    s = "R"
                Case HorizontalAlignment.Center
                    s = "C"
                Case Else
                    s = "L"
            End Select
            Justify = s
        End Get
        Set(ByVal jt As String)
            Select Case jt
                Case "R"
                    Me.TextBox1.TextAlign = HorizontalAlignment.Right
                Case "C"
                    Me.TextBox1.TextAlign = HorizontalAlignment.Center
                Case Else
                    Me.TextBox1.TextAlign = HorizontalAlignment.Left
            End Select
        End Set
    End Property

    Public Property WordWrap() As Boolean
        Get
            WordWrap = Me.TextBox1.WordWrap
        End Get
        Set(ByVal ww As Boolean)
            Me.TextBox1.WordWrap = ww
        End Set
    End Property

    Public ReadOnly Property Lines() As String()
        Get
            Lines = Me.TextBox1.Lines
        End Get
    End Property

    Public Property ScrollBars() As ScrollBars
        Get
            ScrollBars = Me.TextBox1.ScrollBars
        End Get
        Set(ByVal sb As ScrollBars)
            Me.TextBox1.ScrollBars = sb
        End Set
    End Property

    Public ReadOnly Property CanUndo() As Boolean
        Get
            CanUndo = Me.TextBox1.CanUndo
        End Get
    End Property

    Public Sub Copy()
        Me.TextBox1.Copy()
    End Sub

    Public Sub Cut()
        Me.TextBox1.Cut()
    End Sub

    Public Sub Paste()
        Me.TextBox1.Paste()
    End Sub

    Public Sub Undo()
        Me.TextBox1.Undo()
    End Sub

    Public Sub SelectAll()
        Me.TextBox1.SelectAll()
    End Sub

    Public ReadOnly Property ErrorMsg() As String
        Get
            Dim s As String
            SetColour()
            s = sErrorMsg
            If s <> "" And sUserMsg <> "" Then s &= vbCrLf
            s &= sUserMsg
            ErrorMsg = s
        End Get
    End Property

    Public Property UserMsg() As String
        Get
            UserMsg = sUserMsg
        End Get
        Set(ByVal um As String)
            sUserMsg = um
            SetColour()
        End Set
    End Property

    Public Overrides Property Text() As String
        Get
            Text = Me.TextBox1.Text
        End Get
        Set(ByVal txt As String)
            Me.TextBox1.Text = txt
        End Set
    End Property

    Public Property SelectedText() As String
        Get
            SelectedText = Me.TextBox1.SelectedText
        End Get
        Set(ByVal txt As String)
            Me.TextBox1.SelectedText = txt
        End Set
    End Property

    Public Overrides Property ContextMenu() As System.Windows.Forms.ContextMenu
        Get
            ContextMenu = Me.TextBox1.ContextMenu
        End Get
        Set(ByVal cm As System.Windows.Forms.ContextMenu)
            Me.TextBox1.ContextMenu = cm
        End Set
    End Property

    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
        If Not Me.TextBox1.Multiline Then
            Me.TextBox1.SelectionStart = 0
            Me.TextBox1.SelectionLength = Len(Me.TextBox1.Text)
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics
        Dim r As Rectangle
        SetColour()
        r = New Rectangle(e.Graphics.VisibleClipBounds.X, e.Graphics.VisibleClipBounds.Y, e.Graphics.VisibleClipBounds.Width - 1, e.Graphics.VisibleClipBounds.Height - 1)
        g.FillRectangle(New SolidBrush(Me.TextBox1.BackColor), r)
        g.DrawRectangle(New Pen(borderColour, DialogStyle.BorderWidth), r)
    End Sub

    Private Sub SetColour()
        Dim bgColour, borColour As Color
        Dim b As Boolean = False

        If bRequired And bEnabled Then
            bgColour = DialogStyle.NameToColour(DialogStyle.BackRequired)
            If Me.TextBox1.Text = "" Then
                sErrorMsg = "Must not be empty"
            Else
                sErrorMsg = ""
            End If
        Else
            bgColour = DialogStyle.NameToColour(DialogStyle.BackNormal)
        End If

        If bBorder Then
            If sUserMsg = "" And sErrorMsg = "" Then
                borColour = DialogStyle.NameToColour(DialogStyle.BorderNormal)
            Else
                borColour = DialogStyle.NameToColour(DialogStyle.BorderError)
            End If
        Else
            borColour = bgColour
        End If

        If Me.TextBox1.BackColor <> bgColour Then
            MyBase.BackColor = bgColour
            Me.TextBox1.BackColor = bgColour
            b = True
        End If
        If borderColour <> borColour Then
            borderColour = borColour
            b = True
        End If
        If b Then Me.Invalidate()
    End Sub
End Class
