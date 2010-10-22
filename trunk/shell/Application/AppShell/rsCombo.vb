Public Class rsCombo
    Private borderColour As Color = DialogStyle.NameToColour(DialogStyle.BorderNormal)
    Private bRequired As Boolean = False

    Public Sub New()
        InitializeComponent()
        Dim i As Integer = DialogStyle.BorderWidth
        Me.ComboBox1.Top = i
        Me.ComboBox1.Left = i
        SetColour()
        Me.ComboBox1.ForeColor = DialogStyle.NameToColour(DialogStyle.ForeColour)
        Me.Height = Me.ComboBox1.Height + 2 * i
        Me.Width = Me.ComboBox1.Width + 2 * i
    End Sub

    Public Overloads Property Height() As Integer
        Get
            Height = MyBase.Height
        End Get
        Set(ByVal h As Integer)
            Dim i As Integer = DialogStyle.BorderWidth
            Dim j As Integer
            j = Me.ComboBox1.Height + 2 * i
            If h <= j Then
                MyBase.Height = j
            Else
                MyBase.Height = h
                j = (h - j) / 2 + i
                Me.ComboBox1.Top = j
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
            If w < 20 Then
                MyBase.Width = 20
            Else
                Me.ComboBox1.Width = w - 2 * i
                MyBase.Width = w
            End If
            Me.Invalidate()
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

    Public Property DataSource() As Object
        Get
            DataSource = Me.ComboBox1.DataSource
        End Get
        Set(ByVal ds As Object)
            Me.ComboBox1.DataSource = ds
            Me.ComboBox1.SelectedIndex = -1
            SetColour()
        End Set
    End Property

    Public Property ValueMember() As String
        Get
            ValueMember = Me.ComboBox1.ValueMember
        End Get
        Set(ByVal vm As String)
            Me.ComboBox1.ValueMember = vm
        End Set
    End Property

    Public Property DisplayMember() As String
        Get
            DisplayMember = Me.ComboBox1.DisplayMember
        End Get
        Set(ByVal dm As String)
            Me.ComboBox1.DisplayMember = dm
        End Set
    End Property

    Public Overrides Property Text() As String
        Get
            Text = Me.ComboBox1.Text
        End Get
        Set(ByVal txt As String)
            Me.ComboBox1.Text = txt
            SetColour()
        End Set
    End Property

    Public Property SelectedIndex() As Integer
        Get
            SelectedIndex = Me.ComboBox1.SelectedIndex
        End Get
        Set(ByVal si As Integer)
            Me.ComboBox1.SelectedIndex = si
            SetColour()
        End Set
    End Property

    Public Property SelectedText() As String
        Get
            SelectedText = Me.ComboBox1.SelectedText
        End Get
        Set(ByVal st As String)
            Me.ComboBox1.SelectedText = st
            SetColour()
        End Set
    End Property

    Public Property SelectedValue() As Object
        Get
            SelectedValue = Me.ComboBox1.SelectedValue
        End Get
        Set(ByVal sv As Object)
            Me.ComboBox1.SelectedValue = sv
            'If Me.ComboBox1.SelectedIndex <> -1 Then
            '    If Me.ComboBox1.SelectionLength = 0 Then
            '        Me.ComboBox1.SelectionLength = Len(Me.ComboBox1.Text)
            '    End If
            'End If
            SetColour()
        End Set
    End Property

    Public Overloads Property xTag() As Object
        Get
            xTag = Me.ComboBox1.Tag
        End Get
        Set(ByVal tg As Object)
            Me.ComboBox1.Tag = tg
        End Set
    End Property

    Public Overrides Property ContextMenu() As System.Windows.Forms.ContextMenu
        Get
            ContextMenu = Me.ComboBox1.ContextMenu
        End Get
        Set(ByVal cm As System.Windows.Forms.ContextMenu)
            Me.ComboBox1.ContextMenu = cm
        End Set
    End Property

    Public Sub SelectAll()
        Me.ComboBox1.SelectAll()
    End Sub

    Public ReadOnly Property items() As ComboBox.ObjectCollection
        Get
            items = Me.ComboBox1.Items
        End Get
    End Property

    Public ReadOnly Property ErrorMsg() As String
        Get
            If Me.ComboBox1.SelectedIndex = -1 And bRequired Then
                ErrorMsg = "Must not be empty"
            Else
                ErrorMsg = ""
            End If
        End Get
    End Property

    'Auto fill combo box input functionality...
    Private Sub ComboBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyUp
        Dim s As String
        Dim ss As String
        Dim i As Integer

        Try
            Select Case e.KeyCode
                Case Keys.Left, Keys.Right, Keys.Up, Keys.Delete, Keys.Down, _
                     Keys.Shift, Keys.ShiftKey, Keys.Tab, Keys.F1, Keys.Home, Keys.End
                Case Keys.Back
                    s = Me.ComboBox1.Text
                    If Len(s) > 1 Then
                        s = Mid(s, 1, Len(s) - 1)
                    Else
                        s = ""
                    End If
                Case Else
                    If Me.ComboBox1.Text <> "" And Me.ComboBox1.SelectionStart = Len(Me.ComboBox1.Text) Then
                        s = Me.ComboBox1.Text
                        i = Me.ComboBox1.FindString(Me.ComboBox1.Text)
                        If i > -1 Then
                            Me.ComboBox1.SelectedIndex = i
                            ss = Me.ComboBox1.Text
                            If Len(s) < Len(ss) Then
                                Me.ComboBox1.Select(Len(s), Len(ss) - Len(s))
                            Else
                                Me.ComboBox1.Select(Len(s), 1)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
        End Try
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics
        SetColour()
        Dim r As New Rectangle(e.ClipRectangle.X, e.ClipRectangle.Y, e.ClipRectangle.Width - 1, e.ClipRectangle.Height - 1)
        g.FillRectangle(New SolidBrush(Me.ComboBox1.BackColor), r)
        g.DrawRectangle(New Pen(borderColour, DialogStyle.BorderWidth), r)
    End Sub

    Private Sub SetColour()
        Dim bgColour, borColour As Color
        Dim b As Boolean = False

        If bRequired Then
            bgColour = DialogStyle.NameToColour(DialogStyle.BackRequired)
        Else
            bgColour = DialogStyle.NameToColour(DialogStyle.BackNormal)
        End If

        If Me.ComboBox1.SelectedIndex = -1 And bRequired Then
            borColour = DialogStyle.NameToColour(DialogStyle.BorderError)
        Else
            borColour = DialogStyle.NameToColour(DialogStyle.BorderNormal)
        End If

        If MyBase.BackColor <> bgColour Then
            MyBase.BackColor = bgColour
            Me.ComboBox1.BackColor = bgColour
            b = True
        End If
        If borderColour <> borColour Then
            borderColour = borColour
            b = True
        End If
        If b Then Me.Invalidate()
    End Sub

    Private Sub ComboBox1_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDownClosed
        Me.Invalidate()
    End Sub

    Private Sub ComboBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Leave
        SetColour()
    End Sub
End Class
