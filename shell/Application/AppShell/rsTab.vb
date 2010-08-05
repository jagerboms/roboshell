Imports System.Drawing.Drawing2D

Public Class rsTab
    Public Sub New()
        InitializeComponent()
        MyBase.BackColor = DialogStyle.NameToColour(DialogStyle.BorderNormal)
    End Sub

    Protected Overrides Sub Onresize(ByVal e As EventArgs)
        Dim i As Integer = DialogStyle.BorderWidth
        Me.ToolStrip1.Top = i
        Me.ToolStrip1.Left = i
    End Sub

    Public Overloads Property Height() As Integer
        Get
            Height = MyBase.Height
        End Get
        Set(ByVal h As Integer)
            Dim i As Integer = DialogStyle.BorderWidth
            Dim j As Integer
            If h < 50 Then
                j = 50
            Else
                j = h
            End If
            MyBase.Height = j
            Me.ToolStrip1.Top = i
            Me.ToolStrip1.Left = i
            Me.ToolStrip1.Width = MyBase.Width - 2 * i
        End Set
    End Property

    Public Overloads Property Width() As Integer
        Get
            Width = MyBase.Width
        End Get
        Set(ByVal w As Integer)
            Dim i As Integer = DialogStyle.BorderWidth
            Dim j As Integer

            If w < 50 Then
                j = 50
            Else
                j = w
            End If
            MyBase.Width = j
            Me.ToolStrip1.Top = i
            Me.ToolStrip1.Left = i
            Me.ToolStrip1.Width = MyBase.Width - 2 * i
        End Set
    End Property

    Public Function AddPanel(ByVal Name As String, ByVal Label As String) As Panel
        Dim p As New Panel
        Dim i As Integer = DialogStyle.BorderWidth
        Dim c, j As Integer
        Dim b As Boolean = True

        c = Me.Controls.Count
        With p
            j = Me.ToolStrip1.Top + Me.ToolStrip1.Height
            .Top = j
            .Left = i
            .Width = MyBase.Width - 2 * i
            .Height = MyBase.Height - i - j
            .BackColor = DialogStyle.NameToColour(DialogStyle.BackNormal)
            .Name = Name
        End With
        Me.Controls.Add(p)

        If Me.ToolStrip1.Items.Count > 0 Then
            Me.ToolStrip1.Items.Add(New ToolStripSeparator)
            p.Visible = False
            b = False
        End If

        Dim tsb As New ToolStripButton
        With tsb
            .Name = Name
            .DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
            .Text = Label
            .Tag = Name
            If b Then
                .BackColor = DialogStyle.NameToColour(DialogStyle.BackNormal)
            Else
                .BackColor = Color.Transparent
            End If
            AddHandler .Click, AddressOf ToolStripButton_Click
        End With
        Me.ToolStrip1.Items.Add(tsb)

        Return p
    End Function

    Public Sub SetError(ByVal PanelName As String, ByVal OnOff As Boolean)
        Dim tb As ToolStripButton

        tb = DirectCast(Me.ToolStrip1.Items(PanelName), ToolStripButton)
        If OnOff Then
            tb.ForeColor = DialogStyle.NameToColour(DialogStyle.ForeError)
        Else
            tb.ForeColor = DialogStyle.NameToColour(DialogStyle.ForeColour)
        End If
    End Sub

    Private Sub ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim tsb As ToolStripButton = DirectCast(sender, ToolStripButton)
        Dim tb As ToolStripButton
        Dim p As Panel
        Dim s As String = tsb.Tag.ToString

        For Each c As Control In Me.Controls
            If c.GetType.Name = "Panel" Then
                p = DirectCast(c, Panel)
                If c.Name = s Then
                    p.Visible = True
                    For Each cc As Control In p.Controls
                        If cc.TabStop = True Then
                            cc.Focus()
                            Exit For
                        End If
                    Next
                ElseIf p.Visible = True Then
                    p.Visible = False
                End If
            End If
        Next

        For Each c As ToolStripItem In Me.ToolStrip1.Items
            If c.GetType.Name = "ToolStripButton" Then
                tb = DirectCast(c, ToolStripButton)
                tb.BackColor = Color.Transparent
            End If
        Next
        tsb.BackColor = DialogStyle.NameToColour(DialogStyle.BackNormal)
    End Sub

    Private Sub ToolBar_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles ToolStrip1.Paint
        Dim tb As ToolStrip = DirectCast(sender, ToolStrip)
        Dim r As New Rectangle(0, 0, tb.Width, tb.Height)
        Dim Br As New LinearGradientBrush(r, DialogStyle.NameToColour(DialogStyle.ToolStart), DialogStyle.NameToColour(DialogStyle.ToolEnd), LinearGradientMode.Vertical)
        e.Graphics.FillRectangle(Br, e.ClipRectangle)
    End Sub
End Class
