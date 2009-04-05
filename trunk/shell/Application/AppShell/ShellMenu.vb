Option Explicit On 
Option Strict On

Public Class MenuDefn
    Inherits ObjectDefn

    Public Sub New(ByVal sName As String)
        MyBase.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New ShellMenu(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case Name
            'Case "Parameter"
            '    Parameter = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Menu object")
        End Select
    End Sub
End Class

Public Class ShellMenu
    Inherits ShellObject

    Private Actions As ActionDefns
    Public fForm As System.Windows.Forms.Form

    Public Event Action(ByVal sAction As String)
    Private ts As System.Windows.Forms.ToolStrip
    Private MyParams As New ShellParameters

    Public Sub New(ByVal defn As ObjectDefn)
        Actions = defn.Actions
        For Each a As ActionDefn In Actions
            If a.LinkedParam <> "" Then
                MyParams.Add(a.LinkedParam, Nothing, DbType.String, True, True, 100)
            End If
        Next
    End Sub

    Public Shadows ReadOnly Property parms() As ShellParameters
        Get
            Dim p As shellParameter
            Dim b As ToolStripButton
            Dim values() As String

            For Each a As ActionDefn In Actions
                If a.LinkedParam <> "" Then
                    p = MyParams.Item(a.LinkedParam)
                    values = Split(a.ParamValue, "||")
                    b = GetButton(a)
                    If Not b Is Nothing Then
                        If b.CheckState = CheckState.Unchecked Then
                            p.Value = values(0)
                        Else
                            p.Value = values(1)
                        End If
                    End If
                End If
            Next
            Return MyParams
        End Get
    End Property

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Dim p As shellParameter
            Dim b As ToolStripButton
            Dim chk As Boolean
            Dim values() As String

            MyParams.MergeValues(Parms)

            For Each a As ActionDefn In Actions
                If a.LinkedParam <> "" Then
                    p = MyParams.Item(a.LinkedParam)
                    values = Split(a.ParamValue, "||")
                    chk = a.Checked
                    If GetString(p.Value) = values(1) Then
                        a.Checked = True
                    Else
                        a.Checked = False
                    End If
                    If chk <> a.Checked and Not ts Is Nothing Then
                        b = GetButton(a)
                        If chk Then
                            b.CheckState = CheckState.Unchecked
                        Else
                            b.CheckState = CheckState.Checked
                        End If
                    End If
                End If
            Next

            If ts Is Nothing Then
                InitialiseStrip()
            End If
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                Me.Messages.Add("E", ex.ToString)
            Else
                Dim ex2 As Exception = ex.InnerException
                Do While Not ex2 Is Nothing
                    Me.Messages.Add("E", ex2.ToString)
                    ex2 = ex2.InnerException
                Loop
            End If
            Me.OnExitFail()
        End Try
    End Sub

    Public Overrides Sub Listener(ByVal Parms As ShellParameters)
    End Sub

    Public Overrides Sub Suspend(ByVal Mode As Boolean)
    End Sub

    Private Sub InitialiseStrip()
        ts = New System.Windows.Forms.ToolStrip
        Dim b As ToolStripButton

        For Each a As ActionDefn In Actions
            If a.IsButton And a.Parent = "" Then
                If a.MenuType = "S" Then
                    Dim ddb As New System.Windows.Forms.ToolStripDropDownButton
                    With ddb
                        .DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
                        .Name = a.Name
                        Try
                            .Image = Publics.GetImage(a.ImageFile)
                        Catch
                        End Try
                        .ImageTransparentColor = System.Drawing.Color.Magenta

                        .ToolTipText = a.ToolTip
                    End With

                    For Each sa As ActionDefn In Actions
                        If sa.Parent = a.Name Then
                            Dim si As New System.Windows.Forms.ToolStripMenuItem
                            si.Name = sa.Name
                            si.Text = sa.MenuText
                            si.Tag = sa
                            Try
                                si.Image = Publics.GetImage(sa.ImageFile)
                            Catch
                            End Try
                            AddHandler si.Click, AddressOf Menu_Click
                            ddb.DropDownItems.Add(si)
                        End If
                    Next

                    ts.Items.Add(ddb)
                Else
                    b = New ToolStripButton
                    With b
                        .DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
                        .Name = a.Name
                        Try
                            .Image = Publics.GetImage(a.ImageFile)
                        Catch
                        End Try
                        .ImageTransparentColor = System.Drawing.Color.Magenta

                        .ToolTipText = a.ToolTip
                        AddHandler .Click, AddressOf Button_Click
                        .Tag = a
                        If a.LinkedParam <> "" Then
                            .CheckOnClick = True
                            If a.Checked Then
                                .Checked = True
                            End If
                        End If
                    End With
                    ts.Items.Add(b)
                End If
            End If
        Next
        fForm.Controls.Add(ts)
    End Sub

    Public Sub Enable(ByVal Action As ActionDefn)
        If Action.IsButton Or Action.MenuType = "S" Or Action.MenuType = "I" Then
            For Each s As ToolStripItem In ts.Items
                If s.Name = Action.Name Then
                    s.Enabled = Action.Enabled
                End If
            Next
        End If
    End Sub

    Private Function GetButton(ByVal Action As ActionDefn) As ToolStripButton
        If Action.IsButton And Action.MenuType <> "S" Then
            For Each b As ToolStripItem In ts.Items
                If b.Name = Action.Name Then
                    Return CType(b, ToolStripButton)
                End If
            Next
        End If
        Return Nothing
    End Function

    Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim b As ToolStripButton = CType(sender, ToolStripButton)
        RaiseEvent Action(b.Name)
    End Sub

    Protected Sub Menu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim m As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        RaiseEvent Action(m.Name)
    End Sub
End Class
