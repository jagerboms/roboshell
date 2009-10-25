Option Explicit On 
Option Strict On

Public Class DialogForm
    Inherits System.Windows.Forms.Form
    Friend oOwner As Dialog
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
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents statusBar As System.Windows.Forms.StatusBar
    Friend WithEvents Message As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Stale As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.statusBar = New System.Windows.Forms.StatusBar
        Me.Message = New System.Windows.Forms.StatusBarPanel
        Me.Stale = New System.Windows.Forms.StatusBarPanel
        CType(Me.Message, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Stale, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Russell"
        '
        'ContextMenu2
        '
        '
        'statusBar
        '
        Me.statusBar.ContextMenu = Me.ContextMenu2
        Me.statusBar.Location = New System.Drawing.Point(0, 249)
        Me.statusBar.Name = "statusBar"
        Me.statusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.Message, Me.Stale})
        Me.statusBar.ShowPanels = True
        Me.statusBar.Size = New System.Drawing.Size(224, 28)
        Me.statusBar.TabIndex = 1
        '
        'Message
        '
        Me.Message.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.Message.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None
        Me.Message.Name = "Message"
        Me.Message.Width = 170
        '
        'Stale
        '
        Me.Stale.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.Stale.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None
        Me.Stale.Name = "Stale"
        Me.Stale.Width = 38
        '
        'DialogForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(224, 277)
        Me.ContextMenu = Me.ContextMenu2
        Me.Controls.Add(Me.statusBar)
        Me.ForeColor = System.Drawing.Color.Blue
        Me.KeyPreview = True
        Me.Name = "DialogForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Dialog"
        CType(Me.Message, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Stale, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Dialog_DoubleClick(ByVal sender As Object, _
                        ByVal e As System.EventArgs) Handles MyBase.DoubleClick
        If DblClkKey <> "" Then
            oOwner.ProcessAction(DblClkKey)
        End If
    End Sub

    Private Sub Dialog_KeyUp(ByVal sender As Object, _
                 ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        oOwner.ProcessKey(e.KeyCode, e.Modifiers.ToString)
    End Sub

    Private Sub Dialog_Load(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = Publics.ShellIcon
    End Sub

    Private Sub Dialog_Closing(ByVal sender As Object, _
             ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        oOwner.ProcessClose()
    End Sub

    Private Sub ContextMenu1_Popup(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles ContextMenu1.Popup
        oOwner.DoMenu(1)
    End Sub

    Private Sub ContextMenu2_Popup(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles ContextMenu2.Popup
        oOwner.DoMenu(2)
    End Sub

    Protected Overrides Function ProcessMnemonic(ByVal charCode As Char) As Boolean
        If CheckControls(Me, charCode) Then
            Return True
        End If
        Return MyBase.ProcessMnemonic(charCode)
    End Function

    Private Function CheckControls(ByVal ctrl As Control, ByVal charCode As Char) As Boolean
        For Each c As Control In ctrl.Controls
            If Mid(c.ToString.ToLower, 1, 32) = "system.windows.forms.tabcontrol," Then
                Dim t As System.Windows.Forms.TabControl = CType(c, System.Windows.Forms.TabControl)
                For Each tp As TabPage In t.TabPages
                    If IsMnemonic(charCode, tp.Text) Then
                        t.SelectedTab = tp
                        Return True
                    End If
                Next
                If CheckControls(t.SelectedTab, charCode) Then
                    Return True
                End If
            End If
            If CheckControls(c, charCode) Then
                Return True
            End If
        Next
        Return False
    End Function
End Class

Public Class DialogDefn
    Inherits ObjectDefn

    Private sTitle As String
    Private sTitleParameter() As String
    Private sHelpPage As String

    Public ReadOnly Property Title() As String
        Get
            Title = sTitle
        End Get
    End Property

    Public ReadOnly Property HelpPage() As String
        Get
            HelpPage = sHelpPage
        End Get
    End Property

    Public ReadOnly Property TitleParameter() As String()
        Get
            TitleParameter = sTitleParameter
        End Get
    End Property

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New Dialog(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "Title"
                sTitle = GetString(Value)
            Case "TitleParameters"
                sTitleParameter = Split(GetString(Value), "||")
            Case "HelpPage"
                sHelpPage = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Dialog object")
        End Select
    End Sub
End Class

Public Class Dialog
    Inherits ShellObject

    Private sDefn As DialogDefn
    Private fForm As DialogForm
    Private bLoading As Boolean = False
    Private bInit As Boolean = True
    Private bEditing As Boolean = True
    Private WithEvents mAction As ShellMenu
    Private ActiveField As String = ""
    Private FieldText As String
    Private FirstField As String
    Private LocalParms As New ShellParameters
    Private dlogf As New DialogFields
    Private bCloseState As Boolean = False
    Private sTitle As String
    Private TxtHeight As Integer

    Public Sub New(ByVal Defn As DialogDefn)
        Dim r As ObjectRegister
        Dim fld As Field
        Dim d As DialogField

        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
        MyBase.parms.Clone(LocalParms)

        For Each c As Field In sDefn.Fields
            fld = Nothing
            c.Clone(fld)
            d = dlogf.Add(fld)
            For Each f As Field In sDefn.Fields
                If f.LinkField = c.Name Then
                    d.AddLinkedField(c.Name)
                End If
            Next
        Next

        r = Register.Add(CType(Me, ShellObject))
        Me.RegKey = r.Key
        For Each p As ShellProperty In sDefn.Properties
            If p.Type = "lk" Then
                Register.Listen.Add(p.Name, "B", Me.RegKey)
            End If
        Next
    End Sub

    Public Shadows ReadOnly Property parms() As ShellParameters
        Get
            Dim p As shellParameter
            Dim d As DialogField
            Try
                If Not bLoading Then
                    For Each d In dlogf
                        p = MyBase.parms.Item(d.Name)
                        If Not p Is Nothing Then
                            If p.Output Then
                                p.Value = GetFieldValue(d.Name)
                                p.InputText = d.Text
                            End If
                        End If
                    Next
                End If
            Catch
            End Try
            Return MyBase.parms
        End Get
    End Property

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim s As String
        Dim b As Boolean
        Dim p As shellParameter
        Dim p2 As shellParameter
        Dim prop As ShellProperty
        Dim d As DialogField
        Dim obj As Object
        Dim a As ActionDefn
        Dim i As Integer

        Try
            MyBase.Parms.MergeValues(Parms)
            If fForm Is Nothing Then
                bInit = True
                LocalParms.MergeValues(Parms)
                fForm = New DialogForm
                fForm.oOwner = Me
                fForm.BackColor = Publics.GetBackColour()
                fForm.Name = sDefn.Title
                SetTitle()
                If Not Publics.MDIParent Is Nothing Then
                    fForm.MdiParent = Publics.MDIParent
                End If
                If Not InitialiseDialog() Then
                    ProcessClose()
                    fForm.Close()
                    Exit Sub
                End If
                InitialiseAction()
                Publics.SetFormPosition(CType(fForm, Form), sDefn.Name)
                For Each d In dlogf
                    If d.Field.DisplayType = "D" Then
                        CType(d.Control, ComboBox).SelectedIndex = -1
                        CType(d.Control, ComboBox).SelectedIndex = -1
                    End If
                Next

                fForm.Show()
                For Each d In dlogf
                    b = False
                    If Not LocalParms Is Nothing Then
                        If Not LocalParms.Item(d.Name) Is Nothing Then
                            If d.Field.DisplayType = "T" Or d.Field.DisplayType = "P" Then
                                s = LocalParms.Item(d.Name).InputText
                            Else
                                s = ""
                            End If
                            obj = LocalParms.Item(d.Name).Value
                            SetFieldValue(d.Name, obj, s)
                            b = Not LocalParms.Item(d.Name).Input
                        End If
                    End If

                    ' set user default properties

                    If Not b Then
                        If Not Parms Is Nothing Then
                            p = Parms.Item(d.Name)
                            If p Is Nothing Then
                                b = True
                            Else
                                If Not (p.Output And p.Initialised) Then
                                    b = True
                                End If
                            End If
                        End If
                    End If
                    If b Then
                        prop = sDefn.Properties.Item("_" & d.Name, "u")
                        If Not prop Is Nothing Then
                            obj = prop.Value
                            If Not obj Is Nothing Then
                                SetFieldValue(d.Name, obj, "")
                            End If
                        End If
                    End If
                    Application.DoEvents()
                Next
                bInit = False
                If FirstField <> "" Then
                    FieldEnter(FirstField)
                End If

                Dim hProcs() As String = Nothing ' Call each field based action but only
                Dim bDo As Boolean               ' once for each underlying process...
                For Each a In sDefn.Actions
                    If a.FieldName <> Nothing Then
                        If a.Enabled Then
                            bDo = True
                            If Not hProcs Is Nothing Then
                                For Each sP As String In hProcs
                                    If sP = a.Process Then
                                        bDo = False
                                        Exit For
                                    End If
                                Next
                            End If
                            If bDo Then
                                If DoProcess(a) Then
                                    If hProcs Is Nothing Then
                                        i = 0
                                    Else
                                        i = hProcs.GetUpperBound(0) + 1
                                    End If
                                    ReDim Preserve hProcs(i)
                                    hProcs(i) = a.Process
                                End If
                            End If
                        End If
                    End If
                Next
                SetActions()

                ProcessAction("Refresh")
            Else
                SetTitle()
                If Not Parms Is Nothing And Not bLoading Then
                    For Each p In Parms
                        If p.Output Then
                            p2 = MyBase.Parms.Item(p.Name)
                            If Not p2 Is Nothing Then
                                If p2.Input Then
                                    d = dlogf.Item(p.Name)
                                    If Not d Is Nothing Then
                                        If Not MyBase.Parms.Item(d.Name) Is Nothing Then
                                            obj = MyBase.Parms.Item(d.Name).Value
                                            p2 = LocalParms.Item(d.Name)
                                            If Not p2 Is Nothing Then
                                                p2.Value = obj
                                            End If
                                            Try
                                                d.Last = d.Value
                                            Catch ex As Exception
                                            End Try
                                            SetFieldValue(d.Name, obj, "")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    SetActions()
                End If
            End If
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Private Sub SetTitle()
        Dim s As String
        Dim ss As String
        Dim sTemp As String

        s = sDefn.Title
        If Not sDefn.TitleParameter Is Nothing Then
            For Each ss In sDefn.TitleParameter
                sTemp = GetString(MyBase.Parms.Item(ss).Value)
                If sTemp <> "" Then
                    s &= " - " & sTemp
                End If
            Next
        End If
        If s <> sTitle Then
            fForm.Text = s
            sTitle = s
        End If
    End Sub

    Public Overrides Sub Listener(ByVal Parms As ShellParameters)
        Dim b As Boolean
        Dim p2 As shellParameter
        Dim d As DialogField

        Try
            If Not Parms Is Nothing Then

                'match primary key columns to input parms 

                b = True
                For Each d In dlogf
                    If d.Field.Primary Then
                        Try
                            If GetString(GetFieldValue(d.Name)) <> _
                                            GetString(Parms.Item(d.Name).Value) Then
                                b = False
                                Exit For
                            End If
                        Catch
                            b = False
                            Exit For
                        End Try
                    End If
                Next

                ' if found update all the matching columns with data from parms

                If b Then
                    For Each p As shellParameter In Parms
                        p2 = MyBase.Parms.Item(p.Name)
                        If Not p2 Is Nothing Then
                            If Not MyBase.Parms.Item(p.Name) Is Nothing Then
                                MyBase.Parms.Item(p.Name).Value = p.Value
                            End If
                            d = dlogf.Item(p.Name)
                            If Not d Is Nothing Then
                                SetFieldValue(p.Name, p.Value, "")
                            End If
                        End If
                    Next
                End If
            End If
            SetActions()
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Public Overrides Sub Suspend(ByVal Mode As Boolean)
        If Mode Then
            fForm.Cursor = Cursors.WaitCursor
        Else
            fForm.Cursor = Cursors.Default
        End If
    End Sub

    Private Function InitialiseDialog() As Boolean
        Dim bRet As Boolean = True

        Try
            bLoading = True
            LoadContainer("", fForm.Controls)
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
            bRet = False
        End Try
        bLoading = False
        Return bRet
    End Function

    Private Function LoadContainer(ByVal sCont As String, ByVal ctrs As Control.ControlCollection) As Boolean
        Dim l As Label
        Dim i As Integer = 0
        Dim iH As Integer

        Dim cl As Integer
        Dim ct As Integer
        Dim cw As Integer
        Dim ch As Integer

        Dim gl As Integer = 5
        Dim gt As Integer
        Dim gw As Integer

        Dim fw As Integer = 100
        Dim ft As Integer = 30

        Dim ci As String = ""
        Dim d As DialogField

        If sCont = "" Then
            gt = 30
            ch = 30
        Else
            gt = 0
            ch = 0
        End If

        For Each f As Field In sDefn.Fields
            d = dlogf.Item(f.Name)
            If d.Field.DisplayType <> "REL" And d.Field.DisplayType <> "H" And d.Field.Container = sCont Then ' do nothing for hidden fields
                Select Case d.Field.Locate
                    Case "G"
                        gl = 5
                        gt = ft
                        gw = 0
                        ct = gt
                        ch = 0
                    Case "C"
                        gl += gw + 5
                        gw = 0
                        ct = gt
                        ch = 0
                End Select

                If UCase(d.Field.DisplayType) = "TAB" Then
                    Dim tb As New TabControl

                    cl = gl
                    ct += ch + 5
                    cw = d.Field.LabelWidth
                    If d.Field.DisplayHeight > 1 Then
                        iH = 17 + 13 * d.Field.DisplayHeight
                    Else
                        iH = 17
                    End If
                    With tb
                        .Name = d.Name
                        .Appearance = TabAppearance.FlatButtons
                        .Top = ct - 2
                        .Left = cl + cw
                        .Height = iH
                        .Width = d.Field.DisplayWidth
                    End With
                    cw += d.Field.DisplayWidth
                    ch = iH
                    AddHandler tb.DrawItem, AddressOf TabDrawItem
                    tb.DrawMode = TabDrawMode.OwnerDrawFixed
                    ctrs.Add(tb)
                    LoadContainer(d.Name, tb.Controls)

                ElseIf UCase(d.Field.DisplayType) = "TBP" Then
                    Dim s As String = ""
                    If d.Field.LinkField <> "" Then
                        Dim c2 As DialogField
                        c2 = dlogf.Item(d.Field.LinkField)
                        s = c2.Value.ToString
                    End If
                    If s = "" Then
                        s = d.Field.Label
                    End If
                    Dim tp As New TabPage(s)
                    With tp
                        .Name = d.Name
                        .AutoScroll = True
                        .BackColor = GetBackColour()
                        .BorderStyle = BorderStyle.None
                        .Margin = New System.Windows.Forms.Padding(0)
                    End With
                    ctrs.Add(tp)
                    LoadContainer(d.Name, tp.Controls)

                ElseIf d.Field.Locate <> "P" Or i = 0 Then
                    ct += ch + 5
                    cl = gl
                    cw = d.Field.LabelWidth
                    l = New Label
                    With l
                        .Text = d.Field.Label
                        .Top = ct - 1
                        Select Case d.Field.Justify
                            Case "L"
                                .TextAlign = ContentAlignment.MiddleLeft
                            Case "C"
                                .TextAlign = ContentAlignment.MiddleCenter
                            Case Else
                                .TextAlign = ContentAlignment.MiddleRight
                        End Select
                        .Left = cl
                        .Width = cw
                        .ForeColor = System.Drawing.Color.Blue
                        .Visible = True
                    End With
                    cw += 5
                    ctrs.Add(l)
                    d.AddLabel(l)
                    ci = d.Name
                    d.ErrField = ci
                    ch = l.Height
                Else
                    'ct = gt
                    cw += 2
                    d.ErrField = ci

                End If

                Select Case UCase(d.Field.DisplayType)
                    Case "L", "B"       ' Label

                        If d.Field.DisplayHeight > 1 Then
                            Dim t As New TextBox
                            With t
                                .Multiline = True
                                .ReadOnly = True
                                .ScrollBars = System.Windows.Forms.ScrollBars.None
                                .Cursor = System.Windows.Forms.Cursors.Default
                                .TabStop = False
                                .WordWrap = True
                                .BackColor = Publics.GetBackColour
                                Select Case d.Field.Justify
                                    Case "R"
                                        .TextAlign = HorizontalAlignment.Right
                                    Case "C"
                                        .TextAlign = HorizontalAlignment.Center
                                    Case Else
                                        .TextAlign = HorizontalAlignment.Left
                                End Select
                                If d.Field.DisplayType = "B" Then
                                    .BorderStyle = BorderStyle.FixedSingle
                                Else
                                    .BorderStyle = BorderStyle.None
                                End If
                            End With
                            TxtHeight = t.Height
                            iH = TxtHeight * d.Field.DisplayHeight
                            AddControl(d, CType(t, Control), ctrs, ct - 1, _
                                    cl + cw, d.Field.DisplayWidth, iH)
                            iH = t.Height + 3

                        Else
                            l = New Label
                            With l
                                Select Case d.Field.Justify
                                    Case "R"
                                        .TextAlign = ContentAlignment.MiddleRight
                                    Case "C"
                                        .TextAlign = ContentAlignment.MiddleCenter
                                    Case Else
                                        .TextAlign = ContentAlignment.MiddleLeft
                                End Select
                                If d.Field.DisplayType = "B" Then
                                    .BorderStyle = BorderStyle.Fixed3D
                                End If
                            End With
                            iH = -1
                            AddControl(d, CType(l, Control), ctrs, ct - 1, _
                                    cl + cw, d.Field.DisplayWidth, iH)
                            iH = l.Height
                        End If
                        cw += d.Field.DisplayWidth
                        If iH > ch Then ch = iH

                    Case "T", "P"            'Textbox, DatePicker
                        Dim t As New TextBox
                        With t
                            Select Case d.Field.Justify
                                Case "R"
                                    .TextAlign = HorizontalAlignment.Right
                                Case "C"
                                    .TextAlign = HorizontalAlignment.Center
                                Case Else
                                    .TextAlign = HorizontalAlignment.Left
                            End Select
                            If d.Field.Enabled Then
                                If d.Field.Required Then
                                    .BackColor = Color.FromArgb(255, 255, 100)
                                End If
                            Else
                                .ReadOnly = True
                            End If
                            If d.Field.DisplayHeight > 1 And UCase(d.Field.DisplayType) = "T" Then
                                .Multiline = True
                                iH = t.Height * d.Field.DisplayHeight
                            Else
                                iH = -1
                            End If
                            .BorderStyle = BorderStyle.FixedSingle
                        End With
                        AddControl(d, CType(t, Control), ctrs, ct, _
                                cl + cw, d.Field.DisplayWidth, iH)
                        cw += d.Field.DisplayWidth
                        iH = t.Height
                        If d.Field.DisplayHeight > 1 Then
                            iH += 3
                        End If
                        If iH > ch Then ch = iH

                        If UCase(d.Field.DisplayType) = "P" Then   'DatePicker
                            Dim dp As New DateTimePicker
                            With dp
                                .Top = ct + 1
                                .Left = cl + cw
                                .Width = 20
                                .TabStop = False
                                .DropDownAlign = LeftRightAlignment.Right
                                .Tag = d.Name
                                .Format = DateTimePickerFormat.Short

                                AddHandler .ValueChanged, AddressOf DatePickChanged
                                AddHandler .Enter, AddressOf DatePickEnter
                            End With
                            ctrs.Add(dp)
                            cw += 20
                        End If

                    Case "D"        'Dropdownlist
                        Dim cbox As New ComboBox
                        With cbox
                            'd.Sorted = True
                            If d.Field.Required Then
                                .BackColor = Color.FromArgb(255, 255, 100)
                            End If
                            .DropDownStyle = ComboBoxStyle.DropDown
                            AddHandler .KeyUp, AddressOf Combo_KeyUp
                            If d.Field.FillProcess <> "" Then
                                Dim p As New ShellProcess(d.Field.FillProcess, _
                                                                Me, Me.parms)
                                .DisplayMember = d.Field.TextField
                                .ValueMember = d.Field.ValueField
                                .DataSource = CType( _
                                    Me.parms.Item(d.Field.FillProcess).Value, DataTable)
                            Else
                                Dim ds As New ArrayList
                                Dim s As String
                                If d.Field.ValueField = "" Then
                                    s = "Y||Yes||N||No"
                                Else
                                    s = d.Field.ValueField
                                End If
                                Dim a() As String = Split(s, "||")
                                For j As Integer = 0 To a.GetUpperBound(0) Step 2
                                    ds.Add(New ComboSource(a(j), a(j + 1)))
                                Next
                                .DataSource = ds
                                .DisplayMember = "Text"
                                .ValueMember = "Value"
                            End If
                        End With
                        AddControl(d, CType(cbox, Control), ctrs, ct, _
                              cl + cw, d.Field.DisplayWidth, -1)
                        cw += d.Field.DisplayWidth
                        AddHandler cbox.SelectedIndexChanged, AddressOf Field_Change
                        If cbox.Height > ch Then ch = cbox.Height

                    Case "LST"            'Listbox
                        Dim lbox As New ListBox
                        'With lbox
                        '    If d.Field.Required Then
                        '        .BackColor = Color.FromArgb(255, 255, 100)
                        '    End If
                        'End With
                        If d.Field.DisplayHeight > 1 Then
                            iH = 17 + 13 * d.Field.DisplayHeight
                        Else
                            iH = 17
                        End If
                        AddControl(d, CType(lbox, Control), ctrs, ct, _
                              cl + cw, d.Field.DisplayWidth, iH)
                        cw += d.Field.DisplayWidth
                        AddHandler lbox.SelectedIndexChanged, AddressOf Field_Change
                        If lbox.Height > ch Then ch = lbox.Height

                    Case "CHK", "C"            'Checkbox
                        Dim cb As New CheckBox
                        AddControl(d, CType(cb, Control), ctrs, ct - 2, _
                                          cl + cw, 15, -1)
                        cw += 15
                        AddHandler cb.CheckedChanged, AddressOf Field_Change
                        If ch = 0 Then ch = 23

                End Select
                If ct + ch > ft Then ft = ct + ch
                If cw > gw Then gw = cw
                If cl + cw > fw Then fw = cl + cw
                i += 1
            End If
        Next
        If sCont = "" Then
            fForm.Height = ft + 70
            fForm.Width = fw + 35
        End If
    End Function

    Private Sub AddControl(ByRef d As DialogField, ByRef ctl As Control, _
                ByVal ctrs As Control.ControlCollection, _
                ByVal top As Integer, ByVal left As Integer, _
                ByVal width As Integer, ByVal Height As Integer)

        With ctl
            .Enabled = d.Field.Enabled
            AddHandler .Enter, AddressOf Field_Enter
            AddHandler .Leave, AddressOf Field_Leave
            AddHandler .TextChanged, AddressOf Field_Change
            .ForeColor = System.Drawing.Color.Blue
            .Top = top
            .Left = left
            .Width = width
            .Visible = True
            If Height > 0 Then
                .Height = Height
            End If
            If d.Field.Enabled Then
                .ContextMenu = fForm.ContextMenu1
            End If

            For Each a As ActionDefn In sDefn.Actions
                If a.FieldName = d.Name Then
                    d.AddAction(a.Name)
                End If
            Next
            .Name = d.Name
            .Tag = d.Name
        End With
        ctrs.Add(ctl)
        d.AddControl(ctl)
    End Sub

    Private Sub DatePickEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dp As DateTimePicker
        Dim s As String
        Dim dt As Date
        Dim d As DialogField
        Dim t As TextBox

        dp = DirectCast(sender, DateTimePicker)
        s = CType(dp.Tag, String)
        d = dlogf.Item(s)
        t = DirectCast(d.Control, TextBox)
        s = t.Text
        If IsDate(s) Then
            dt = CDate(s)
            dp.Value = dt
        End If
    End Sub

    Private Sub DatePickChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dp As DateTimePicker
        Dim s As String
        Dim d As DialogField
        Dim t As TextBox

        dp = DirectCast(sender, DateTimePicker)
        s = CType(dp.Tag, String)
        d = dlogf.Item(s)
        t = DirectCast(d.Control, TextBox)
        If d.Field.Format = "" Then
            t.Text = Format(dp.Value, "d-MMM-yyyy")
        Else
            t.Text = Format(dp.Value, d.Field.Format)
        End If
        SetFieldText(s)
    End Sub

    Private Sub TabDrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        Dim tc As System.Windows.Forms.TabControl = DirectCast(sender, System.Windows.Forms.TabControl)
        Dim sf As New StringFormat
        'Dim r As Rectangle
        'Dim rf As RectangleF

        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.HotkeyPrefix = Drawing.Text.HotkeyPrefix.Show
        If tc.SelectedIndex = e.Index Then
            e.Graphics.FillRectangle(New SolidBrush(GetBackColour()), e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height)

            'r = tc.GetTabRect(tc.TabPages.Count - 1)
            'rf = New RectangleF(r.X + r.Width, r.Y - 5, _
            '  tc.Width - (r.X + r.Width), r.Height + 5)

            'e.Graphics.FillRectangle(New SolidBrush(tc.Parent.BackColor), rf)
            'Else
            '    e.Graphics.FillRectangle(Brushes.Gainsboro, e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height)
        End If
        e.Graphics.DrawString(tc.TabPages(e.Index).Text, tc.Font, Brushes.Blue, RectangleF.op_Implicit(e.Bounds), sf)
        sf.Dispose()
    End Sub

    'Auto fill combo box input functionality...

    Private Sub Combo_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
        Dim s As String
        Dim ss As String
        Dim i As Integer

        Try
            Dim c As ComboBox = CType(sender, ComboBox)

            Select Case e.KeyCode
                Case Keys.Left, Keys.Right, Keys.Up, Keys.Delete, Keys.Down, _
                     Keys.Shift, Keys.ShiftKey, Keys.Tab, Keys.F1, Keys.Home, Keys.End
                Case Keys.Back
                    s = c.Text
                    If Len(s) > 1 Then
                        s = Mid(s, 1, Len(s) - 1)
                    Else
                        s = ""
                    End If
                Case Else
                    If c.Text <> "" And c.SelectionStart = Len(c.Text) Then
                        s = c.Text
                        i = c.FindString(c.Text)
                        If i > -1 Then
                            c.SelectedIndex = i
                            ss = c.Text
                            If Len(s) < Len(ss) Then
                                c.Select(Len(s), Len(ss) - Len(s))
                            Else
                                c.Select(Len(s), 1)
                            End If
                        End If
                    End If
            End Select
            'ProcessKey(e.KeyCode, e.Modifiers.ToString)

        Catch ex As Exception
        End Try
    End Sub

    Private Sub Field_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim s As String = CType(CType(sender, Control).Tag, String)
        If bInit Then
            If FirstField Is Nothing Then
                FirstField = s
            End If
        Else
            FieldEnter(s)
        End If
    End Sub

    Private Sub FieldEnter(ByVal sField As String)
        Dim d As DialogField = dlogf.Item(sField)
        Dim c As Control
        fForm.statusBar.Panels(0).Text = d.Field.HelpText
        d.Last = d.Value
        ActiveField = d.Name

        c = d.Control
        If Not c Is Nothing Then
            If c.GetType.Name = "TextBox" Then
                Dim t As TextBox = DirectCast(c, TextBox)
                t.Text = d.Text
                t.SelectionStart = 0
                t.SelectionLength = Len(t.Text)
            End If
        End If
        bEditing = True
    End Sub

    Private Sub Field_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim d As DialogField = dlogf.Item(CType(CType(sender, Control).Tag, String))
        ActiveField = ""
        fForm.statusBar.Panels(0).Text = ""
        SetFieldText(d.Name)
        bEditing = False
        fForm.statusBar.Panels(1).Style = StatusBarPanelStyle.Text
        fForm.statusBar.Panels(1).Text = ""
    End Sub

    Private Sub FieldLeave(ByVal sName As String)
        ActiveField = ""
        fForm.statusBar.Panels(0).Text = ""
        SetFieldText(sName)
        bEditing = False
        fForm.statusBar.Panels(1).Style = StatusBarPanelStyle.Text
        fForm.statusBar.Panels(1).Text = ""
    End Sub

    Private Sub Field_Change(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim d As DialogField = dlogf.Item(CType(CType(sender, Control).Tag, String))
        If ActiveField = d.Name Then
            If Not d.Actions Is Nothing And bEditing Then
                fForm.statusBar.Panels(1).Style = StatusBarPanelStyle.OwnerDraw
                fForm.statusBar.Panels(1).Text = "Stale"
            End If
        End If
    End Sub

    Private Function ValidateFields() As Boolean
        Dim bRet As Boolean = False
        Dim d As DialogField

        Try
            For Each d In dlogf
                If Not d.Errs Is Nothing Then
                    If d.Errs.count > 0 Then
                        Return False
                    End If
                End If
            Next
            bRet = True
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
        Return bRet
    End Function

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
            If ActiveField <> "" Then
                FieldLeave(ActiveField)
            End If
            Me.ProcessAction(sAction)
        End If
    End Sub

    Friend Sub SetActions()
        Dim bSet As Boolean
        Dim bRule As Boolean
        Dim vRule As String
        Dim d As DialogField

        ' Scan through the Actions and Enable / Disable them for the selected row

        For Each a As ActionDefn In sDefn.Actions

            If GetActionProcess(a) = "" Then
                bSet = False
            ElseIf Not a.Rules Is Nothing Then      ' rules always enabled
                bSet = True
                For Each cRuleD As ActionRuleDefn In a.Rules
                    bRule = True
                    For Each cRule As ActionRule In cRuleD.Rules
                        vRule = (GetString(GetFieldValue(cRule.FieldName)))
                        Select Case cRule.Type
                            Case ActionRule.ValidationType.EQ
                                If vRule <> cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.NE
                                If vRule = cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.NN
                                If vRule = "" Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.GT
                                If vRule <= cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.GE
                                If vRule < cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.LT
                                If vRule >= cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.LE
                                If vRule > cRule.Value.ToString Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.VL
                                d = dlogf.Item(cRule.FieldName)
                                If d.Errs.count > 0 Then
                                    bRule = False
                                    Exit For
                                End If
                            Case ActionRule.ValidationType.NV
                                d = dlogf.Item(cRule.FieldName)
                                If d.Errs.count = 0 Then
                                    bRule = False
                                    Exit For
                                End If
                        End Select
                    Next
                    If Not bRule Then
                        bSet = False
                        Exit For
                    End If
                Next
            Else
                bSet = True
            End If

            a.Enabled = bSet
            mAction.Enable(a)
        Next
    End Sub

    Friend Sub ProcessAction(ByVal sKey As String)
        Dim a As ActionDefn
        If Not sDefn.Actions Is Nothing Then
            a = sDefn.Actions.Item(sKey)
            If Not a Is Nothing Then
                If a.Enabled Then
                    DoProcess(a)
                End If
            End If
        End If
    End Sub

    Friend Sub ProcessKey(ByVal KeyCode As Integer, ByVal Shift As String)
        Dim b As Boolean = False
        If KeyCode = Keys.F1 Then
            If Shift = "Shift" Then
                b = True
            End If
            Publics.RaiseHelp(b, sDefn.HelpPage)
        Else
            ' Don't process a return key if currently entering a multi-line input.
            If KeyCode = 13 And Shift = "None" Then
                If ActiveField <> "" Then
                    If dlogf.Item(ActiveField).Field.DisplayHeight > 1 Then
                        Exit Sub
                    End If
                End If
            End If

            For Each a As ActionDefn In sDefn.Actions
                If a.KeyCode = KeyCode And a.IsKey Then
                    If a.Shift = Shift Or a.Shift Is Nothing Then
                        If ActiveField <> "" Then
                            SetFieldText(ActiveField)
                        End If
                        If a.Enabled Then
                            DoProcess(a)
                        End If
                        Exit Sub
                    End If
                End If
            Next
        End If
    End Sub

    Friend Sub ProcessClose()
        SaveFormPosition(CType(fForm, Form), sDefn.Name)
        Register.Remove(Me.RegKey)
        fForm.Hide()

        If bCloseState Then
            Me.OnExitOkay()
        Else
            Me.OnExitFail()
        End If
    End Sub

    Private Function DoProcess(ByRef a As ActionDefn) As Boolean
        Dim bValid As Boolean = ValidateFields()

        ' If all input is not valid and action requires validation then don't run it

        If a.Validate And Not bValid Then
            Return False
        End If

        Select Case a.CloseObject
            Case "P"
                If MsgBox("Are you sure you want to exit?", _
                        MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = _
                                                    MsgBoxResult.No Then
                    Return False
                End If
            Case "Q"
                If MsgBox("Are you sure you want to exit?", _
                        MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = _
                                                    MsgBoxResult.No Then
                    Return False
                End If
            Case Else
        End Select

        Dim p As New ShellProcess(GetActionProcess(a), Me, Me.parms)
        If p.Success Then
            Select Case a.CloseObject
                Case "Y", "P"
                    fForm.Close()
                Case "O", "Q"
                    bCloseState = True
                    fForm.Close()
                Case Else
            End Select
        End If
        Return True
    End Function

    Private Function GetActionProcess(ByVal Action As ActionDefn) As String
        Dim s As String
        Dim p As ProcessDefn

        If Action.ProcessField <> "" And Not Action.Processes Is Nothing Then
            s = GetString(GetFieldValue(Action.ProcessField))
            For Each cRule As ActionProcessRuleDefn In Action.Processes
                If s = GetString(cRule.Value) Then
                    p = Processes.Item(cRule.Process)
                    If Not p Is Nothing Then
                        Return cRule.Process
                    End If
                End If
            Next
        End If
        Return Action.Process
    End Function

    Private Function GetFieldValue(ByVal Field As String) As Object
        Dim c As DialogField

        Try
            c = dlogf.Item(Field)
            Select Case c.Field.DisplayType
                Case "REL", "H"    'Related or Hidden field return parameter value if it exists
                    Dim c2 As DialogField
                    Dim i As Integer
                    Dim ds As Object
                    Dim o As Object

                    If c.Field.LinkField = "" Then
                        If MyBase.parms.Item(Field) Is Nothing Then
                            Return Nothing
                        Else
                            Return MyBase.parms.Item(Field).Value
                        End If
                    End If

                    c2 = dlogf.Item(c.Field.LinkField)
                    Select Case c2.Field.DisplayType
                        Case "D"
                            Dim cb As ComboBox = DirectCast(c2.Control, ComboBox)
                            i = cb.SelectedIndex
                            ds = cb.DataSource

                        Case "LST"
                            Dim lb As ListBox = DirectCast(c2.Control, ListBox)
                            i = lb.SelectedIndex
                            ds = lb.DataSource

                        Case Else
                            Return c2.Value
                    End Select
                    If i = -1 Then
                        Return Nothing
                    End If
                    Select Case LCase(ds.GetType.ToString)
                        Case "system.data.datatable", "datatable"
                            o = CType(ds, DataTable).Rows.Item(i)(c.Field.LinkColumn)

                        Case "system.collections.arraylist", "arraylist"
                            Select Case c.Field.LinkColumn
                                Case "Text"
                                    o = CType(CType(ds, ArrayList).Item(i), ComboSource).Text
                                Case "Value"
                                    o = CType(CType(ds, ArrayList).Item(i), ComboSource).Value
                                Case Else
                                    Return Nothing
                            End Select

                        Case Else
                            Return Nothing
                    End Select
                    If IsDBNull(o) Then
                        Return Nothing
                    Else
                        Return o
                    End If

                    'Case "TBP"
                    '    Dim i As Integer
                    '    i = CType(dlogf.Item(c.Name).Control, TabPage).
                    '    i = CType(dlogf.Item(c.Field.Container).Control, TabControl).SelectedIndex
                    '    If i = c.ControlIndex Then
                    '        Return "Y"
                    '    Else
                    '        Return "N"
                    '    End If

                Case Else
                    Return c.Value
            End Select
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub SetFieldText(ByVal Field As String)
        Dim s As String
        Dim d As DialogField = dlogf.Item(Field)

        Try
            d.Text = d.Control.Text

            If GetString(d.Text) = "" And d.Field.DisplayType <> "C" And d.Field.DisplayType <> "CHK" Then
                d.Value = Nothing
            Else
                d.Value = d.Text
                Select Case d.Field.DisplayType
                    Case "D"            'DropdownList
                        Dim cb As ComboBox = DirectCast(d.Control, ComboBox)
                        If cb.SelectedIndex = -1 Then
                            d.Value = Nothing
                        Else
                            d.Value = cb.SelectedValue
                        End If

                        'Case "LST"            'Listbox
                        '    Dim lb As ListBox = directcast(d.Control, ListBox)
                        '    If lb.SelectedIndex = -1 Then
                        '        d.Value = Nothing
                        '    Else
                        '        d.Value = lb.SelectedValue
                        '    End If

                    Case "CHK", "C"            'Checkbox
                        If Len(d.Field.ValueField) > 1 Then
                            s = UCase(d.Field.ValueField)
                        Else
                            s = "YN"
                        End If
                        If DirectCast(d.Control, CheckBox).Checked Then
                            d.Value = Mid(s, 1, 1)
                        Else
                            d.Value = Mid(s, 2, 1)
                        End If

                    Case "T", "P", "L", "B"
                        s = d.Control.Text
                        d.Value = s
                        d.BusDateRelated = False
                        If s <> "" Then
                            Select Case d.Field.ValueType
                                Case DbType.Date, DbType.DateTime
                                    Dim dt As System.DateTime
                                    If UCase(s) = "T" _
                                        And Publics.BusinessDate <> Nothing Then
                                        d.Value = Publics.BusinessDate
                                        d.BusDateRelated = True
                                    Else
                                        Try
                                            If IsNumeric(s) Then
                                                If (Mid(s, 1, 1) = "+" _
                                                Or Mid(s, 1, 1) = "-") _
                                                And Publics.BusinessDate <> Nothing Then
                                                    dt = DateAdd(DateInterval.Day, _
                                                        CDbl(s), Publics.BusinessDate)
                                                Else
                                                    Select Case Len(s)
                                                        Case 1, 2
                                                            Dim i As Integer
                                                            i = CInt(s) - _
                                                            Publics.BusinessDate.Day()
                                                            If i = 0 Then
                                                                dt = Publics.BusinessDate
                                                            Else
                                                                dt = DateAdd( _
                                                                DateInterval.Day, _
                                                                i, Publics.BusinessDate)
                                                            End If
                                                        Case 4
                                                            dt = CType(Mid(s, 1, 2) & _
                                                                "-" & Mid(s, 3, 2), Date)
                                                        Case 6, 8
                                                            dt = CType(Mid(s, 1, 2) & _
                                                                "-" & Mid(s, 3, 2) & _
                                                                "-" & Mid(s, 5, 4), Date)
                                                    End Select
                                                End If
                                            Else
                                                dt = CType(s, Date)
                                            End If
                                        Catch
                                        End Try
                                        d.Value = dt
                                    End If

                                Case DbType.Currency
                                    Dim dc As Decimal
                                    Try
                                        dc = CType(s, Decimal)
                                        d.Value = dc
                                    Catch
                                    End Try

                                Case DbType.Double
                                    Dim db As Double
                                    Try
                                        db = CType(s, Double)
                                        d.Value = db
                                    Catch
                                    End Try

                                Case DbType.Int32
                                    Dim i2 As Int32
                                    Try
                                        i2 = CType(s, Integer)
                                        d.Value = i2
                                    Catch
                                    End Try

                                Case DbType.Int64
                                    Dim i4 As Int64
                                    Try
                                        i4 = CType(s, Int64)
                                        d.Value = i4
                                    Catch
                                    End Try
                            End Select
                        End If
                    Case Else
                        d.Value = d.Control.Text
                End Select
            End If
            CheckField(Field)
        Catch ex As Exception
            Dim i As Integer = 9
        End Try
    End Sub

    Private Sub CheckField(ByVal Field As String)
        Dim s As String
        Dim a As ActionDefn
        Dim d As DialogField = dlogf.Item(Field)
        Dim con As Control = d.Control
        Dim sComp1 As String
        Dim sComp2 As String = ""
        Dim b As Boolean

        s = GetString(d.Value)

        If Not d.LinkedFields Is Nothing Then
            Dim dx As DialogField
            Dim cb As ComboBox
            Dim dt As DataTable
            Dim save As String
            For Each ss As String In d.LinkedFields
                dx = dlogf.Item(ss)
                If Not dx Is Nothing Then
                    Select Case dx.Field.DisplayType
                        Case "D"   ' dropdown list
                            If dx.Field.FillProcess <> "" Then
                                If dx.Field.LinkColumn <> "" Then
                                    cb = DirectCast(dx.Control, ComboBox)
                                    save = cb.Text
                                    dt = CType(cb.DataSource, DataTable)
                                    dt.DefaultView.RowFilter = dx.Field.LinkColumn & " = '" & s & "'"
                                    cb.Text = save
                                    If cb.Text <> save Or cb.SelectedIndex = -1 Then
                                        cb.Text = ""
                                        cb.SelectedIndex = -1
                                        cb.SelectedIndex = -1
                                    End If
                                End If
                            Else
                                Dim i As Integer
                                i = 9 ' link to user data not yet supported!!
                            End If

                        Case "REL", "H"
                            dx.Value = GetFieldValue(ss)
                            If ss <> Field Then
                                CheckField(ss)
                            End If
                    End Select
                End If
            Next
        End If

        If d.Field.DisplayType <> "REL" And d.Field.DisplayType <> "H" Then
            d.Errs.Clear()

            If s = "" Then
                If d.Field.Required Then
                    d.Errs.Add("R", "Must not be empty")
                End If
                If d.BusDateRelated Then
                    If d.Field.Required = True Then
                        con.BackColor = Color.FromArgb(255, 255, 100)
                    Else
                        con.BackColor = Color.White
                    End If
                    d.BusDateRelated = False
                End If

            Else
                If d.Field.DisplayType = "T" Or d.Field.DisplayType = "P" _
                    Or d.Field.DisplayType = "L" Or d.Field.DisplayType = "B" Then
                    Select Case d.Field.ValueType
                        Case DbType.Date, DbType.DateTime
                            Dim dt As System.DateTime
                            Try
                                dt = CType(d.Value, Date)
                                If dt = DateTime.MinValue Then
                                    d.Errs.Add("V", "Invalid date")
                                End If
                            Catch
                                d.Errs.Add("V", "Invalid date")
                            End Try
                            If d.Errs.count = 0 Then
                                If d.Field.Format = "" Then
                                    con.Text = Format(dt, "d-MMM-yyyy")
                                Else
                                    con.Text = Format(dt, d.Field.Format)
                                End If
                            End If
                            If (d.Field.DisplayType = "T" Or d.Field.DisplayType = "P") _
                              And d.Field.Enabled Then
                                If d.BusDateRelated And _
                                    Format(Publics.BusinessDate, "yyyyMMdd") <> Format(Now(), "yyyyMMdd") Then
                                    con.BackColor = Color.GreenYellow
                                Else
                                    If d.Field.Required = True Then
                                        con.BackColor = Color.FromArgb(255, 255, 100)
                                    Else
                                        con.BackColor = Color.White
                                    End If
                                End If
                            End If

                        Case DbType.Currency
                            Dim dc As Decimal
                            Try
                                dc = CType(d.Value, Decimal)
                            Catch
                                d.Errs.Add("V", "Invalid numeric value")
                            End Try
                            If d.Errs.count = 0 Then
                                dc = CType(d.Text, Decimal)
                                con.Text = Format(dc, d.Field.Format)
                            End If

                        Case DbType.Double
                            Dim db As Double
                            Try
                                db = CType(d.Value, Double)
                            Catch
                                d.Errs.Add("V", "Invalid numeric value")
                            End Try
                            If d.Errs.count = 0 Then
                                db = CType(d.Text, Double)
                                con.Text = Format(db, d.Field.Format)
                            End If

                        Case DbType.Int32
                            Dim i2 As Int32
                            Try
                                i2 = CType(d.Value, Integer)
                            Catch
                                d.Errs.Add("V", "Invalid integer value")
                            End Try
                            If d.Errs.count = 0 Then
                                i2 = CType(d.Text, Integer)
                                con.Text = Format(i2, d.Field.Format)
                            End If

                        Case DbType.Int64
                            Dim i4 As Int64
                            Try
                                i4 = CType(d.Value, Int64)
                            Catch
                                d.Errs.Add("V", "Invalid integer value")
                            End Try
                            If d.Errs.count = 0 Then
                                i4 = CType(d.Text, Int64)
                                con.Text = Format(i4, d.Field.Format)
                            End If

                        Case DbType.String
                            If Len(d.Text) > d.Field.Width Then
                                d.Errs.Add("V", "Too many characters")
                            End If
                            If d.Errs.count = 0 Then
                                con.Text = d.Text
                            End If

                        Case Else
                            con.Text = d.Text
                    End Select
                End If

                For Each vv As ValidationDefn In sDefn.Validations
                    If vv.FieldName = Field Then
                        d.Errs.Remove("-" & vv.Name)
                        sComp1 = s
                        Select Case vv.ValueType
                            Case ValidationDefn.ValType.Process
                                Dim p As New ShellProcess(vv.Process, Me, Me.parms)
                                sComp1 = GetString(Me.parms.Item(vv.ReturnParameter).Value)
                                sComp2 = vv.Value.ToString
                            Case ValidationDefn.ValType.Constant
                                sComp2 = vv.Value.ToString
                            Case ValidationDefn.ValType.Field
                                sComp2 = GetString(GetFieldValue(vv.Value.ToString))
                        End Select
                        b = True
                        Select Case vv.Type
                            Case ValidationDefn.ValidationType.EQ
                                If sComp1 = sComp2 Then
                                    b = False
                                End If
                            Case ValidationDefn.ValidationType.NE
                                If sComp1 <> sComp2 Then
                                    b = False
                                End If
                            Case ValidationDefn.ValidationType.GT
                                If sComp1 > sComp2 Then
                                    b = False
                                End If
                            Case ValidationDefn.ValidationType.GE
                                If sComp1 >= sComp2 Then
                                    b = False
                                End If
                            Case ValidationDefn.ValidationType.LT
                                If sComp1 < sComp2 Then
                                    b = False
                                End If
                            Case ValidationDefn.ValidationType.LE
                                If sComp1 <= sComp2 Then
                                    b = False
                                End If
                        End Select
                        If Not b Then
                            d.Errs.Add("-" & vv.Name, vv.Message)
                        End If
                    ElseIf Not vv.AssociatedFields Is Nothing Then
                        For Each ss As String In vv.AssociatedFields
                            If ss = Field Then
                                Dim dff As DialogField = dlogf.Item(vv.FieldName)
                                dff.Errs.Remove("-" & vv.Name)
                                sComp1 = GetString(dff.Value)
                                If sComp1 <> "" Then
                                    Select Case vv.ValueType
                                        Case ValidationDefn.ValType.Process
                                            Dim p As New ShellProcess(vv.Process, Me, Me.parms)
                                            sComp1 = GetString(Me.parms.Item( _
                                                                vv.ReturnParameter).Value)
                                            sComp2 = vv.Value.ToString
                                        Case ValidationDefn.ValType.Constant
                                            sComp2 = vv.Value.ToString
                                        Case ValidationDefn.ValType.Field
                                            sComp2 = GetString(GetFieldValue( _
                                                                vv.Value.ToString))
                                    End Select
                                    b = True
                                    Select Case vv.Type
                                        Case ValidationDefn.ValidationType.EQ
                                            If sComp1 = sComp2 Then
                                                b = False
                                            End If
                                        Case ValidationDefn.ValidationType.NE
                                            If sComp1 <> sComp2 Then
                                                b = False
                                            End If
                                        Case ValidationDefn.ValidationType.GT
                                            If sComp1 > sComp2 Then
                                                b = False
                                            End If
                                        Case ValidationDefn.ValidationType.GE
                                            If sComp1 >= sComp2 Then
                                                b = False
                                            End If
                                        Case ValidationDefn.ValidationType.LT
                                            If sComp1 < sComp2 Then
                                                b = False
                                            End If
                                        Case ValidationDefn.ValidationType.LE
                                            If sComp1 <= sComp2 Then
                                                b = False
                                            End If
                                    End Select
                                    If Not b Then
                                        dff.Errs.Add("-" & vv.Name, vv.Message)
                                    End If
                                End If
                                SetErrorState(dff)
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If

            SetErrorState(d)
        End If

        SetActions()
        If Not GetString(d.Last) = GetString(d.Value) Then   ' data has changed
            If Not d.Actions Is Nothing And Not bInit Then
                For Each s In d.Actions
                    a = sDefn.Actions.Item(s)
                    If Not a Is Nothing Then
                        If a.Enabled Then
                            DoProcess(a)
                        End If
                    End If
                Next
            End If
            d.Last = d.Value
        End If
    End Sub

    Private Sub SetErrorState(ByVal d As DialogField)
        Dim b As Boolean
        Dim sField As String = d.ErrField
        Dim dErr As DialogField = dlogf.Item(sField)
        Dim lab As Label = dErr.Label()
        Dim con As Control = d.Control

        If d.Errs.count = 0 Then
            If fForm.ToolTip1.GetToolTip(con) <> "" Then
                fForm.ToolTip1.SetToolTip(con, "")
            End If

            b = True
            For Each dd As DialogField In dlogf
                If dd.ErrField = sField And dd.Errs.count > 0 Then
                    b = False
                    Exit For
                End If
            Next
            If b Then
                lab.ForeColor = Color.Blue
            Else
                lab.ForeColor = Color.Red
            End If
        Else
            lab.ForeColor = Color.Red
            fForm.ToolTip1.SetToolTip(con, d.Errs.Message)
        End If
    End Sub


    Private Sub SetFieldValue(ByVal Field As String, ByVal Value As Object, _
                                                            ByVal sText As String)
        Try
            Dim d As DialogField = dlogf.Item(Field)
            If d Is Nothing Then
                Exit Sub
            End If
            If d.Field.DisplayType = "H" Then
                ' Hidden field do Nothing
                Exit Sub
            End If

            Dim cc As Control = d.Control

            If sText <> "" Then
                d.Text = sText
            Else
                d.Text = GetString(Value)
            End If
            If d.Text = "" Then
                d.Value = Nothing
            Else
                d.Value = Value
            End If
            Select Case d.Field.DisplayType
                Case "D"            'Dropdownlist
                    If GetString(Value) = "" Then
                        CType(cc, ComboBox).SelectedIndex = -1
                        d.Value = Nothing
                    Else
                        Dim cb As ComboBox = CType(cc, ComboBox)
                        cb.SelectedValue = GetString(Value)
                        If cb.SelectedIndex <> -1 And Field = FirstField Then
                            If cb.SelectionLength = 0 Then
                                cb.SelectionLength = Len(cb.Text)
                            End If
                        End If
                    End If

                Case "LST"            'Listbox
                    CType(cc, ListBox).DataSource = Value
                    d.Value = Value

                Case "CHK", "C"            'Checkbox
                    Dim so As String = "YN"

                    If Len(d.Field.ValueField) > 1 Then
                        so = UCase(d.Field.ValueField)
                    End If
                    If Mid(so, 1, 1) = UCase(GetString(Value)) Then
                        d.Value = Mid(so, 1, 1)
                        CType(cc, CheckBox).Checked = True
                    Else
                        d.Value = Mid(so, 2, 1)
                        CType(cc, CheckBox).Checked = False
                    End If

                Case "T", "P", "L", "B"
                    d.BusDateRelated = False
                    If d.Text <> "" Then
                        Select Case d.Field.ValueType
                            Case DbType.Date, DbType.DateTime
                                If UCase(d.Text) = "T" _
                                    And Publics.BusinessDate <> Nothing Then
                                    d.BusDateRelated = True
                                    d.Value = Publics.BusinessDate
                                ElseIf UCase(d.Text) = "Y" _
                                    And Publics.GetVariable("Yesterday") <> "" Then
                                    d.BusDateRelated = True
                                    d.Value = Publics.GetVariable("Yesterday")
                                ElseIf UCase(d.Text) = "M" _
                                    And Publics.GetVariable("Tomorrow") <> "" Then
                                    d.BusDateRelated = True
                                    d.Value = Publics.GetVariable("Tomorrow")
                                End If

                            Case DbType.Currency
                                Dim dc As Decimal
                                Try
                                    dc = CType(Value, Decimal)
                                    d.Text = GetString(dc)
                                Catch
                                End Try

                            Case DbType.Double
                                Dim db As Double
                                Try
                                    db = CType(Value, Double)
                                    If d.Field.Decimals > -1 Then
                                        d.Text = _
                                            Trim(Str(Publics.Round(db, d.Field.Decimals)))
                                    Else
                                        d.Text = GetString(db)
                                    End If
                                Catch
                                End Try

                            Case DbType.Int32
                                Dim i2 As Int32
                                Try
                                    i2 = CType(Value, Integer)
                                    d.Text = GetString(i2)
                                Catch
                                End Try

                            Case DbType.Int64
                                Dim i4 As Int64
                                Try
                                    i4 = CType(Value, Int64)
                                    d.Text = GetString(i4)
                                Catch
                                End Try
                        End Select
                    End If
                    cc.Text = d.Text

                    If d.Field.DisplayHeight > 1 Then
                        Dim txt As TextBox = CType(cc, TextBox)
                        If txt.Lines.Length > d.Field.DisplayHeight Then
                            txt.ScrollBars = ScrollBars.Vertical
                        Else
                            txt.ScrollBars = ScrollBars.None
                        End If
                    End If

            End Select
            CheckField(Field)
        Catch ex As Exception
            Dim i As Integer = 9
        End Try
    End Sub

    Friend Sub DoMenu(ByVal index As Integer)
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim mi As MenuItem
        Dim bDef As Boolean = True

        If index = 1 Then
            fForm.ContextMenu1.MenuItems.Clear()

            Select Case fForm.ContextMenu1.SourceControl.GetType.Name
                Case "TextBox"
                    Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)

                    ''Handle Undo text.
                    mi = fForm.ContextMenu1.MenuItems.Add("Undo", _
                                        New EventHandler(AddressOf mnuUndo_Click))
                    If Not t.CanUndo Then
                        mi.Enabled = False
                    End If

                    fForm.ContextMenu1.MenuItems.Add("-")

                    Dim blnEnable As Boolean = (t.SelectedText.Length > 0)
                    mi = fForm.ContextMenu1.MenuItems.Add("Cut", _
                                        New EventHandler(AddressOf mnuCut_Click))
                    mi.Enabled = blnEnable
                    mi = fForm.ContextMenu1.MenuItems.Add("Copy", _
                                        New EventHandler(AddressOf mnuCopy_Click))
                    mi.Enabled = blnEnable

                    mi = fForm.ContextMenu1.MenuItems.Add("Paste", _
                                        New EventHandler(AddressOf mnuPaste_Click))
                    mi.Enabled = iData.GetDataPresent(GetType(String))

                    mi = fForm.ContextMenu1.MenuItems.Add("Delete", _
                                        New EventHandler(AddressOf mnuDelete_Click))
                    mi.Enabled = blnEnable

                    fForm.ContextMenu1.MenuItems.Add("-")

                    mi = fForm.ContextMenu1.MenuItems.Add("Select All", _
                                        New EventHandler(AddressOf mnuSelect_Click))
                    If t.SelectedText = t.Text Then
                        mi.Enabled = False
                    End If

                    fForm.ContextMenu1.MenuItems.Add("-")
                    'If t.Multiline Then
                    FieldText = t.Text
                    mi = fForm.ContextMenu1.MenuItems.Add("View text...", _
                                        New EventHandler(AddressOf DisplayText))
                    'End If
                Case "ComboBox"
                    Dim t As ComboBox = _
                                CType(fForm.ContextMenu1.SourceControl, ComboBox)

                    Dim blnEnable As Boolean = (t.SelectedText.Length > 0)
                    mi = fForm.ContextMenu1.MenuItems.Add("Cut", _
                                        New EventHandler(AddressOf mnuCut_Click))
                    mi.Enabled = blnEnable
                    mi = fForm.ContextMenu1.MenuItems.Add("Copy", _
                                        New EventHandler(AddressOf mnuCopy_Click))
                    mi.Enabled = blnEnable

                    mi = fForm.ContextMenu1.MenuItems.Add("Paste", _
                                        New EventHandler(AddressOf mnuPaste_Click))
                    mi.Enabled = iData.GetDataPresent(GetType(String))

                    mi = fForm.ContextMenu1.MenuItems.Add("Delete", _
                                        New EventHandler(AddressOf mnuDelete_Click))
                    mi.Enabled = blnEnable

                    fForm.ContextMenu1.MenuItems.Add("-")

                    mi = fForm.ContextMenu1.MenuItems.Add("Select All", _
                                        New EventHandler(AddressOf mnuSelect_Click))
                    If t.SelectedText = t.Text Then
                        mi.Enabled = False
                    End If

                    fForm.ContextMenu1.MenuItems.Add("-")

                    Dim d As DialogField = dlogf.Item(CType(t.Tag, String))
                    If d.Field.FillProcess <> "" Then
                        mi = fForm.ContextMenu1.MenuItems.Add("Refresh data", _
                                     New EventHandler(AddressOf mnuComboRefresh_Click))
                    End If

                Case "Label"
                    Dim l As Label = CType(fForm.ContextMenu1.SourceControl, Label)
                    Dim d As DialogField = dlogf.Item(CType(l.Tag, String))
                    If d.Field.DisplayHeight > 1 Then
                        FieldText = l.Text
                        mi = fForm.ContextMenu1.MenuItems.Add("View text...", _
                                                New EventHandler(AddressOf DisplayText))
                    End If
                    bDef = False

                Case "CheckBox"
            End Select

            If bDef Then
                mi = fForm.ContextMenu1.MenuItems.Add("Set Default", _
                                    New EventHandler(AddressOf mnuDefault_Click))

            End If
        Else
            fForm.ContextMenu2.MenuItems.Clear()
            mi = fForm.ContextMenu2.MenuItems.Add("Copy to Clipboard", _
                                        New EventHandler(AddressOf mnuClip_Click))
        End If
    End Sub

    Private Sub mnuClip_Click(ByVal sender As System.Object, _
                                            ByVal e As System.EventArgs)
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim Line As String = ""
        Dim Obj As Object
        Dim s As String
        Dim d As DialogField

        For Each d In dlogf
            If d.Field.DisplayType <> "H" Then
                Obj = GetFieldValue(d.Name)
                If d.Field.Label = "" Then
                    s = d.Name
                Else
                    s = d.Field.Label
                End If
                If Obj Is Nothing Then
                    Line &= s & vbTab & "" & vbLf
                Else
                    Line &= s & vbTab & Obj.ToString & vbLf
                End If
            End If
        Next
        Clipboard.SetDataObject(New DataObject(DataFormats.Text, Line))
    End Sub

    Private Sub mnuComboRefresh_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        If ActiveField <> "" Then
            Dim d As DialogField = dlogf.Item(ActiveField)
            Dim cb As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
            Dim s As String

            If d.Field.FillProcess <> "" Then
                Dim p As New ShellProcess(d.Field.FillProcess, Me, Me.parms)
                s = cb.Text
                cb.DisplayMember = d.Field.TextField
                cb.ValueMember = d.Field.ValueField
                cb.DataSource = CType( _
                    Me.parms.Item(d.Field.FillProcess).Value, DataTable)
                cb.Text = s
            End If
        End If
    End Sub

    Private Sub mnuDefault_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        If ActiveField <> "" Then
            SetFieldText(ActiveField)
            SaveProperty(sDefn.Name, "_" & ActiveField, "u", _
                                    GetString(GetFieldValue(ActiveField)))
        End If
    End Sub

    Private Sub mnuUndo_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
        t.Undo()
    End Sub

    Private Sub mnuCut_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        Select Case fForm.ContextMenu1.SourceControl.GetType.Name
            Case "TextBox"
                Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
                t.Cut()
            Case "ComboBox"
                Dim c As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
                Clipboard.SetDataObject(New DataObject(DataFormats.Text, c.SelectedText))
                c.SelectedText = ""
        End Select
    End Sub

    Private Sub mnuCopy_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        Select Case fForm.ContextMenu1.SourceControl.GetType.Name
            Case "TextBox"
                Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
                t.Copy()
            Case "ComboBox"
                Dim c As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
                Dim iData As IDataObject = Clipboard.GetDataObject()

                Clipboard.SetDataObject(New DataObject(DataFormats.Text, c.SelectedText))
        End Select
    End Sub

    Private Sub mnuPaste_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        Select Case fForm.ContextMenu1.SourceControl.GetType.Name
            Case "TextBox"
                Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
                t.Paste()
            Case "ComboBox"
                Dim c As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
                c.SelectedText = CStr(Clipboard.GetDataObject().GetData(DataFormats.Text))
        End Select
    End Sub

    Private Sub mnuDelete_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)

        Select Case fForm.ContextMenu1.SourceControl.GetType.Name
            Case "TextBox"
                Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
                t.SelectedText = ""
            Case "ComboBox"
                Dim c As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
                c.SelectedText = ""
        End Select
    End Sub

    Private Sub mnuSelect_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs)
        Select Case fForm.ContextMenu1.SourceControl.GetType.Name
            Case "TextBox"
                Dim t As TextBox = CType(fForm.ContextMenu1.SourceControl, TextBox)
                t.SelectAll()
            Case "ComboBox"
                Dim c As ComboBox = CType(fForm.ContextMenu1.SourceControl, ComboBox)
                c.SelectAll()
        End Select
    End Sub

    Private Sub DisplayText(ByVal sender As Object, ByVal e As EventArgs)
        Dim ff As New Form
        ff.MaximizeBox = False
        ff.MinimizeBox = False
        ff.Text = fForm.Text  '' & " - " & d.Name
        ff.Icon = fForm.Icon
        ff.FormBorderStyle = FormBorderStyle.Sizable
        Dim l As New Label
        l.Text = FieldText
        l.Top = 0
        l.Left = 0
        l.Width = ff.Width
        l.Height = ff.Height
        l.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom Or AnchorStyles.Right
        ff.Controls.Add(l)
        ff.Show()
    End Sub

    Private Class ComboSource
        Private tValue As String
        Private tText As String

        Public ReadOnly Property Value() As String
            Get
                Return tValue
            End Get
        End Property

        Public ReadOnly Property Text() As String
            Get
                Return tText
            End Get
        End Property

        Public Sub New(ByVal sValue As String, ByVal sText As String)
            tValue = sValue
            tText = sText
        End Sub

        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class
End Class
