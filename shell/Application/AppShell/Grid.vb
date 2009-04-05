Option Explicit On
Option Strict On
Imports System.Drawing.Drawing2D

Public Class Grid
    Inherits System.Windows.Forms.Form
    Friend oOwner As GridForm
    Friend WithEvents Context As System.Windows.Forms.ContextMenuStrip
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
    Friend WithEvents Grid1 As System.Windows.Forms.DataGridView
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents statusBar As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Grid1 = New System.Windows.Forms.DataGridView
        Me.Context = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.statusBar = New System.Windows.Forms.StatusBar
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Grid1
        '
        Me.Grid1.AllowUserToAddRows = False
        Me.Grid1.AllowUserToDeleteRows = False
        Me.Grid1.AllowUserToOrderColumns = True
        Me.Grid1.BackgroundColor = System.Drawing.Color.White
        Me.Grid1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Grid1.ContextMenuStrip = Me.Context
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Grid1.DefaultCellStyle = DataGridViewCellStyle1
        Me.Grid1.Location = New System.Drawing.Point(0, 23)
        Me.Grid1.MultiSelect = False
        Me.Grid1.Name = "Grid1"
        Me.Grid1.ReadOnly = True
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.Grid1.RowHeadersWidth = 23
        Me.Grid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.Grid1.Size = New System.Drawing.Size(312, 314)
        Me.Grid1.TabIndex = 1
        '
        'Context
        '
        Me.Context.Name = "Context"
        Me.Context.Size = New System.Drawing.Size(61, 4)
        '
        'statusBar
        '
        Me.statusBar.ContextMenuStrip = Me.Context
        Me.statusBar.Location = New System.Drawing.Point(0, 343)
        Me.statusBar.Name = "statusBar"
        Me.statusBar.Size = New System.Drawing.Size(312, 22)
        Me.statusBar.TabIndex = 3
        '
        'Grid
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(312, 365)
        Me.Controls.Add(Me.statusBar)
        Me.Controls.Add(Me.Grid1)
        Me.KeyPreview = True
        Me.Name = "Grid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Grid"
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Grid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Grid1.DoubleClick
        If DblClkKey <> "" Then
            oOwner.ProcessAction(DblClkKey)
        End If
    End Sub

    Private Sub Grid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        oOwner.ProcessKey(e.KeyCode, e.Modifiers.ToString)
    End Sub

    Private Sub Context_Popup(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Context.Opening
        oOwner.DoMenu()
    End Sub

    Private Sub Grid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Grid1.CurrentCellChanged
        oOwner.SetActions()
    End Sub

    Private Sub Grid_RowPrePaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) Handles Grid1.RowPrePaint
        oOwner.SetRowColour(sender, e)
    End Sub

    Private Sub dataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles Grid1.CellFormatting
        e.CellStyle.ForeColor = oOwner.rowforecolour
        e.CellStyle.BackColor = oOwner.rowbackcolour
    End Sub

    Private Sub Grid_Load(ByVal sender As System.Object, _
                             ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = Publics.ShellIcon

        For i As Integer = 2 To Grid1.Controls.Count - 1
            AddHandler Grid1.Controls(i).KeyDown, AddressOf Grid1_KeyUp
        Next
    End Sub

    Private Sub Grid_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Me.Grid1.Width = Me.Width - 10
        Me.Grid1.Height = Me.statusBar.Top - Me.Grid1.Top
    End Sub

    Private Sub Grid_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        oOwner.ProcessClose()
    End Sub
End Class

Public Class GridDefn
    Inherits ObjectDefn

    Public Title As String
    Public DataParameter As String
    Public ColourColumn As String
    Public TitleParameter() As String
    Public HelpPage As String
    Public StateFilter As Boolean = False

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New GridForm(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case LCase(Name)
            Case "title"
                Title = GetString(Value)
            Case "dataparameter"
                DataParameter = GetString(Value)
            Case "colourcolumn"
                ColourColumn = GetString(Value)
            Case "statefilter"
                StateFilter = True
            Case "titleparameters"
                TitleParameter = Split(GetString(Value), "||")
            Case "helppage"
                HelpPage = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Grid object")
        End Select
    End Sub
End Class

Public Class GridForm
    Inherits ShellObject

    Public rowforecolour As System.Drawing.Color = Color.Black
    Public rowbackcolour As System.Drawing.Color = Color.White

    Private sDefn As GridDefn
    Private fForm As Grid

    Private bFilter As Boolean = True
    Private WithEvents mAction As ShellMenu
    Private bFormOff As Boolean = False
    Private ActionStates As New ActionStates
    Private bCloseState As Boolean = False
    Private sTitle As String

    Public Sub New(ByVal Defn As GridDefn)
        Dim r As ObjectRegister
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)

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
            Dim i As Integer
            Dim p As shellParameter
            Try
                i = fForm.Grid1.CurrentRow.Index
                For Each c As Field In sDefn.Fields
                    p = MyBase.parms.Item(c.Name)
                    If Not p Is Nothing Then
                        If p.Output Then
                            p.Value = GetGridValue(i, c.Name)
                        End If
                    End If
                Next
            Catch
            End Try
            Return MyBase.parms
        End Get
    End Property

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim sData As String = ""
        Dim ss As String
        Dim dt As DataTable = Nothing
        Dim i As Integer
        Dim gc As DataGridViewTextBoxColumn

        Try
            Me.parms.MergeValues(Parms)
            If Not Parms Is Nothing Then
                sData = sDefn.DataParameter
                If Not Parms.Item(sData) Is Nothing Then
                    If Not Parms.Item(sData).Value Is Nothing Then
                        If Parms.Item(sData).Output Then
                            dt = CType(Parms.Item(sData).Value, DataTable)
                        End If
                    End If
                End If
            End If

            If fForm Is Nothing Then
                fForm = New Grid
                fForm.oOwner = Me
                fForm.BackColor = Publics.GetBackColour
                fForm.Grid1.BackgroundColor = fForm.BackColor

                fForm.Name = sDefn.Title
                SetTitle()

                If Not Publics.MDIParent Is Nothing Then
                    fForm.MdiParent = CType(Publics.MDIParent, Form)
                End If
                InitialiseGrid()
                InitialiseAction()
                Publics.SetFormPosition(CType(fForm, Form), sDefn.Name)
                fForm.Show()
                If dt Is Nothing Then
                    ProcessAction("Refresh")
                Else
                    i = 9
                End If
            Else
                mAction.Update(Me.parms)
                SetTitle()
            End If

            If Not dt Is Nothing Then
                Dim f As Field
                For Each t As DataGridViewColumn In fForm.Grid1.Columns
                    f = sDefn.Fields.Item(t.Name)
                    If Not f Is Nothing Then
                        If f.DisplayType = "F" Then
                            Try
                                ss = dt.Columns.Item(f.Name).Caption
                                t.HeaderText = ss
                            Catch ex As Exception
                                i = i
                            End Try
                        End If
                    End If
                Next

                For Each tr As DataColumn In dt.Columns
                    If fForm.Grid1.Columns(tr.ColumnName) Is Nothing Then
                        gc = New DataGridViewTextBoxColumn
                        gc.Name = tr.ColumnName
                        gc.DataPropertyName = tr.ColumnName
                        gc.Visible = False
                        fForm.Grid1.Columns.Add(gc)
                    End If
                Next
                fForm.Grid1.DataSource = dt
                SetFilter()

                Parms.Item(sData).Value = Nothing
                fForm.statusBar.Text = dt.Rows.Count & " rows."
            End If
            SetActions()
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
                sTemp = GetString(parms.Item(ss).Value)
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

    Private Sub SetFilter()
        Dim dt As DataTable
        If sDefn.StateFilter Then
            dt = CType(fForm.Grid1.DataSource, DataTable)
            If dt Is Nothing Then
                Return
            End If
            If bFilter Then
                dt.DefaultView.RowFilter() = "State<>'dl'"
            Else
                dt.DefaultView.RowFilter() = ""
            End If
        End If
    End Sub

    Public Overrides Sub Listener(ByVal Params As ShellParameters)
        Dim b As Boolean
        Try
            If Not Params Is Nothing Then

                'find grid dataset row with a primary key equal to input parameter data 

                For Each r As DataRow In CType(fForm.Grid1.DataSource, DataTable).Rows
                    b = True
                    For Each f As Field In sDefn.Fields
                        If f.Primary Then
                            If GetString(r.Item(f.Name)) <> _
                                            GetString(Params.Item(f.Name).Value) Then
                                b = False
                                Exit For
                            End If
                        End If
                    Next

                    ' if found update all the matching columns with data from params

                    If b Then
                        fForm.Grid1.SuspendLayout()
                        r.BeginEdit()
                        For Each p As shellParameter In Params
                            Try
                                r.Item(p.Name) = p.Value
                            Catch
                            End Try
                        Next
                        r.EndEdit()
                        fForm.Grid1.ResumeLayout()
                        SetActions()
                        Exit For    'this a primary key so there can only be 1 matching row!
                    End If
                Next
            End If
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Public Overrides Sub Suspend(ByVal Mode As Boolean)
        Application.DoEvents()
        If Mode Then
            fForm.Cursor = Cursors.WaitCursor
            bFormOff = True
            SetActions()
        Else
            bFormOff = False
            SetActions()
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

    Private Sub InitialiseGrid()
        Dim a() As String
        Dim cb() As Boolean
        Dim i As Integer = 0
        Dim b As Boolean = True
        Dim s As String
        Dim c As Field
        Dim col As DataGridViewColumn
        Dim w As Integer = 0

        Try
            ReDim a(sDefn.Fields.count - 1)
            ReDim cb(sDefn.Fields.count - 1)

            ''restore Grid column order and widths

            If Not sDefn.Properties.Item("Columns", "u") Is Nothing Then
                Dim cc() As String
                Dim ca() As String

                cc = Split(GetString(sDefn.Properties.Item("Columns", "u").Value), "| |")
                For Each s In cc
                    ca = Split(s, "||")
                    c = sDefn.Fields.Item(ca(0))
                    If Not c Is Nothing Then
                        If Not cb(c.Index) Then
                            a(i) = ca(0)
                            If c.DisplayWidth > 0 And UCase(c.DisplayType) <> "H" Then
                                c.DisplayWidth = CType(ca(1), Integer)
                            End If
                            cb(c.Index) = True
                        End If
                        i += 1
                    End If
                Next
            End If

            For Each c In sDefn.Fields
                If Not cb(c.Index) Then
                    cb(c.Index) = True
                    a(i) = c.Name
                    i += 1
                End If
            Next

            For Each s In a
                c = sDefn.Fields.Item(s)
                col = New DataGridViewTextBoxColumn

                col.Name = c.Name
                col.DataPropertyName = c.Name
                col.DefaultCellStyle.Format = c.Format
                col.DefaultCellStyle.NullValue = c.NullText
                ''col.DefaultCellStyle.Font() = new Font(
                Select Case c.Justify
                    Case "C"
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Case "R"
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Case "D"
                        If c.ValueType = DbType.Currency Or _
                           c.ValueType = DbType.Double Or _
                           c.ValueType = DbType.Int32 Or _
                           c.ValueType = DbType.Int64 Then
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        End If
                End Select
                If c.DisplayWidth < 1 Or UCase(c.DisplayType) = "H" Then
                    col.Visible = False
                Else
                    col.Width = c.DisplayWidth
                    w += c.DisplayWidth
                    col.HeaderText = c.Label
                End If
                fForm.Grid1.Columns.Add(col)
            Next
            If w > 1000 Then w = 1000
            fForm.Width = w + 50

        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Public Sub SetRowColour(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs)
        Dim sColumn As String
        Dim cs As shellProperties
        Dim c As ShellProperty

        rowforecolour = Color.Black
        rowbackcolour = Color.White
        sColumn = sDefn.ColourColumn
        If sColumn <> "" Then
            cs = sDefn.Properties
            If Not cs Is Nothing Then
                c = cs.Item(CType(GetGridValue(e.RowIndex, sColumn), String), "cl")
                If Not c Is Nothing Then
                    rowforecolour = System.Drawing.Color.FromName(CType(c.Value, String))
                End If
                c = cs.Item(CType(GetGridValue(e.RowIndex, sColumn), String), "cb")
                If Not c Is Nothing Then
                    rowbackcolour = System.Drawing.Color.FromName(CType(c.Value, String))
                End If
            End If
        End If
    End Sub

    Private Sub InitialiseAction()
        mAction = New ShellMenu(sDefn)
        mAction.fForm = fForm
        mAction.Update(Me.parms)

        For Each a As ActionDefn In sDefn.Actions
            ActionStates.Add(a.Name, a.Enabled)
            If a.IsDblClick Then
                fForm.DblClkKey = a.Name
            End If
        Next
    End Sub

    Friend Sub ProcessKey(ByVal KeyCode As Integer, ByVal Shift As String)
        Dim b As Boolean = False
        If KeyCode = Keys.F1 Then
            If Shift = "Shift" Then
                b = True
            End If
            Publics.RaiseHelp(b, sDefn.HelpPage)
        Else
            For Each a As ActionDefn In sDefn.Actions
                Dim ae As ActionState = ActionStates.Item(a.Name)
                If Not ae Is Nothing Then
                    If a.KeyCode = KeyCode And ae.Enabled And a.IsKey Then
                        If a.Shift = Shift Or a.Shift Is Nothing Then
                            Dim p As New ShellProcess(a.Process, Me, Me.parms)
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub mAction_Action(ByVal sAction As String) Handles mAction.Action
        If sAction <> "" Then
            Me.ProcessAction(sAction)
        End If
    End Sub

    Friend Sub ProcessAction(ByVal sKey As String)
        Dim a As ActionDefn
        If Not sDefn.Actions Is Nothing Then
            a = sDefn.Actions.Item(sKey)
            If Not a Is Nothing Then
                Dim ae As ActionState = ActionStates.Item(a.Name)
                If Not ae Is Nothing Then
                    If ae.Enabled Then
                        DoProcess(a)
                    End If
                End If
            End If
        End If
    End Sub

    Private Function DoProcess(ByRef a As ActionDefn) As Boolean
        If a.LinkedParam <> "" Then
            Me.parms.MergeValues(mAction.parms)
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

        If Action.ProcessField <> "" And Not Action.Processes Is Nothing And Not fForm.Grid1.CurrentRow Is Nothing Then
            s = GetString(GetGridValue(fForm.Grid1.CurrentRow.Index, Action.ProcessField))
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

    Friend Sub ProcessClose()
        Dim s As String = ""
        Dim Columns() As String

        Try
            ReDim Columns(fForm.Grid1.ColumnCount)
            Register.Remove(Me.RegKey)
            SaveFormPosition(CType(fForm, Form), sDefn.Name)
            For Each c As DataGridViewColumn In fForm.Grid1.Columns
                Columns(c.DisplayIndex) = c.Name & "||" & c.Width
            Next
            For Each t As String In Columns
                If s <> "" Then s &= "| |"
                s &= t
            Next
            Publics.SaveProperty(sDefn.Name, "Columns", "u", s)

            If bCloseState Then
                Me.OnExitOkay()
            Else
                Me.OnExitFail()
            End If
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Friend Sub SetActions()
        Dim bSet As Boolean
        Dim bRule As Boolean
        Dim vRule As Object

        ' Scan through the Actions and Enable / Disable them for the current grid data row

        For Each a As ActionDefn In sDefn.Actions

            ' Determine Enabled Properties

            If bFormOff Or a.RowBased And fForm.Grid1.CurrentRow Is Nothing Then
                bSet = False
            ElseIf GetActionProcess(a) = "" Then
                bSet = False
            Else
                bSet = True
                If Not a.Rules Is Nothing And Not fForm.Grid1.DataSource Is Nothing Then      ' no rules always enabled
                    If CType(fForm.Grid1.DataSource, DataTable).Rows.Count > 0 And Not fForm.Grid1.CurrentRow Is Nothing Then
                        bSet = False
                        For Each cRuleD As ActionRuleDefn In a.Rules
                            bRule = True
                            For Each cRule As ActionRule In cRuleD.Rules
                                vRule = (GetGridValue(fForm.Grid1.CurrentRow.Index, _
                                                                    cRule.FieldName))
                                Select Case cRule.Type
                                    Case ActionRule.ValidationType.EQ
                                        If vRule.ToString <> cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.NE
                                        If vRule.ToString = cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.NN
                                        If vRule.ToString = "" Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.GT
                                        If vRule.ToString <= cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.GE
                                        If vRule.ToString < cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.LT
                                        If vRule.ToString >= cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                    Case ActionRule.ValidationType.LE
                                        If vRule.ToString > cRule.Value.ToString Then
                                            bRule = False
                                            Exit For
                                        End If
                                End Select
                            Next
                            If bRule Then
                                bSet = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            Dim ae As ActionState = ActionStates.Item(a.Name)
            If Not ae Is Nothing Then
                ae.Enabled = bSet
            End If

            a.Enabled = bSet
            mAction.Enable(a)
        Next
    End Sub

    Private Function GetGridValue(ByVal Row As Integer, ByVal Column As String) As Object
        Try
            Return fForm.Grid1.Rows(Row).Cells(Column).Value
        Catch
            Return Nothing
        End Try
    End Function

    Friend Sub DoMenu()
        fForm.Context.Items.Clear()
        fForm.Context.Items.Add("Copy to Clipboard", Nothing, _
                                    New EventHandler(AddressOf mnuCopy_Click))
        fForm.Context.Items.Add("Export to file", Nothing, _
                                    New EventHandler(AddressOf mnuExport_Click))
        If sDefn.StateFilter Then
            If bFilter Then
                fForm.Context.Items.Add("Display all records", Nothing, _
                                            New EventHandler(AddressOf mnuSetFilter))
            Else
                fForm.Context.Items.Add("Hide disabled rows", Nothing, _
                                            New EventHandler(AddressOf mnuSetFilter))
            End If
        End If
    End Sub

    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim iData As IDataObject = Clipboard.GetDataObject()
        Dim names() As String = Nothing
        Dim i As Integer = -1
        Dim sName As String
        Dim Line As String = ""
        Dim s As String

        For Each t As DataGridViewColumn In fForm.Grid1.Columns
            If t.Visible Then
                If Not CType(fForm.Grid1.DataSource, DataTable).Columns.Item(t.Name) Is Nothing Then
                    s = t.HeaderText
                    Line &= vbTab & s
                    i += 1
                    ReDim Preserve names(i)
                    names(i) = t.Name
                End If
            End If
        Next

        s = Mid(Line, 2) & vbLf

        Dim dv As DataView = CType(fForm.Grid1.DataSource, DataTable).DefaultView
        For Each drv As DataRowView In dv
            Line = ""
            For Each sName In names
                Line &= vbTab & Publics.GetString(drv(sName))
            Next
            s &= Mid(Line, 2) & vbLf
            Application.DoEvents()
        Next

        Clipboard.SetDataObject(New DataObject(DataFormats.Text, s))
    End Sub

    Private Sub mnuExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim names() As String = Nothing
        Dim i As Integer = -1
        Dim sName As String
        Dim Line As String = ""
        Dim s As String

        Dim sFileName As String = System.IO.Path.GetTempFileName()
        Kill(sFileName)
        sFileName = sFileName.Replace(".tmp", ".csv")
        Dim fs As New System.IO.FileStream(sFileName, System.IO.FileMode.Create, _
                        System.IO.FileAccess.Write, System.IO.FileShare.None)
        Dim w As New System.IO.StreamWriter(fs)

        For Each t As DataGridViewColumn In fForm.Grid1.Columns
            If t.Visible Then
                If Not CType(fForm.Grid1.DataSource, DataTable).Columns.Item(t.Name) Is Nothing Then
                    s = t.HeaderText
                    s = s.Replace("""", "'")
                    Line &= ",""" & s & """"
                    i += 1
                    ReDim Preserve names(i)
                    names(i) = t.Name
                End If
            End If
        Next

        w.WriteLine(Mid(Line, 2))

        Dim dv As DataView = CType(fForm.Grid1.DataSource, DataTable).DefaultView
        For Each drv As DataRowView In dv
            Line = ""
            For Each sName In names
                s = Publics.GetString(drv(sName))
                Line &= ",""" & s.Replace("""", "'") & """"
            Next
            w.WriteLine(Mid(Line, 2))
            Application.DoEvents()
        Next

        w.Close()
        fs.Close()

        Dim myProcess As New Process
        myProcess.StartInfo.FileName = sFileName
        Try
            myProcess.Start()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuSetFilter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bFilter = Not bFilter
        SetFilter()
    End Sub
End Class
