Option Explicit On 
Option Strict On

Imports System.Drawing.Drawing2D

Public Class TreeForm
    Inherits System.Windows.Forms.Form
    Friend oOwner As Tree
    Friend WithEvents statusBar As System.Windows.Forms.StatusStrip
    Friend WithEvents panel0 As System.Windows.Forms.ToolStripStatusLabel
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
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.statusBar = New System.Windows.Forms.StatusStrip
        Me.panel0 = New System.Windows.Forms.ToolStripStatusLabel
        Me.statusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'TreeView1
        '
        Me.TreeView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TreeView1.ContextMenu = Me.ContextMenu2
        Me.TreeView1.Location = New System.Drawing.Point(0, 26)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(296, 246)
        Me.TreeView1.TabIndex = 0
        '
        'ContextMenu2
        '
        '
        'ImageList
        '
        Me.ImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'statusBar
        '
        Me.statusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.panel0})
        Me.statusBar.Location = New System.Drawing.Point(0, 271)
        Me.statusBar.Name = "statusBar"
        Me.statusBar.Size = New System.Drawing.Size(296, 22)
        Me.statusBar.TabIndex = 1
        Me.statusBar.Text = "StatusStrip1"
        '
        'panel0
        '
        Me.panel0.BackColor = System.Drawing.Color.Transparent
        Me.panel0.Name = "panel0"
        Me.panel0.Size = New System.Drawing.Size(11, 17)
        Me.panel0.Text = " "
        '
        'TreeForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(296, 293)
        Me.Controls.Add(Me.statusBar)
        Me.Controls.Add(Me.TreeView1)
        Me.KeyPreview = True
        Me.Name = "TreeForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Tree"
        Me.statusBar.ResumeLayout(False)
        Me.statusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Tree_DoubleClick(ByVal sender As Object, _
                            ByVal e As System.EventArgs) Handles MyBase.DoubleClick
        If DblClkKey <> "" Then
            oOwner.ProcessAction(DblClkKey)
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, _
                        ByVal e As System.Windows.Forms.TreeViewEventArgs) _
                                                   Handles TreeView1.AfterSelect
        oOwner.SetActions()
    End Sub

    Private Sub Tree_Load(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = Publics.ShellIcon
    End Sub

    Private Sub Tree_Closing(ByVal sender As Object, _
            ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.MdiFormClosing Then
            e.Cancel = True
        Else
            oOwner.ProcessClose()
        End If
    End Sub

    Private Sub Tree_KeyDown(ByVal sender As Object, _
                    ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        oOwner.ProcessKey(e.KeyCode, e.Modifiers.ToString)
    End Sub

    Private Sub ContextMenu2_Popup(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles ContextMenu2.Popup
        oOwner.DoMenu()
    End Sub

    Private Sub statusBar_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles statusBar.Paint
        Dim tb As StatusStrip = DirectCast(sender, StatusStrip)
        Dim r As New Rectangle(0, 0, tb.Width, tb.Height)
        Dim Br As New LinearGradientBrush(r, DialogStyle.NameToColour(DialogStyle.ToolEnd), DialogStyle.NameToColour(DialogStyle.ToolStart), LinearGradientMode.Vertical)
        e.Graphics.FillRectangle(Br, e.ClipRectangle)
        'tb.Items(0).Width = tb.Width - 20
    End Sub
End Class

Public Class TreeDefn
    Inherits ObjectDefn

    Private sTitle As String
    Private sDataParameter As String
    Private sKeyColumn As String
    Private sDescriptionColumn As String
    Private sParentColumn As String
    Private sTypeColumn As String
    Private sColourColumn As String
    Private sDefaultImage As String
    Private sTitleParameter() As String
    Private bRefreshTree As Boolean = False
    Private sHelpPage As String

    Public ReadOnly Property Title() As String
        Get
            Title = sTitle
        End Get
    End Property

    Public ReadOnly Property DataParameter() As String
        Get
            DataParameter = sDataParameter
        End Get
    End Property

    Public ReadOnly Property KeyColumn() As String
        Get
            KeyColumn = sKeyColumn
        End Get
    End Property

    Public ReadOnly Property DescriptionColumn() As String
        Get
            DescriptionColumn = sDescriptionColumn
        End Get
    End Property

    Public ReadOnly Property ParentColumn() As String
        Get
            ParentColumn = sParentColumn
        End Get
    End Property

    Public ReadOnly Property TypeColumn() As String
        Get
            TypeColumn = sTypeColumn
        End Get
    End Property

    Public ReadOnly Property ColourColumn() As String
        Get
            ColourColumn = sColourColumn
        End Get
    End Property

    Public ReadOnly Property DefaultImage() As String
        Get
            DefaultImage = sDefaultImage
        End Get
    End Property

    Public ReadOnly Property TitleParameter() As String()
        Get
            TitleParameter = sTitleParameter
        End Get
    End Property

    Public ReadOnly Property RefreshTree() As Boolean
        Get
            RefreshTree = bRefreshTree
        End Get
    End Property

    Public ReadOnly Property HelpPage() As String
        Get
            HelpPage = sHelpPage
        End Get
    End Property

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New Tree(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case LCase(Name)
            Case "title"
                sTitle = GetString(Value)
            Case "dataparameter"
                sDataParameter = GetString(Value)
            Case "keycolumn"
                sKeyColumn = GetString(Value)
            Case "descriptioncolumn"
                sDescriptionColumn = GetString(Value)
            Case "parentcolumn"
                sParentColumn = GetString(Value)
            Case "typecolumn"
                sTypeColumn = GetString(Value)
            Case "colourcolumn"
                sColourColumn = GetString(Value)
            Case "defaultimage"
                sDefaultImage = GetString(Value)
            Case "titleparameters"
                sTitleParameter = Split(GetString(Value), "||")
            Case "helppage"
                sHelpPage = GetString(Value)
            Case "refreshtree"
                If GetString(Value) = "Y" Then
                    bRefreshTree = True
                End If
            Case Else
                Publics.MessageOut(Name & " property is not supported by Tree object")
        End Select
    End Sub
End Class

Public Class Tree
    Inherits ShellObject

    Private sDefn As TreeDefn
    Private fForm As TreeForm
    Private WithEvents mAction As ShellMenu
    Private Images() As String
    Private bFormOff As Boolean = False
    Private bCloseState As Boolean = False
    Private sTitle As String
    Private bListen As Boolean = False

    Public Sub New(ByVal Defn As TreeDefn)
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
            Dim p As shellParameter
            Try
                For Each p In MyBase.parms
                    If p.Output Then
                        If p.Name = sDefn.KeyColumn Then
                            p.Value = GetTreeValue("Key")
                        ElseIf p.Name = sDefn.ParentColumn Then
                            p.Value = GetTreeValue("Parent")
                        ElseIf p.Name = sDefn.DescriptionColumn Then
                            p.Value = GetTreeValue("Description")
                        ElseIf p.Name = sDefn.TypeColumn Then
                            p.Value = GetTreeValue("Type")
                        ElseIf p.Name = sDefn.ColourColumn Then
                            p.Value = GetTreeValue("Colour")
                        ElseIf Not p.Input Then
                            p.Value = GetTreeValue(p.Name)
                        End If
                    End If
                Next
            Catch ex As Exception
                Dim i As Integer = 9
            End Try
            Return MyBase.parms
        End Get
    End Property

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim s As String
        Try
            Me.parms.MergeValues(Parms)
            If fForm Is Nothing Then
                bListen = False
                fForm = New TreeForm
                fForm.oOwner = Me
                fForm.BackColor = Publics.GetBackColour()
                fForm.TreeView1.BackColor = fForm.BackColor
                fForm.Name = sDefn.Title
                SetTitle()
                If Not Publics.MDIParent Is Nothing Then
                    fForm.MdiParent = Publics.MDIParent
                End If
                InitialiseAction()
                Publics.SetFormPosition(CType(fForm, Form), sDefn.Name)
                fForm.Show()
            Else
                SetTitle()
            End If

            s = sDefn.DataParameter
            If Me.parms.Item(s).Value Is Nothing Then
                ProcessAction("Refresh")
            Else
                If bListen Then
                    ReInitTree(CType(Parms.Item(s).Value, DataTable))
                Else
                    InitImages()
                    InitTree(CType(Parms.Item(s).Value, DataTable))
                End If
                Parms.Item(s).Value = Nothing
                SetActions()
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

    Public Overrides Sub Listener(ByVal Parms As ShellParameters)
        Dim n As TreeNode
        Dim s As String

        If sDefn.RefreshTree Then   ' update all rows
            s = sDefn.DataParameter
            If s <> "" Then
                If Not Parms.Item(s) Is Nothing Then
                    If Not Parms.Item(s).Value Is Nothing Then
                        ReInitTree(CType(Parms.Item(s).Value, DataTable))
                        SetActions()
                    Else
                        s = ""
                    End If
                Else
                    s = ""
                End If
            End If
            If s = "" Then
                bListen = True
                ProcessAction("Refresh")
                bListen = False
            End If
        Else                     ' update single row...
            n = FindNode(fForm.TreeView1.Nodes, GetString(Parms.Item(sDefn.KeyColumn).Value))
            If Not n Is Nothing Then
                Dim dr As DataRow = CType(n.Tag, DataRow)
                For Each pm As shellParameter In Parms
                    Try
                        dr.Item(pm.Name) = pm.Value
                    Catch ex As Exception
                    End Try
                Next
                SetNode(n, dr)
            End If
        End If
    End Sub

    Public Overrides Sub Suspend(ByVal Mode As Boolean)
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

    Public Overrides Sub MsgOut(ByVal msgs As ShellMessages)
        Dim s As String = ""
        Dim b As Boolean = False

        For Each ms As ShellMessage In msgs
            s &= ms.Message & vbCrLf
            If ms.Type = "U" And Not b Then
                ProcessAction("Refresh")
                b = True
            End If
        Next
        Publics.MessageOut(s)
    End Sub

    Private Sub InitImages()
        Dim i As Integer = 0
        Dim s As String

        Try
            If sDefn.DefaultImage <> "" Then
                fForm.TreeView1.ImageList = fForm.ImageList

                ReDim Images(0)
                fForm.ImageList.Images.Add(Publics.GetImage(sDefn.DefaultImage))
                Images(0) = sDefn.DefaultImage

                For Each img As ShellProperty In sDefn.Properties
                    If img.Type = "im" Then
                        s = CType(img.Value, String)
                        If GetImageIndex(s) = -1 Then
                            i += 1
                            ReDim Preserve Images(i)
                            Images(i) = s
                            fForm.ImageList.Images.Add(Publics.GetImage(s))
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        End Try
    End Sub

    Private Function GetImageIndex(ByVal sImage As String) As Integer
        For i As Integer = 0 To Images.GetUpperBound(0)
            If Images(i) = sImage Then
                Return i
                Exit For
            End If
        Next
        Return -1
    End Function

    Private Sub InitTree(ByVal dt As DataTable)
        Dim dr As DataRow
        Dim tn As TreeNode
        Dim b As Boolean = True

        Try
            fForm.TreeView1.Nodes.Clear()
            fForm.TreeView1.BeginUpdate()
            fForm.TreeView1.ShowLines = True
            fForm.TreeView1.ShowPlusMinus = True
            fForm.TreeView1.ShowRootLines = True
            fForm.TreeView1.HideSelection = False

            Do While b
                b = False
                For Each dr In dt.Rows
                    If FindNode(fForm.TreeView1.Nodes, _
                        GetString(dr.Item(sDefn.KeyColumn))) Is Nothing Then
                        If GetString(dr.Item(sDefn.ParentColumn)) = _
                                    GetString(dr.Item(sDefn.KeyColumn)) Then
                            AddNode(fForm.TreeView1, Nothing, dr)
                        Else
                            tn = FindNode(fForm.TreeView1.Nodes, _
                                    GetString(dr.Item(sDefn.ParentColumn)))
                            If tn Is Nothing Then
                                b = True
                            Else
                                AddNode(fForm.TreeView1, tn, dr)
                            End If
                        End If
                    End If
                Next
            Loop

        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        Finally
            fForm.TreeView1.ExpandAll()
            fForm.TreeView1.EndUpdate()
        End Try
        Try
            fForm.TreeView1.SelectedImageIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ResetNodes(ByVal tn As TreeNode, ByVal del As Boolean)
        If del And tn.Checked Then
            tn.Remove()
        Else
            tn.Checked = True
            For Each n As TreeNode In tn.Nodes
                ResetNodes(n, del)
            Next
        End If
    End Sub

    Private Sub ReInitTree(ByVal dt As DataTable)
        Dim dr As DataRow
        Dim tn As TreeNode
        Dim b As Boolean = True

        Try
            fForm.TreeView1.BeginUpdate()
            ' initialise nodes so missing items can be identified after updates
            For Each tn In fForm.TreeView1.Nodes
                ResetNodes(tn, False)
            Next

            ' update changed rows
            Do While b
                b = False
                For Each dr In dt.Rows
                    tn = FindNode(fForm.TreeView1.Nodes, GetString(dr.Item(sDefn.KeyColumn)))
                    If tn Is Nothing Then
                        If GetString(dr.Item(sDefn.ParentColumn)) = _
                                    GetString(dr.Item(sDefn.KeyColumn)) Then
                            AddNode(fForm.TreeView1, Nothing, dr)
                        Else
                            tn = FindNode(fForm.TreeView1.Nodes, _
                                    GetString(dr.Item(sDefn.ParentColumn)))
                            If tn Is Nothing Then
                                b = True
                            Else
                                AddNode(fForm.TreeView1, tn, dr)
                            End If
                        End If
                    Else
                        SetNode(tn, dr)
                    End If
                Next
            Loop

            '' remove missing rows
            For Each tn In fForm.TreeView1.Nodes
                ResetNodes(tn, True)
            Next

        Catch ex As Exception
            Publics.MessageOut(ex.Message)
        Finally
            fForm.TreeView1.EndUpdate()
        End Try
    End Sub

    Private Sub AddNode(ByVal T As TreeView, ByVal tn As TreeNode, ByVal dr As DataRow)
        Dim n As TreeNode

        n = New TreeNode(GetString(dr.Item(sDefn.DescriptionColumn)))
        SetNode(n, dr)
        If tn Is Nothing Then
            T.Nodes.Add(n)
        Else
            tn.Nodes.Add(n)
        End If
    End Sub

    Private Sub SetNode(ByVal n As TreeNode, ByVal dr As DataRow)
        Dim s As String
        Dim i As Integer
        Dim p As ShellProperty
        Dim sColour As String = ""

        n.Tag = dr
        n.Checked = False
        n.Text = GetString(dr.Item(sDefn.DescriptionColumn))
        If Not Images Is Nothing Then
            If Images.GetUpperBound(0) > 0 Then
                p = sDefn.Properties.Item(GetString(dr.Item(sDefn.TypeColumn)), "im")
                If Not p Is Nothing Then
                    s = GetString(p.Value)
                    If s <> "" Then
                        i = GetImageIndex(s)
                    End If
                End If
                If i > 0 Then
                    n.ImageIndex = i
                    n.SelectedImageIndex = i
                Else
                    n.ImageIndex = 0
                    n.SelectedImageIndex = 0
                End If
            End If
        End If

        If Not sDefn.ColourColumn Is Nothing Then
            sColour = GetString(dr.Item(sDefn.ColourColumn))
        End If
        If sColour <> "" Then
            p = sDefn.Properties.Item(sColour, "cl")
            If Not p Is Nothing Then
                s = GetString(p.Value)
                If s <> "" Then
                    n.ForeColor = Color.FromName(s)
                End If
            Else
                n.ForeColor = Color.Blue
            End If
        Else
            n.ForeColor = Color.Blue
        End If
    End Sub

    Private Function FindNode(ByRef T As TreeNodeCollection, _
                              ByVal sID As String) As TreeNode
        Dim retNode As TreeNode = Nothing
        Dim dr As DataRow

        For Each oNode As TreeNode In T
            dr = CType(oNode.Tag, DataRow)
            If GetString(dr(sDefn.KeyColumn)) = sID Then
                retNode = oNode
                Exit For
            End If

            If oNode.Nodes.Count > 0 Then
                retNode = FindNode(oNode.Nodes, sID)
                If Not retNode Is Nothing Then
                    Exit For
                End If
            End If
        Next
        Return retNode
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
            Me.ProcessAction(sAction)
        End If
    End Sub

    Friend Sub ProcessClose()
        SaveFormPosition(CType(fForm, Form), sDefn.Name)
        Register.Remove(Me.RegKey)

        If bCloseState Then
            Me.OnExitOkay()
        Else
            Me.OnExitFail()
        End If
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

    Private Function DoProcess(ByRef a As ActionDefn) As Boolean
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
            s = GetString(GetTreeValue(Action.ProcessField))
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

    Friend Sub SetActions()
        Dim bSet As Boolean
        Dim vRule As String
        Dim bRule As Boolean

        ' Scan through the Actions and Enable / Disable them for the selected row

        For Each a As ActionDefn In sDefn.Actions

            ' Determine Enabled Properties
            bSet = True

            If bFormOff Or GetActionProcess(a) = "" Then
                bSet = False
            ElseIf a.RowBased And fForm.TreeView1.SelectedNode Is Nothing Then
                bSet = False
            ElseIf Not a.Rules Is Nothing Then      ' rules always enabled
                If Not fForm.TreeView1.SelectedNode Is Nothing Then
                    bSet = True
                    For Each cRuleD As ActionRuleDefn In a.Rules
                        bRule = False
                        For Each cRule As ActionRule In cRuleD.Rules
                            If cRule.FieldName = sDefn.KeyColumn Then
                                vRule = GetTreeValue("Key")
                            ElseIf cRule.FieldName = sDefn.ParentColumn Then
                                vRule = GetTreeValue("Parent")
                            ElseIf cRule.FieldName = sDefn.DescriptionColumn Then
                                vRule = GetTreeValue("Description")
                            ElseIf cRule.FieldName = sDefn.TypeColumn Then
                                vRule = GetTreeValue("Type")
                            ElseIf cRule.FieldName = sDefn.ColourColumn Then
                                vRule = GetTreeValue("Colour")
                            Else
                                vRule = GetTreeValue(cRule.FieldName)
                            End If

                            Select Case cRule.Type
                                Case ActionRule.ValidationType.EQ
                                    If vRule.ToString = cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.NE
                                    If vRule.ToString <> cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.NN
                                    If vRule.ToString <> "" Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.GT
                                    If vRule.ToString > cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.GE
                                    If vRule.ToString >= cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.LT
                                    If vRule.ToString < cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                                Case ActionRule.ValidationType.LE
                                    If vRule.ToString <= cRule.Value.ToString Then
                                        bRule = True
                                        Exit For
                                    End If
                            End Select
                        Next
                        If Not bRule Then
                            bSet = False
                            Exit For
                        End If
                    Next
                End If
            End If

            a.Enabled = bSet
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

    Private Function GetTreeValue(ByVal sKey As String) As String
        Dim dr As DataRow

        If fForm.TreeView1.SelectedNode Is Nothing Then
            Return Nothing
        End If
        Select Case sKey
            Case "Key"
                dr = CType(fForm.TreeView1.SelectedNode.Tag, DataRow)
                Return GetString(dr.Item(sDefn.KeyColumn))
            Case "Parent"
                If fForm.TreeView1.SelectedNode.Parent Is Nothing Then
                    dr = CType(fForm.TreeView1.SelectedNode.Tag, DataRow)
                Else
                    dr = CType(fForm.TreeView1.SelectedNode.Parent.Tag, DataRow)
                End If
                Return GetString(dr.Item(sDefn.KeyColumn))
            Case "Description"
                Return fForm.TreeView1.SelectedNode.Text
            Case "Type"
                dr = CType(fForm.TreeView1.SelectedNode.Tag, DataRow)
                Return GetString(dr.Item(sDefn.TypeColumn))
            Case "Colour"
                dr = CType(fForm.TreeView1.SelectedNode.Tag, DataRow)
                Return GetString(dr.Item(sDefn.ColourColumn))
            Case Else
                dr = CType(fForm.TreeView1.SelectedNode.Tag, DataRow)
                Try
                    Return GetString(dr.Item(sKey))
                Catch ex As Exception
                    Return Nothing
                End Try
        End Select
    End Function

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

        For Each n As TreeNode In fForm.TreeView1.Nodes
            Line &= GetNode(n, 0)
        Next
        Clipboard.SetDataObject(New DataObject(DataFormats.Text, Line))
    End Sub

    Private Function GetNode(ByVal nd As TreeNode, ByVal level As Integer) As String
        Dim i As Integer
        Dim s As String = ""
        Dim dr As DataRow

        For i = 1 To level
            s &= vbTab
        Next
        If nd.Tag Is Nothing Then
            s &= nd.Text & vbLf
        Else
            dr = CType(nd.Tag, DataRow)
            s &= nd.Text & vbTab & GetString(dr.Item(sDefn.TypeColumn)) & vbLf
        End If
        For Each n As TreeNode In nd.Nodes
            s &= GetNode(n, level + 1)
        Next
        Return s
    End Function
End Class
