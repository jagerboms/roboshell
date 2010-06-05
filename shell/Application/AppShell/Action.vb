Option Explicit On 
Option Strict On

Public Class ActionDefn
    Private sName As String
    Private sProcess As String
    Private bEnabled As Boolean = True
    Private bChecked As Boolean
    Private bRowBased As Boolean
    Private bValidate As Boolean
    Private sCloseObject As String
    Private oRules As ActionRuleDefns
    Private oProcesses As ActionProcessRuleDefns
    Private bIsDblClick As Boolean
    Private bIsButton As Boolean
    Private sImageFile As String
    Private sToolTip As String
    Private sMenuType As String    ' (N)one, (I)tem, (S)ub menu
    Private sMenuText As String
    Private sParent As String  ' identifies the menu object for sub menus
    Private bIsKey As Boolean
    Private iKeyCode As Integer
    Private sShift As String
    Private sFieldName As String      ' Action is fired when field data is changed by user
    Private sProcessField As String   ' Field used to select process to call
    Private sLinkedParam As String    ' Parameter linked to button state
    Private sParamValue As String

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Process() As String
        Get
            Process = sProcess
        End Get
        Set(ByVal v As String)
            sProcess = v
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Enabled = bEnabled
        End Get
        Set(ByVal v As Boolean)
            bEnabled = v
        End Set
    End Property

    Public Property Checked() As Boolean
        Get
            Checked = bChecked
        End Get
        Set(ByVal v As Boolean)
            bChecked = v
        End Set
    End Property

    Public Property RowBased() As Boolean
        Get
            RowBased = bRowBased
        End Get
        Set(ByVal v As Boolean)
            bRowBased = v
        End Set
    End Property

    Public Property Validate() As Boolean
        Get
            Validate = bValidate
        End Get
        Set(ByVal v As Boolean)
            bValidate = v
        End Set
    End Property

    Public Property CloseObject() As String
        Get
            CloseObject = sCloseObject
        End Get
        Set(ByVal v As String)
            sCloseObject = v
        End Set
    End Property

    Public Property Rules() As ActionRuleDefns
        Get
            Rules = oRules
        End Get
        Set(ByVal v As ActionRuleDefns)
            oRules = v
        End Set
    End Property

    Public Property Processes() As ActionProcessRuleDefns
        Get
            Processes = oProcesses
        End Get
        Set(ByVal v As ActionProcessRuleDefns)
            oProcesses = v
        End Set
    End Property

    Public Property IsDblClick() As Boolean
        Get
            IsDblClick = bIsDblClick
        End Get
        Set(ByVal v As Boolean)
            bIsDblClick = v
        End Set
    End Property

    Public Property IsButton() As Boolean
        Get
            IsButton = bIsButton
        End Get
        Set(ByVal v As Boolean)
            bIsButton = v
        End Set
    End Property

    Public Property ImageFile() As String
        Get
            ImageFile = sImageFile
        End Get
        Set(ByVal v As String)
            sImageFile = v
        End Set
    End Property

    Public Property ToolTip() As String
        Get
            ToolTip = sToolTip
        End Get
        Set(ByVal v As String)
            sToolTip = v
        End Set
    End Property

    Public Property MenuType() As String
        Get
            MenuType = sMenuType
        End Get
        Set(ByVal v As String)
            sMenuType = v
        End Set
    End Property

    Public Property MenuText() As String
        Get
            MenuText = sMenuText
        End Get
        Set(ByVal v As String)
            sMenuText = v
        End Set
    End Property

    Public Property Parent() As String
        Get
            Parent = sParent
        End Get
        Set(ByVal v As String)
            sParent = v
        End Set
    End Property

    Public Property IsKey() As Boolean
        Get
            IsKey = bIsKey
        End Get
        Set(ByVal v As Boolean)
            bIsKey = v
        End Set
    End Property

    Public Property KeyCode() As Integer
        Get
            KeyCode = iKeyCode
        End Get
        Set(ByVal v As Integer)
            iKeyCode = v
        End Set
    End Property

    Public Property Shift() As String
        Get
            Shift = sShift
        End Get
        Set(ByVal v As String)
            sShift = v
        End Set
    End Property

    Public Property FieldName() As String
        Get
            FieldName = sFieldName
        End Get
        Set(ByVal v As String)
            sFieldName = v
        End Set
    End Property

    Public Property ProcessField() As String
        Get
            ProcessField = sProcessField
        End Get
        Set(ByVal v As String)
            sProcessField = v
        End Set
    End Property

    Public Property LinkedParam() As String
        Get
            LinkedParam = sLinkedParam
        End Get
        Set(ByVal v As String)
            sLinkedParam = v
        End Set
    End Property

    Public Property ParamValue() As String
        Get
            ParamValue = sParamValue
        End Get
        Set(ByVal v As String)
            sParamValue = v
        End Set
    End Property
End Class

Public Class ActionDefns
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As ActionDefn)
        List.Add(Parm)
    End Sub

    Public Function Add(ByVal Name As String, _
                    ByVal Process As String, _
                    ByVal Validate As Boolean, _
                    ByVal RowBased As Boolean, _
                    ByVal CloseObj As String, _
                    ByVal IsDblClick As Boolean, _
                    ByVal ImageFile As String, _
                    ByVal ToolTip As String) As ActionDefn

        Return Add(Name, Process, Validate, RowBased, CloseObj, IsDblClick, _
                ImageFile, ToolTip, "N", "", "", Nothing, 0, "", "y||n", "", "")
    End Function

    Public Function Add(ByVal Name As String, _
                    ByVal Process As String, _
                    ByVal Validate As Boolean, _
                    ByVal RowBased As Boolean, _
                    ByVal CloseObj As String, _
                    ByVal IsDblClick As Boolean, _
                    ByVal ImageFile As String, _
                    ByVal ToolTip As String, _
                    ByVal MenuType As String, _
                    ByVal MenuText As String, _
                    ByVal Parent As String, _
                    ByVal Rules As ActionRuleDefns, _
                    ByVal KeyCode As Integer, _
                    ByVal FieldName As String, _
                    ByVal ProcessField As String, _
                    ByVal LinkedParam As String, _
                    ByVal ParamValue As String) As ActionDefn
        Dim parm As New ActionDefn

        With parm
            .Name = Name
            .Process = Process
            .Validate = Validate
            .RowBased = RowBased
            .CloseObject = CloseObj
            .IsDblClick = IsDblClick
            If ImageFile <> "" Then
                .IsButton = True
            End If
            .ImageFile = ImageFile
            .ToolTip = ToolTip
            .MenuType = MenuType
            .MenuText = MenuText
            .Parent = Parent
            If KeyCode <> 0 Then
                .IsKey = True
                .KeyCode = KeyCode
            End If
            .Rules = Rules
            .FieldName = FieldName
            .ProcessField = ProcessField
            .LinkedParam = LinkedParam
            .ParamValue = ParamValue
        End With
        Me.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As ActionDefn
        Get
            Return CType(List.Item(Index), ActionDefn)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ActionDefn
        Get
            For Each ic As ActionDefn In Me
                If ic.Name = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public Class ActionRule
    ' EQ - equal
    ' NE - not equal
    ' NN - not null
    ' GT - greater than
    ' GE - greater than or equal
    ' LT - less than
    ' LE - less than or equal
    ' VL - valid
    ' LE - not valid
    Enum ValidationType
        EQ = 0
        NE = 1
        NN = 2
        GT = 3
        GE = 4
        LT = 5
        LE = 6
        VL = 7
        NV = 8
    End Enum

    Private iID As Integer
    Private sFieldName As String
    Private oType As ValidationType
    Private oValue As Object

    Public Property ID() As Integer
        Get
            ID = iID
        End Get
        Set(ByVal v As Integer)
            iID = v
        End Set
    End Property

    Public Property FieldName() As String
        Get
            FieldName = sFieldName
        End Get
        Set(ByVal v As String)
            sFieldName = v
        End Set
    End Property

    Public Property Type() As ValidationType
        Get
            Type = oType
        End Get
        Set(ByVal v As ValidationType)
            oType = v
        End Set
    End Property

    Public Property Value() As Object
        Get
            Value = oValue
        End Get
        Set(ByVal v As Object)
            oValue = v
        End Set
    End Property
End Class

Public Class ActionRules
    Inherits CollectionBase

    Public Function Add(ByVal ID As Integer, _
                        ByVal FieldName As String, _
                        ByVal Type As ActionRule.ValidationType, _
                        ByVal Value As Object) As ActionRule
        Dim parm As New ActionRule

        With parm
            .ID = ID
            .FieldName = FieldName
            .Type = Type
            .Value = Value
        End With
        List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ActionRule
        Get
            For Each ic As ActionRule In Me
                If ic.FieldName = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public Class ActionRuleDefn
    Private sName As String
    Private oRules As New ActionRules

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Rules() As ActionRules
        Get
            Rules = oRules
        End Get
        Set(ByVal v As ActionRules)
            oRules = v
        End Set
    End Property
End Class

Public Class ActionRuleDefns
    Inherits CollectionBase

    Public Function Add(ByVal ID As Integer, _
                        ByVal Name As String, _
                        ByVal FieldName As String, _
                        ByVal Type As ActionRule.ValidationType, _
                        ByVal Value As Object) As ActionRuleDefn
        Dim parm As ActionRuleDefn = Nothing
        Dim b As Boolean = True
        Dim rule As New ActionRule

        For Each obj As Object In Me
            parm = CType(obj, ActionRuleDefn)
            If parm.Name = Name Then
                b = False
                Exit For
            End If
        Next
        If b Then
            parm = New ActionRuleDefn
            parm.Name = Name
            List.Add(parm)
        End If

        With rule
            .ID = ID
            .FieldName = FieldName
            .Type = Type
            .Value = Value
        End With
        parm.Rules.Add(ID, FieldName, Type, Value)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As ActionRuleDefn
        Get
            Return CType(List.Item(Index), ActionRuleDefn)
        End Get
    End Property
End Class

Public Class ActionProcessRuleDefn
    Private oValue As Object
    Private sProcess As String

    Public Property Value() As Object
        Get
            Value = oValue
        End Get
        Set(ByVal v As Object)
            oValue = v
        End Set
    End Property

    Public Property Process() As String
        Get
            Process = sProcess
        End Get
        Set(ByVal v As String)
            sProcess = v
        End Set
    End Property
End Class

Public Class ActionProcessRuleDefns
    Inherits CollectionBase

    Public Function Add(ByVal Process As String, _
                        ByVal Value As Object) As ActionProcessRuleDefn
        Dim parm As New ActionProcessRuleDefn

        parm.Process = Process
        parm.Value = Value
        List.Add(parm)

        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As ActionProcessRuleDefn
        Get
            Return CType(List.Item(Index), ActionProcessRuleDefn)
        End Get
    End Property
End Class

Public Class ActionState
    Private sAction As String
    Private bEnabled As Boolean

    Public Property Action() As String
        Get
            Action = sAction
        End Get
        Set(ByVal v As String)
            sAction = v
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Enabled = bEnabled
        End Get
        Set(ByVal v As Boolean)
            bEnabled = v
        End Set
    End Property
End Class

Public Class ActionStates
    Inherits CollectionBase

    Public Function Add(ByVal Action As String, _
                        ByVal Enabled As Boolean) As ActionState
        Dim parm As New ActionState

        parm.Action = Action
        parm.Enabled = Enabled
        List.Add(parm)

        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As ActionState
        Get
            Return CType(List.Item(Index), ActionState)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ActionState
        Get
            For Each ic As ActionState In Me
                If ic.Action = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class
