Option Explicit On 
Option Strict On

Public Class ActionDefn
    Public Name As String
    Public Process As String
    Public Enabled As Boolean = True
    Public Checked As Boolean = False
    Public RowBased As Boolean = False
    Public Validate As Boolean = False
    Public CloseObject As String
    Public Rules As ActionRuleDefns
    Public Processes As ActionProcessRuleDefns

    Public IsDblClick As Boolean = False
    Public IsButton As Boolean = False
    Public ImageFile As String
    Public ToolTip As String

    Public MenuType As String    ' (N)one, (I)tem, (S)ub menu
    Public MenuText As String
    Public Parent As String  ' identifies the menu object for sub menus

    Public IsKey As Boolean
    Public KeyCode As Integer
    Public Shift As String

    Public FieldName As String      ' Action is fired when field data is changed by user
    Public ProcessField As String   ' Field used to select process to call
    Public LinkedParam As String    ' Parameter linked to button state
    Public ParamValue As String
End Class

Public Class ActionDefns
#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        If Sorted Then
            Return New ActionsEnum(SKeys, Values)
        Else
            Return New ActionsEnum(Keys, Values)
        End If
    End Function

    Public Class ActionsEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Hashtable
        Dim Keys() As String
        Private EnumeratorPosition As Integer = -1

        Public Sub New(ByVal aKeys() As String, ByVal Hash As Hashtable)
            Keys = aKeys
            Values = Hash
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(Keys(EnumeratorPosition)), ActionDefn)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition < Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = -1
        End Sub
    End Class
#End Region

    Private Values As New Hashtable
    Private Keys() As String
    Private SKeys() As String
    Public Sorted As Boolean = False

    Public Sub Add(ByVal Parm As ActionDefn)
        Dim i As Integer

        i = Values.Count
        ReDim Preserve Keys(i)
        ReDim Preserve SKeys(i)
        Values.Add(Parm.Name, Parm)
        Keys(i) = Parm.Name
        SKeys(i) = Parm.Name
        Array.Sort(SKeys)
    End Sub

    Public Function Add(ByVal Name As String, _
                    ByVal Process As String, _
                    ByVal Validate As Boolean, _
                    ByVal RowBased As Boolean, _
                    ByVal CloseObj As String, _
                    ByVal IsDblClick As Boolean, _
                    ByVal ImageFile As String, _
                    ByVal ToolTip As String, _
                    Optional ByVal MenuType As String = "N", _
                    Optional ByVal MenuText As String = "", _
                    Optional ByVal Parent As String = "", _
                    Optional ByVal Rules As ActionRuleDefns = Nothing, _
                    Optional ByVal KeyCode As Integer = 0, _
                    Optional ByVal FieldName As String = "", _
                    Optional ByVal ProcessField As String = "y||n", _
                    Optional ByVal LinkedParam As String = "", _
                    Optional ByVal ParamValue As String = "") As ActionDefn
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

    Public ReadOnly Property Item(ByVal index As Object) As ActionDefn
        Get
            Try
                Return CType(Values.Item(index), ActionDefn)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
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

    Public ID As Integer
    Public FieldName As String
    Public Type As ValidationType
    Public Value As Object
End Class

Public Class ActionRules

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActionRulesEnum(Values)
    End Function

    Public Class ActionRulesEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal col As Collection)
            Values = col
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ActionRule)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition <= Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection

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
        Values.Add(parm, parm.ID.ToString)
        Return CType(Values.Item(ID.ToString), ActionRule)
    End Function

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class

Public Class ActionRuleDefn
    Public Name As String
    Public Rules As New ActionRules
End Class

Public Class ActionRuleDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActionRulesEnum(Values)
    End Function

    Public Class ActionRulesEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal col As Collection)
            Values = col
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ActionRuleDefn)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition <= Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection

    'Public Sub Add(ByVal Parm As ActionRuleDefn)
    '    Values.Add(Parm.Name, Parm)
    'End Sub

    Public Function Add(ByVal ID As Integer, _
                        ByVal Name As String, _
                        ByVal FieldName As String, _
                        ByVal Type As ActionRule.ValidationType, _
                        ByVal Value As Object) As ActionRuleDefn
        Dim parm As ActionRuleDefn = Nothing
        Dim rule As New ActionRule

        For Each obj As Object In Values
            If CType(obj, ActionRuleDefn).Name = Name Then
                parm = CType(obj, ActionRuleDefn)
                Exit For
            End If
        Next
        If parm Is Nothing Then
            parm = New ActionRuleDefn
            parm.Name = Name
            Values.Add(parm, Name)
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

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class

Public Class ActionProcessRuleDefn
    Public Value As Object
    Public Process As String
End Class

Public Class ActionProcessRuleDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActProcRulesEnum(Values)
    End Function

    Public Class ActProcRulesEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal col As Collection)
            Values = col
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ActionProcessRuleDefn)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition <= Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection

    Public Function Add(ByVal Process As String, _
                        ByVal Value As Object) As ActionProcessRuleDefn
        Dim parm As New ActionProcessRuleDefn

        parm.Process = Process
        parm.Value = Value
        Values.Add(parm, parm.Value.ToString)

        Return parm
    End Function

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class

Public Class ActionState
    Public Action As String
    Public Enabled As Boolean
End Class

Public Class ActionStates

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActStatesEnum(Values)
    End Function

    Public Class ActStatesEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal col As Collection)
            Values = col
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ActionState)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition <= Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection

    Public Function Add(ByVal Action As String, _
                        ByVal Enabled As Boolean) As ActionState
        Dim parm As New ActionState

        parm.Action = Action
        parm.Enabled = Enabled
        Values.Add(parm, parm.Action)

        Return CType(Values.Item(Action), ActionState)
    End Function

    Public ReadOnly Property Item(ByVal index As String) As ActionState
        Get
            Try
                Return CType(Values.Item(index), ActionState)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class
