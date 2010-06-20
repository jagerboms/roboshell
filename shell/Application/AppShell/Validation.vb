Option Explicit On 
Option Strict On

Public Class ValidationDefn
    Private sName As String
    Private sFieldName As String
    Private sAssociatedFields() As String
    Private sProcess As String
    Private oType As ValidationType
    Private oValueType As ValType
    Private oValue As Object
    Private sMessage As String
    Private sReturnParameter As String

    Enum ValidationType
        EQ = 0  ' equal
        NE = 1  ' not equal
        GT = 3  ' greater than
        GE = 4  ' greater than or equal
        LT = 5  ' less than
        LE = 6  ' less than or equal
    End Enum

    Enum ValType
        Process = 0
        Constant = 1
        Field = 2
    End Enum

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
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

    Public Property AssociatedFields() As String()
        Get
            AssociatedFields = sAssociatedFields
        End Get
        Set(ByVal v As String())
            sAssociatedFields = v
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

    Public Property Type() As ValidationType
        Get
            Type = oType
        End Get
        Set(ByVal v As ValidationType)
            oType = v
        End Set
    End Property

    Public Property ValueType() As ValType
        Get
            ValueType = oValueType
        End Get
        Set(ByVal v As ValType)
            oValueType = v
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

    Public Property Message() As String
        Get
            Message = sMessage
        End Get
        Set(ByVal v As String)
            sMessage = v
        End Set
    End Property

    Public Property ReturnParameter() As String
        Get
            ReturnParameter = sReturnParameter
        End Get
        Set(ByVal v As String)
            sReturnParameter = v
        End Set
    End Property
End Class

Public Class ValidationDefns
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As ValidationDefn)
        Me.List.Add(Parm)
    End Sub

    Public Function Add(ByVal Name As String, _
                    ByVal FieldName As String, _
                    ByVal Process As String, _
                    ByVal Type As ValidationDefn.ValidationType, _
                    ByVal ValueType As ValidationDefn.ValType, _
                    ByVal Value As Object, _
                    ByVal Message As String, _
                    ByVal RetParameter As String) As ValidationDefn
        Dim parm As New ValidationDefn

        With parm
            .Name = Name
            .FieldName = FieldName
            .Process = Process
            .Type = Type
            .ValueType = ValueType
            .Value = Value
            .Message = Message
            .ReturnParameter = RetParameter
        End With
        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ValidationDefn
        Get
            For Each vd As ValidationDefn In Me
                If vd.Name = Name Then
                    Return vd
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class
