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

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActionsCollection(Keys, Values)
    End Function

    Public Class ActionsCollection
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), ValidationDefn)
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

    Public Sub Add(ByVal Parm As ValidationDefn)
        Dim i As Integer

        i = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(Parm.Name, Parm)
        Keys(i) = Parm.Name
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
        Me.Add(parm)
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As ValidationDefn
        Get
            Try
                Return CType(Values.Item(index), ValidationDefn)
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
