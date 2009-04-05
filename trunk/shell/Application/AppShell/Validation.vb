Option Explicit On 
Option Strict On

Public Class ValidationDefn
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

    Public Name As String
    Public FieldName As String
    Public AssociatedFields() As String
    Public Process As String
    Public Type As ValidationType
    Public ValueType As ValType
    Public Value As Object
    Public Message As String
    Public ReturnParameter As String
End Class

Public Class ValidationDefns
#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ActionsEnum(Keys, Values)
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
        Return CType(Values.Item(Name), ValidationDefn)
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
