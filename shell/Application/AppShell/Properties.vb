Option Explicit On 
Option Strict On

Public Class ShellProperty
    Public Name As String
    Public Type As String
    Public UserSpecific As Boolean
    Public Value As Object
End Class

Public Class shellProperties

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New PropertyEnum(Keys, Values)
    End Function

    Public Class PropertyEnum
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), ShellProperty)
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

    Public Sub Add(ByVal Parm As ShellProperty)
        Dim i As Integer

        i = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(Parm.Type & ":" & Parm.Name, Parm)
        Keys(i) = Parm.Type & ":" & Parm.Name
    End Sub

    Public Function Add(ByVal PropertyName As String, ByVal PropertyType As String, _
                    ByVal UserSpecified As Boolean, ByVal PropertyValue As Object) As ShellProperty
        Dim parm As New ShellProperty

        With parm
            .Name = PropertyName
            .Type = PropertyType
            .UserSpecific = UserSpecified
            .Value = PropertyValue
        End With
        Me.Add(parm)
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal Name As String, ByVal Type As String) As ShellProperty
        Get
            Try
                Return CType(Values.Item(Type & ":" & Name), ShellProperty)
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
