Option Explicit On 
Option Strict On

Public Class ShellMessage
    Public Type As String = "E"
    Public Message As String = ""

    Public Sub New(ByVal sType As String, ByVal Msg As String)
        Type = sType
        Message = Msg
    End Sub
End Class

Public Class ShellMessages

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ShellMsgEnum(Values)
    End Function

    Public Class ShellMsgEnum
        Implements IEnumerable, IEnumerator
        Private Values As New ArrayList
        Private EnumeratorPosition As Integer = -1

        Public Sub New(ByVal Val As ArrayList)
            Values = Val
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ShellMessage)
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

    Public Values As New ArrayList

    Public Function Add(ByVal Type As String, ByVal Msg As String) As ShellMessage
        Dim it As New ShellMessage(Type, Msg)
        Values.Add(it)
        Return it
    End Function

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class
