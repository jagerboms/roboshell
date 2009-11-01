Option Explicit On 
Option Strict On

Public Class ShellMessage
    Private sType As String = "E"
    Private sMessage As String = ""

    Public ReadOnly Property Type() As String
        Get
            Type = sType
        End Get
    End Property

    Public ReadOnly Property Message() As String
        Get
            Message = sMessage
        End Get
    End Property

    Public Sub New(ByVal sType As String, ByVal Msg As String)
        sType = sType
        sMessage = Msg
    End Sub
End Class

Public Class ShellMessages

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ShellMsgCollection(Values)
    End Function

    Public Class ShellMsgCollection
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

    Private Values As New ArrayList

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
