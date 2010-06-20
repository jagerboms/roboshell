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
    Inherits CollectionBase

    Public Function Add(ByVal Type As String, ByVal Msg As String) As ShellMessage
        Dim it As New ShellMessage(Type, Msg)
        Me.List.Add(it)
        Return it
    End Function
End Class
