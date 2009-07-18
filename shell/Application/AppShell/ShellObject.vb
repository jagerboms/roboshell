Option Explicit On 
'Option Strict On

Public MustInherit Class ObjectDefn
    Public Name As String
    Public Parms As New ShellParameters
    Public Fields As New Fields
    Public Actions As New ActionDefns
    Public Validations As New ValidationDefns
    Public Properties As New shellProperties

    Public MustOverride Sub SetProperty(ByVal Name As String, ByVal Value As Object)
End Class

Public Class ObjectDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ObjectsEnum(Keys, Values)
    End Function

    Public Class ObjectsEnum
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), ObjectDefn)
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

    Public Sub Add(ByVal Parm As ObjectDefn)
        Dim i As Integer
        Dim s As String = LCase(Parm.Name)

        i = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(s, Parm)
        Keys(i) = s
    End Sub

    Public ReadOnly Property Item(ByVal index As String) As Object
        Get
            Try
                Return Values.Item(LCase(index))
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

Public MustInherit Class ShellObject
    Public Event ExitOkay()
    Public Event ExitFail()
    Public ObjectType As String
    Public RegKey As String
    Public Messages As New ShellMessages
    Public Parent As Object
    Public SuccessFlag As Boolean = True

    Private MyParams As New ShellParameters

    Public ReadOnly Property parms() As ShellParameters
        Get
            Return MyParams
        End Get
    End Property

    Public MustOverride Sub Update(ByVal Parms As ShellParameters)

    Public Overridable Sub Listener(ByVal Parms As ShellParameters)
    End Sub

    Public Overridable Sub Suspend(ByVal Mode As Boolean)
    End Sub

    Public Overridable Sub MsgOut(ByVal msgs As ShellMessages)
        If Parent Is Nothing Then
            Dim s As String = ""
            Dim sType As String = "I"

            For Each ss As ShellMessage In msgs
                s &= ss.Message & vbCrLf
                If ss.Type = "E" Then
                    sType = "E"
                End If
            Next
            Publics.MessageOut(s, sType)
        Else
            Parent.MsgOut(msgs)
        End If
    End Sub

    Friend Sub OnExitOkay()
        RaiseEvent ExitOkay()
    End Sub

    Friend Sub OnExitFail()
        RaiseEvent ExitFail()
    End Sub
End Class
