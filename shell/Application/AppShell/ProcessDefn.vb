'ProcessDefn / ProcessDefns
' contains the definitions of the processes able to be run in the system
' ProcessDefns is the collection class for ProcessDefn
' The ProcessDefns.item returns a ProcessDefn object using 
' the ProcessDefn.Name as the key.
' Using Foreach ... next on the ProcessDefns class returns 
' the ProcessDefn in the order they were created.
Option Explicit On 
Option Strict On

Public Class ProcessDefn
    Public Name As String
    Public SuccessProcess As String
    Public FailProcess As String
    Public ConfirmMsg As String
    Public UpdateParent As Boolean
    Public SuspendParent As Boolean
    Public ObjectKey As String
    Public LoadVariables As Boolean
End Class

Public Class ProcessDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ProcessesDefnEnum(Keys, Values)
    End Function

    Public Class ProcessesDefnEnum
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), ProcessDefn)
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

    Public Sub Add(ByVal Parm As ProcessDefn)
        Dim i As Integer = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(Parm.Name, Parm)
        Keys(i) = Parm.Name
    End Sub

    Public Function Add(ByVal sName As String, _
                    ByVal SuccessProcess As String, _
                    ByVal FailProcess As String, _
                    ByVal ConfirmMsg As String, _
                    ByVal UpdateParent As String, _
                    ByVal ObjectKey As String, _
                    ByVal LoadVariables As Boolean) As ProcessDefn
        Dim parm As New ProcessDefn

        With parm
            .Name = sName
            .SuccessProcess = SuccessProcess
            .FailProcess = FailProcess
            .ConfirmMsg = ConfirmMsg
            Select Case Mid(LCase(UpdateParent), 1, 1)
                Case "y"
                    .UpdateParent = True
                    .SuspendParent = False
                Case "s"
                    .UpdateParent = True
                    .SuspendParent = True
                Case Else
                    .UpdateParent = False
                    .SuspendParent = False
            End Select
            .ObjectKey = ObjectKey
            .LoadVariables = LoadVariables
        End With
        Me.Add(parm)
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As ProcessDefn
        Get
            Try
                Return CType(Values.Item(index), ProcessDefn)
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
