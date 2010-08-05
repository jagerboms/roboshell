Option Explicit On 
Option Strict On

Public Class ObjectRegister
    Private sKey As String
    Private oObject As ShellObject

    Public Property Key() As String
        Get
            Key = sKey
        End Get
        Set(ByVal v As String)
            sKey = v
        End Set
    End Property

    Public Property pObject() As ShellObject
        Get
            pObject = oObject
        End Get
        Set(ByVal v As ShellObject)
            oObject = v
        End Set
    End Property
End Class

Public Class ObjectRegisters

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ObjRegCollection(Values)
    End Function

    Public Class ObjRegCollection
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal Coll As Collection)
            Values = Coll
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), ObjectRegister)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition < Values.Count + 1)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection
    Private iCount As Integer = 1
    Private oListen As New Listeners

    Public Property Listen() As Listeners
        Get
            Listen = oListen
        End Get
        Set(ByVal v As Listeners)
            oListen = v
        End Set
    End Property

    Public Sub Add(ByVal Parm As ObjectRegister)
        Values.Add(Parm, Parm.Key)
    End Sub

    Public Function Add(ByVal pObject As ShellObject) As ObjectRegister
        Dim parm As New ObjectRegister
        Dim sKey As String

        sKey = "LK" & iCount
        iCount += 1
        With parm
            .Key = sKey
            .pObject = pObject
        End With
        Me.Add(parm)
        Return CType(Values.Item(sKey), ObjectRegister)
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As ObjectRegister
        Get
            Try
                Return CType(Values.Item(index), ObjectRegister)
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

    Public Sub Remove(ByVal sKey As String)
        Listen.Remove(sKey)
        Values.Remove(sKey)
    End Sub
End Class

Public Class Listener
    Private iIndex As Integer
    Private sKey As String
    Private sKeyType As String
    Private sObjectKey As String

    Public Property Index() As Integer
        Get
            Index = iIndex
        End Get
        Set(ByVal v As Integer)
            iIndex = v
        End Set
    End Property

    Public Property Key() As String
        Get
            Key = sKey
        End Get
        Set(ByVal v As String)
            sKey = v
        End Set
    End Property

    Public Property KeyType() As String
        Get
            KeyType = sKeyType
        End Get
        Set(ByVal v As String)
            sKeyType = v
        End Set
    End Property

    Public Property ObjectKey() As String
        Get
            ObjectKey = sObjectKey
        End Get
        Set(ByVal v As String)
            sObjectKey = v
        End Set
    End Property
End Class

Public Class Listeners

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New ListenerCollection(Values)
    End Function

    Public Class ListenerCollection
        Implements IEnumerable, IEnumerator
        Private Values As New Collection
        Private EnumeratorPosition As Integer = 0

        Public Sub New(ByVal Coll As Collection)
            Values = Coll
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(EnumeratorPosition), Listener)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition < Values.Count + 1)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = 0
        End Sub
    End Class
#End Region

    Private Values As New Collection

    Public Sub Add(ByVal Parm As Listener)
        Values.Add(Parm, Parm.Key & "||" & Parm.ObjectKey)
    End Sub

    Public Function Add(ByVal sKey As String, _
                    ByVal KeyType As String, _
                    ByVal ObjectKey As String) As Listener
        Dim parm As New Listener

        With parm
            .Index = Values.Count
            .Key = sKey
            .KeyType = KeyType
            .ObjectKey = ObjectKey
        End With
        Me.Add(parm)
        Return CType(Values.Item(sKey & "||" & ObjectKey), Listener)
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As Listener
        Get
            Try
                Return CType(Values.Item(index), Listener)
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

    Public Sub Remove(ByVal sObjectKey As String)
        For Each r As Listener In Values
            If r.ObjectKey = sObjectKey Then
                Values.Remove(r.Key & "||" & sObjectKey)
            End If
        Next
    End Sub
End Class
