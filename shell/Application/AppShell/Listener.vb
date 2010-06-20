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
    Inherits CollectionBase

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
        Me.List.Add(Parm)
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
        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal sKey As String) As ObjectRegister
        Get
            For Each org As ObjectRegister In Me
                If org.Key = sKey Then
                    Return org
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public Sub Remove(ByVal sKey As String)
        Dim org As ObjectRegister

        Listen.Remove(sKey)
        org = Me.Item(sKey)
        If Not org Is Nothing Then
            Me.List.Remove(org)
        End If
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
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As Listener)
        Me.List.Add(Parm)
    End Sub

    Public Function Add(ByVal sKey As String, _
                    ByVal KeyType As String, _
                    ByVal ObjectKey As String) As Listener
        Dim parm As New Listener

        With parm
            .Index = Me.List.Count
            .Key = sKey
            .KeyType = KeyType
            .ObjectKey = ObjectKey
        End With
        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal sKey As String) As Listener
        Get
            For Each l As Listener In Me
                If l.Key = sKey Then
                    Return l
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public Sub Remove(ByVal sObjectKey As String)
        Dim l As Listener = Me.Item(sObjectKey)

        If Not l Is Nothing Then
            Me.List.Remove(l)
        End If
    End Sub
End Class
