Option Explicit On
Option Strict On
'ProcessDefn / ProcessDefns
' contains the definitions of the processes able to be run in the system
' ProcessDefns is the collection class for ProcessDefn
' The ProcessDefns.item returns a ProcessDefn object using 
' the ProcessDefn.Name as the key.
' Using Foreach ... next on the ProcessDefns class returns 
' the ProcessDefn in the order they were created.

Public Class ProcessDefn
    Private sName As String
    Private sSuccessProcess As String
    Private sFailProcess As String
    Private sConfirmMsg As String
    Private bUpdateParent As Boolean
    Private bSuspendParent As Boolean
    Private sObjectKey As String
    Private bLoadVariables As Boolean

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property SuccessProcess() As String
        Get
            SuccessProcess = sSuccessProcess
        End Get
        Set(ByVal v As String)
            sSuccessProcess = v
        End Set
    End Property

    Public Property FailProcess() As String
        Get
            FailProcess = sFailProcess
        End Get
        Set(ByVal v As String)
            sFailProcess = v
        End Set
    End Property

    Public Property ConfirmMsg() As String
        Get
            ConfirmMsg = sConfirmMsg
        End Get
        Set(ByVal v As String)
            sConfirmMsg = v
        End Set
    End Property

    Public Property UpdateParent() As Boolean
        Get
            UpdateParent = bUpdateParent
        End Get
        Set(ByVal v As Boolean)
            bUpdateParent = v
        End Set
    End Property

    Public Property SuspendParent() As Boolean
        Get
            SuspendParent = bSuspendParent
        End Get
        Set(ByVal v As Boolean)
            bSuspendParent = v
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

    Public Property LoadVariables() As Boolean
        Get
            LoadVariables = bLoadVariables
        End Get
        Set(ByVal v As Boolean)
            bLoadVariables = v
        End Set
    End Property
End Class

Public Class ProcessDefns
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As ProcessDefn)
        Me.List.Add(Parm)
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
        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ProcessDefn
        Get
            For Each pd As ProcessDefn In Me
                If pd.Name = Name Then
                    Return pd
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class
