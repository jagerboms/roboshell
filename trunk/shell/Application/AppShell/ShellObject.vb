Option Explicit On 
'Option Strict On

Public MustInherit Class ObjectDefn
    Private sName As String
    Private oParms As New ShellParameters
    Private oFields As New Fields
    Private oActions As New ActionDefns
    Private oValidations As New ValidationDefns
    Private oProperties As New shellProperties

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Parms() As ShellParameters
        Get
            Parms = oParms
        End Get
        Set(ByVal v As ShellParameters)
            oParms = v
        End Set
    End Property

    Public Property Fields() As Fields
        Get
            Fields = oFields
        End Get
        Set(ByVal v As Fields)
            oFields = v
        End Set
    End Property

    Public Property Actions() As ActionDefns
        Get
            Actions = oActions
        End Get
        Set(ByVal v As ActionDefns)
            oActions = v
        End Set
    End Property

    Public Property Validations() As ValidationDefns
        Get
            Validations = oValidations
        End Get
        Set(ByVal v As ValidationDefns)
            oValidations = v
        End Set
    End Property

    Public Property Properties() As shellProperties
        Get
            Properties = oProperties
        End Get
        Set(ByVal v As shellProperties)
            oProperties = v
        End Set
    End Property

    Public MustOverride Sub SetProperty(ByVal Name As String, ByVal Value As Object)

    Protected Sub New()
    End Sub
End Class

Public Class ObjectDefns
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As ObjectDefn)
        Me.List.Add(Parm)
    End Sub

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As Object 'Defn
        Get
            For Each od As ObjectDefn In Me
                If od.Name = Name Then
                    Return od
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public MustInherit Class ShellObject
    Private sObjectType As String
    Private sRegKey As String
    Private oMessages As New ShellMessages
    Private oParent As Object
    Private bSuccessFlag As Boolean = True
    Private MyParams As New ShellParameters

    Public Event ExitOkay()
    Public Event ExitFail()

    Public Property ObjectType() As String
        Get
            ObjectType = sObjectType
        End Get
        Set(ByVal v As String)
            sObjectType = v
        End Set
    End Property

    Public Property RegKey() As String
        Get
            RegKey = sRegKey
        End Get
        Set(ByVal v As String)
            sRegKey = v
        End Set
    End Property

    Public Property Messages() As ShellMessages
        Get
            Messages = oMessages
        End Get
        Set(ByVal v As ShellMessages)
            oMessages = v
        End Set
    End Property

    Public Property Parent() As Object
        Get
            Parent = oParent
        End Get
        Set(ByVal v As Object)
            oParent = v
        End Set
    End Property

    Public Property SuccessFlag() As Boolean
        Get
            SuccessFlag = bSuccessFlag
        End Get
        Set(ByVal v As Boolean)
            bSuccessFlag = v
        End Set
    End Property

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

    Protected Sub New()
    End Sub
End Class
