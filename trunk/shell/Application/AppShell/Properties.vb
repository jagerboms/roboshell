Option Explicit On 
Option Strict On

Public Class ShellProperty
    Private sName As String
    Private sType As String
    Private bUserSpecific As Boolean
    Private oValue As Object

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Type() As String
        Get
            Type = sType
        End Get
        Set(ByVal v As String)
            sType = v
        End Set
    End Property

    Public Property UserSpecific() As Boolean
        Get
            UserSpecific = bUserSpecific
        End Get
        Set(ByVal v As Boolean)
            bUserSpecific = v
        End Set
    End Property

    Public Property Value() As Object
        Get
            Value = oValue
        End Get
        Set(ByVal v As Object)
            oValue = v
        End Set
    End Property
End Class

Public Class shellProperties
    Inherits CollectionBase

    Public Sub zzAdd(ByVal Parm As ShellProperty)
        Me.List.Add(Parm)
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
        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Name As String, ByVal Type As String) As ShellProperty
        Get
            For Each ic As ShellProperty In Me
                If ic.Name = Name And ic.Type = Type Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class
