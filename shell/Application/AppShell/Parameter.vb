Option Explicit On 
Option Strict On
'Parameter / Parameters
' used for passing parameters between process class modules
' ShellParameters is the collection class for ShellParameter
' The ShellParameters.item returns a ShellParameter object using 
' the ShellParameter.Name as the key.
' Using Foreach ... next on the ShellParameters class returns 
' the ShellParameters in the order they were created.

Public Class shellParameter
    Private iIndex As Integer
    Private sName As String
    Private bInput As Boolean = True
    Private bOutput As Boolean = True
    Private iWidth As Integer = 0
    Private oValueType As System.Data.DbType = DbType.Object
    Private oValue As Object
    Private bInitialised As Boolean = False
    Private sInputText As String
    Private sField As String

    Public Property Index() As Integer
        Get
            Index = iIndex
        End Get
        Set(ByVal v As Integer)
            iIndex = v
        End Set
    End Property

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Input() As Boolean
        Get
            Input = bInput
        End Get
        Set(ByVal v As Boolean)
            bInput = v
        End Set
    End Property

    Public Property Output() As Boolean
        Get
            Output = bOutput
        End Get
        Set(ByVal v As Boolean)
            bOutput = v
        End Set
    End Property

    Public Property Width() As Integer
        Get
            Width = iWidth
        End Get
        Set(ByVal v As Integer)
            iWidth = v
        End Set
    End Property

    Public Property ValueType() As System.Data.DbType
        Get
            ValueType = oValueType
        End Get
        Set(ByVal v As System.Data.DbType)
            oValueType = v
        End Set
    End Property

    Public ReadOnly Property Initialised() As Boolean
        Get
            Initialised = bInitialised
        End Get
    End Property

    Public Property InputText() As String
        Get
            InputText = sInputText
        End Get
        Set(ByVal v As String)
            sInputText = v
        End Set
    End Property

    Public Property Field() As String
        Get
            Field = sField
        End Get
        Set(ByVal v As String)
            sField = v
        End Set
    End Property

    Public Property Value() As Object
        Get
            Value = oValue
        End Get
        Set(ByVal Value As Object)
            oValue = Value
            bInitialised = True
        End Set
    End Property
End Class

Public Class ShellParameters
    Inherits CollectionBase

    Public Sub Add(ByVal Parm As shellParameter)
        List.Add(Parm)
    End Sub

    Public Function Add(ByVal sName As String, _
                    ByVal Value As Object) As shellParameter
        Return Add(sName, Value, DbType.String, True, True, 0)
    End Function

    Public Function Add(ByVal sName As String, _
                    ByVal Value As Object, _
                    ByVal ValueType As System.Data.DbType, _
                    ByVal Input As Boolean, _
                    ByVal Output As Boolean, _
                    ByVal Width As Integer) As shellParameter
        Dim parm As New shellParameter

        With parm
            .Index = List.Count
            .Name = sName
            .Input = Input
            .Output = Output
            .Width = Width
            .ValueType = ValueType
            If Not IsDBNull(Value) Then
                .Value = Value
                .InputText = Publics.GetString(Value)
            End If
        End With
        Me.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As shellParameter
        Get
            Return CType(List.Item(Index), shellParameter)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As shellParameter
        Get
            For Each ic As shellParameter In Me
                If ic.Name = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property

    ' set all parameter values to nothing

    Public Function ClearValues() As Boolean
        Dim p As shellParameter

        For Each p In List
            p.Value = Nothing
            p.InputText = Nothing
        Next
        Return True
    End Function

    ' set all input parameters equal to the value of the same named 
    ' output parameters of passed in list

    Public Function MergeValues(ByRef cSource As ShellParameters) As Boolean
        Dim p As shellParameter
        Dim s As shellParameter
        Dim i As Integer

        If Not cSource Is Nothing Then
            For Each p In List
                If p.Input Then
                    s = cSource.Item(p.Name)
                    If Not s Is Nothing Then
                        If s.Output And s.Initialised Then
                            If p.ValueType = DbType.String And Not IsDBNull(s.Value) _
                                    And Not s.Value Is Nothing Then
                                p.Value = Mid(GetString(s.Value), 1, p.Width)
                                p.InputText = s.InputText
                                If s.Width <> p.Width Then
                                    i = 9
                                End If
                            Else
                                p.Value = s.Value
                                p.InputText = s.InputText
                            End If
                        Else
                            If s.Output And Not p.Value Is Nothing Then
                                i = 9
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return True
    End Function

    Public Sub Clone(ByRef parms As ShellParameters)
        Dim param As shellParameter

        For Each p As shellParameter In List
            param = New shellParameter
            With param
                .Index = p.Index
                .Name = p.Name
                .Input = p.Input
                .Output = p.Output
                .Width = p.Width
                .ValueType = p.ValueType
                .Field = p.Field
                If p.Initialised Then
                    .Value = p.Value
                    .InputText = p.InputText
                End If
            End With

            parms.Add(param)
        Next
    End Sub
End Class
