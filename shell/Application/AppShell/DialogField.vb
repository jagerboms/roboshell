Option Explicit On 
Option Strict On

Public Class DialogField
    Private sName As String
    Private oField As Field
    Private sActions() As String
    Private sLinkedFields() As String
    Private sErrField As String
    Private oLabel As Label
    Private oControl As Control
    Private oValue As Object
    Private sText As String = ""
    Private oLast As Object
    Private oErrs As FieldErrors
    Private bBusDateRelated As Boolean = False

    Public ReadOnly Property Name() As String
        Get
            Name = sName
        End Get
    End Property

    Public ReadOnly Property Field() As Field
        Get
            Field = oField
        End Get
    End Property

    Public ReadOnly Property Actions() As String()
        Get
            Actions = sActions
        End Get
    End Property

    Public ReadOnly Property LinkedFields() As String()
        Get
            LinkedFields = sLinkedFields
        End Get
    End Property

    Public Property ErrField() As String
        Get
            ErrField = sErrField
        End Get
        Set(ByVal value As String)
            sErrField = value
        End Set
    End Property

    Public ReadOnly Property Label() As Label
        Get
            Label = oLabel
        End Get
    End Property

    Public ReadOnly Property Control() As Control
        Get
            Control = oControl
        End Get
    End Property

    Public Property Value() As Object
        Get
            Value = oValue
        End Get
        Set(ByVal Value As Object)
            oValue = Value
        End Set
    End Property

    Public Property Text() As String
        Get
            Text = sText
        End Get
        Set(ByVal Value As String)
            sText = Value
        End Set
    End Property

    Public Property Last() As Object
        Get
            Last = oLast
        End Get
        Set(ByVal Value As Object)
            oLast = Value
        End Set
    End Property

    Public Property Errs() As FieldErrors
        Get
            Errs = oErrs
        End Get
        Set(ByVal Value As FieldErrors)
            oErrs = Value
        End Set
    End Property

    Public Property BusDateRelated() As Boolean
        Get
            BusDateRelated = bBusDateRelated
        End Get
        Set(ByVal Value As Boolean)
            bBusDateRelated = Value
        End Set
    End Property

    Public Sub New(ByVal fField As Field)
        sName = fField.Name
        oField = fField
        oErrs = New FieldErrors
    End Sub

    Public Sub AddAction(ByVal Name As String)
        Dim i As Integer
        If sActions Is Nothing Then
            i = 0
        Else
            i = sActions.GetUpperBound(0) + 1
        End If
        ReDim Preserve sActions(i)
        sActions(i) = Name
    End Sub

    Public Sub AddLinkedField(ByVal Name As String)
        Dim i As Integer
        If sLinkedFields Is Nothing Then
            i = 0
        Else
            i = sLinkedFields.GetUpperBound(0) + 1
        End If
        ReDim Preserve sLinkedFields(i)
        sLinkedFields(i) = Name
    End Sub

    Public Sub AddLabel(ByRef lLabel As Label)
        oLabel = lLabel
    End Sub

    Public Sub AddControl(ByRef cControl As Control)
        oControl = cControl
    End Sub
End Class

Public Class DialogFields
    Inherits CollectionBase

    Public Function Add(ByVal fField As Field) As DialogField
        Dim parm As New DialogField(fField)

        Me.List.Add(parm)
        Return parm
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As DialogField
        Get
            For Each df As DialogField In Me
                If df.Name = Name Then
                    Return df
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public Class FieldError
    Private sValidationName As String
    Private sMessage As String = ""

    Public ReadOnly Property Message() As String
        Get
            Message = sMessage
        End Get
    End Property

    Public Sub New(ByVal Name As String, ByVal Message As String)
        sValidationName = Name
        sMessage = Message
    End Sub
End Class

Public Class FieldErrors
    Private Values As New Hashtable

    Public Function Add(ByVal Name As String, _
                        ByVal Message As String) As FieldError
        Dim parm As New FieldError(Name, Message)
        Values.Add(Name, parm)
        Return CType(Values.Item(Name), FieldError)
    End Function

    Public Sub Remove(ByVal index As String)
        Try
            Values.Remove(index)
        Catch
        End Try
    End Sub

    Public Sub Clear()
        Values = New Hashtable
    End Sub

    Public ReadOnly Property Item(ByVal index As String) As FieldError
        Get
            Try
                Return CType(Values.Item(index), FieldError)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Message() As String
        Get
            Try
                Dim s As String = ""
                For Each o As DictionaryEntry In Values
                    s = CType(o.Value, FieldError).Message
                    Exit For
                Next
                Return s
            Catch ex As Exception
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
