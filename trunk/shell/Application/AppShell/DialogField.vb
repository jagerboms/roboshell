Option Explicit On 
Option Strict On

Public Class DialogFieldItem
    Public Key As String
    Public Value As Object
    Public Text As String = ""
    Public Last As Object
    Public Errs As FieldErrors
    Public BusDateRelated As Boolean = False

    Public Sub New(ByVal sKey As String)
        Key = sKey
        Errs = New FieldErrors
    End Sub
End Class

Public Class DialogFieldItems
    Public Values As New Hashtable

    Public Function Add(ByVal Key As String) As DialogFieldItem
        Dim it As New DialogFieldItem(Key)
        Values.Add(Key, it)
        Return CType(Values.Item(Key), DialogFieldItem)
    End Function

    Public Sub Remove(ByVal key As String)
        Try
            Values.Remove(key)
        Catch
        End Try
    End Sub

    Public Sub Clear()
        Try
            Values = New Hashtable
        Catch
        End Try
    End Sub

    Public ReadOnly Property Item(ByVal key As String) As DialogFieldItem
        Get
            Try
                Return CType(Values.Item(key), DialogFieldItem)
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

Public Class DialogField
    Public Name As String
    Public Field As Field
    Public Actions() As String
    Public LinkedFields() As String
    Public Caption As String = ""
    Public LabelIndex As Integer = -1
    Public ControlIndex As Integer = -1
    Public Key As String = "x"
    Public items As DialogFieldItems

    Public Property Value() As Object
        Get
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                Value = Nothing
            Else
                Value = itm.Value
            End If
        End Get
        Set(ByVal Value As Object)
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            itm.Value = Value
        End Set
    End Property

    Public Property Text() As String
        Get
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                Text = ""
            Else
                Text = itm.Text
            End If
        End Get
        Set(ByVal Value As String)
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            itm.Text = Value
        End Set
    End Property

    Public Property Last() As Object
        Get
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                Last = Nothing
            Else
                Last = itm.Last
            End If
        End Get
        Set(ByVal Value As Object)
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            itm.Last = Value
        End Set
    End Property

    Public Property Errs() As FieldErrors
        Get
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            Errs = itm.Errs
        End Get
        Set(ByVal Value As FieldErrors)
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            itm.Errs = Value
        End Set
    End Property

    Public Property BusDateRelated() As Boolean
        Get
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                BusDateRelated = False
            Else
                BusDateRelated = itm.BusDateRelated
            End If
        End Get
        Set(ByVal Value As Boolean)
            Dim itm As DialogFieldItem = items.Item(Key)
            If itm Is Nothing Then
                itm = items.Add(Key)
            End If
            itm.BusDateRelated = Value
        End Set
    End Property

    Public Sub New(ByVal sName As String)
        Name = sName
        items = New DialogFieldItems
    End Sub
End Class

Public Class DialogFields
    Public Values As New Hashtable

    Friend Sub Add(ByVal Parm As DialogField)
        Values.Add(Parm.Name, Parm)
    End Sub

    Public Function Add(ByVal sName As String, _
                        ByVal fField As Field) As DialogField
        Dim parm As New DialogField(sName)

        With parm
            .Field = fField
        End With
        Me.Add(parm)
        Return CType(Values.Item(sName), DialogField)
    End Function

    Public ReadOnly Property Item(ByVal index As String) As DialogField
        Get
            Try
                Return CType(Values.Item(index), DialogField)
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

    Public Sub ClearItems()
        Dim itm As DialogField
        For Each o As DictionaryEntry In Values
            itm = CType(o.Value, DialogField)
            If Not itm Is Nothing Then
                itm.items = New DialogFieldItems
            End If
        Next
    End Sub

    Public Sub ClearKeyItem(ByVal Key As String)
        Dim itm As DialogField
        For Each o As DictionaryEntry In Values
            itm = CType(o.Value, DialogField)
            If Not itm Is Nothing Then
                If Not itm.items.Item(Key) Is Nothing Then
                    itm.items.Remove(Key)
                End If
            End If
        Next
    End Sub
End Class

Public Class FieldError
    Public ValidationName As String
    Public Message As String = ""

    Public Sub New(ByVal sName As String, ByVal sMessage As String)
        ValidationName = sName
        Message = sMessage
    End Sub
End Class

Public Class FieldErrors
    Public Values As New Hashtable

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
        Try
            Values = New Hashtable
        Catch
        End Try
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

Public Class Container
    Private sName As String
    Public Height As Integer = 0
    Public Width As Integer = 0
    Public Index As Integer = 0

    Public Sub New(ByVal Name As String)
        sName = Name
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Name = sName
        End Get
    End Property
End Class

Public Class Containers
    Public Values As New Hashtable

    Public Function Add(ByVal Name As String) As Container
        Dim parm As New Container(Name)
        Values.Add(Name, parm)
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal index As String) As Container
        Get
            Try
                Return CType(Values.Item(index), Container)
            Catch
                Return Nothing
            End Try
        End Get
    End Property
End Class
