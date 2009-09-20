Option Explicit On 
Option Strict On
'ShellField / ShellFields
' used for defining the column of a grid or inputs on a dialog.

Public Class Field
    Public Index As Integer
    Public Name As String
    Public Label As String
    Public Width As Integer
    Public DisplayType As String    ' (T)ext, (L)abel, (D)(r)opdown list, (C)heck, (H)idden ...
    Public FillProcess As String    ' process
    Public TextField As String      '    to
    Public ValueField As String     '       fill
    Public LinkColumn As String     '          dropdown
    Public LinkField As String      '             list
    Public DisplayWidth As Integer
    Public DisplayHeight As Integer
    Public Format As String
    Public Primary As Boolean
    Public Justify As String        ' (L)eft, (R)ight, (C)enter or (D)efault
    Public Enabled As Boolean = True
    Public Required As Boolean
    Public Locate As String         ' (N)ormal, new (C)olumn, new (G)roup, (P)air
    Public ValueType As System.Data.DbType
    Public HelpText As String
    Public LabelWidth As Integer
    Public Decimals As Integer = -1
    Public NullText As String = ""
    Public Container As String = ""
    'Public Value As Object
    'Public eValidationType As colValidationType
    'Public lHelpContext As Long

    Public Sub Clone(ByRef NewField As Field)
        NewField = New Field
        With NewField
            .Index = Index
            .Name = Name
            .Label = Label
            .Width = Width
            .DisplayType = DisplayType
            .FillProcess = FillProcess
            .TextField = TextField
            .ValueField = ValueField
            .LinkColumn = LinkColumn
            .LinkField = LinkField
            .DisplayWidth = DisplayWidth
            .DisplayHeight = DisplayHeight
            .Format = Format
            .Primary = Primary
            .Justify = Justify
            .Enabled = Enabled
            .Required = Required
            .Locate = Locate
            .ValueType = ValueType
            .HelpText = HelpText
            .LabelWidth = LabelWidth
            .Decimals = Decimals
            .NullText = NullText
        End With
    End Sub
End Class

Public Class Fields

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New FieldsEnum(Keys, Values)
    End Function

    Public Class FieldsEnum
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), Field)
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

    Public Sub Add(ByRef Parm As Field)
        Dim s As String
        ReDim Preserve Keys(Parm.Index)
        s = LCase(Parm.Name)
        Values.Add(s, Parm)
        Keys(Parm.Index) = s
    End Sub

    Public Function Add(ByVal Name As String, _
                    ByVal Label As String, _
                    ByVal DisplayType As String, _
                    ByVal DisplayWidth As Integer, _
                    ByVal DisplayHeight As Integer, _
                    ByVal Format As String, _
                    ByVal Primary As Boolean, _
                    ByVal Justify As String, _
                    ByVal Required As Boolean, _
                    ByVal Locate As String, _
                    ByVal ValueType As System.Data.DbType, _
                    ByVal Width As Integer, _
                    ByVal LabelWidth As Integer, _
                    ByVal Decimals As Integer, _
                    ByVal NullText As String, _
                    ByVal HelpText As String) As Field
        Dim parm As New Field

        With parm
            .Index = Values.Count
            .Name = Name
            .Label = Label
            .DisplayType = DisplayType
            .DisplayWidth = CInt(DisplayWidth * 1.1)
            .DisplayHeight = DisplayHeight
            .Format = Format
            .Primary = Primary
            .Justify = Justify
            .Required = Required
            .Locate = Locate
            .ValueType = ValueType
            .Width = Width
            .LabelWidth = CInt(LabelWidth * 1.1)
            .Decimals = Decimals
            .NullText = NullText
            .HelpText = HelpText
        End With
        Me.Add(parm)
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As Field
        Get
            Try
                Return CType(Values.Item(index), Field)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(ByVal Name As String) As Field
        Get
            Try
                Return CType(Values.Item(LCase(Name)), Field)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub Clone(ByRef Fields As Fields)
        Dim param As Field = Nothing
        Dim f As Field

        For i As Integer = 0 To Keys.GetUpperBound(0)
            f = CType(Values.Item(Keys(i)), Field)
            f.Clone(param)
            Fields.Add(param)
        Next
    End Sub

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class

Public Class FieldValidationDefn
    ' EQ - equal
    ' NE - not equal
    ' NN - not null
    ' GT - greater than
    ' GE - greater than or equal
    ' LT - less than
    ' LE - less than or equal

    Enum ValidationType
        EQ = 0
        NE = 1
        NN = 2
        GT = 3
        GE = 4
        LT = 5
        LE = 6
    End Enum

    Public Name As String
    Public Type As ValidationType
    Public Value As Object
    Public FieldName As String
End Class

Public Class FieldValidationDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New FieldValidationEnum(Keys, Values)
    End Function

    Public Class FieldValidationEnum
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), FieldValidationDefn)
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

    Public Sub Add(ByVal Parm As FieldValidationDefn)
        Dim i As Integer

        i = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(Parm.Name, Parm)
        Keys(i) = Parm.Name
    End Sub

    Public Function Add(ByVal Name As String, _
                        ByVal FieldName As String, _
                        ByVal Type As FieldValidationDefn.ValidationType, _
                        ByVal Value As Object) As FieldValidationDefn
        Dim parm As New FieldValidationDefn

        With parm
            .Name = Name
            .FieldName = FieldName
            .Type = Type
            .Value = Value
        End With
        Me.Add(parm)
        Return CType(Values.Item(Name), FieldValidationDefn)
    End Function

    Public ReadOnly Property Item(ByVal index As Object) As Object
        Get
            Try
                Return Values.Item(index)
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
