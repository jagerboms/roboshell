Option Explicit On 
Option Strict On
'ShellField / ShellFields
' used for defining the column of a grid or inputs on a dialog.

Public Class Field
    Private iIndex As Integer
    Private sName As String
    Private sLabel As String
    Private iWidth As Integer
    Private sDisplayType As String    ' (T)ext, (L)abel, (D)(r)opdown list, (C)heck, (H)idden ...
    Private sFillProcess As String    ' process
    Private sTextField As String      '    to
    Private sValueField As String     '       fill
    Private sLinkColumn As String     '          dropdown
    Private sLinkField As String      '             list
    Private iDisplayWidth As Integer
    Private iDisplayHeight As Integer
    Private sFormat As String
    Private bPrimary As Boolean
    Private sJustify As String        ' (L)eft, (R)ight, (C)enter or (D)efault
    Private bEnabled As Boolean = True
    Private bRequired As Boolean
    Private sLocate As String         ' (N)ormal, new (C)olumn, new (G)roup, (P)air
    Private oValueType As System.Data.DbType
    Private sHelpText As String
    Private iLabelWidth As Integer
    Private iDecimals As Integer = -1
    Private sNullText As String = ""
    Private sContainer As String = ""

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

    Public Property Label() As String
        Get
            Label = sLabel
        End Get
        Set(ByVal v As String)
            sLabel = v
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

    Public Property DisplayType() As String
        Get
            DisplayType = sDisplayType
        End Get
        Set(ByVal v As String)
            sDisplayType = v
        End Set
    End Property

    Public Property FillProcess() As String
        Get
            FillProcess = sFillProcess
        End Get
        Set(ByVal v As String)
            sFillProcess = v
        End Set
    End Property

    Public Property TextField() As String
        Get
            TextField = sTextField
        End Get
        Set(ByVal v As String)
            sTextField = v
        End Set
    End Property

    Public Property ValueField() As String
        Get
            ValueField = sValueField
        End Get
        Set(ByVal v As String)
            sValueField = v
        End Set
    End Property

    Public Property LinkColumn() As String
        Get
            LinkColumn = sLinkColumn
        End Get
        Set(ByVal v As String)
            sLinkColumn = v
        End Set
    End Property

    Public Property LinkField() As String
        Get
            LinkField = sLinkField
        End Get
        Set(ByVal v As String)
            sLinkField = v
        End Set
    End Property

    Public Property DisplayWidth() As Integer
        Get
            DisplayWidth = iDisplayWidth
        End Get
        Set(ByVal v As Integer)
            iDisplayWidth = v
        End Set
    End Property

    Public Property DisplayHeight() As Integer
        Get
            DisplayHeight = iDisplayHeight
        End Get
        Set(ByVal v As Integer)
            iDisplayHeight = v
        End Set
    End Property

    Public Property Format() As String
        Get
            Format = sFormat
        End Get
        Set(ByVal v As String)
            sFormat = v
        End Set
    End Property

    Public Property Primary() As Boolean
        Get
            Primary = bPrimary
        End Get
        Set(ByVal v As Boolean)
            bPrimary = v
        End Set
    End Property

    Public Property Justify() As String
        Get
            Justify = sJustify
        End Get
        Set(ByVal v As String)
            sJustify = v
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Enabled = bEnabled
        End Get
        Set(ByVal v As Boolean)
            bEnabled = v
        End Set
    End Property

    Public Property Required() As Boolean
        Get
            Required = bRequired
        End Get
        Set(ByVal v As Boolean)
            bRequired = v
        End Set
    End Property

    Public Property Locate() As String
        Get
            Locate = sLocate
        End Get
        Set(ByVal v As String)
            sLocate = v
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

    Public Property HelpText() As String
        Get
            HelpText = sHelpText
        End Get
        Set(ByVal v As String)
            sHelpText = v
        End Set
    End Property

    Public Property LabelWidth() As Integer
        Get
            LabelWidth = iLabelWidth
        End Get
        Set(ByVal v As Integer)
            iLabelWidth = v
        End Set
    End Property

    Public Property Decimals() As Integer
        Get
            Decimals = iDecimals
        End Get
        Set(ByVal v As Integer)
            iDecimals = v
        End Set
    End Property

    Public Property NullText() As String
        Get
            NullText = sNullText
        End Get
        Set(ByVal v As String)
            sNullText = v
        End Set
    End Property

    Public Property Container() As String
        Get
            Container = sContainer
        End Get
        Set(ByVal v As String)
            sContainer = v
        End Set
    End Property

    Public Sub Clone(ByRef NewField As Field)
        NewField = New Field
        With NewField
            .Index = iIndex
            .Name = sName
            .Label = sLabel
            .Width = iWidth
            .DisplayType = sDisplayType
            .FillProcess = sFillProcess
            .TextField = sTextField
            .ValueField = sValueField
            .LinkColumn = sLinkColumn
            .LinkField = sLinkField
            .DisplayWidth = iDisplayWidth
            .DisplayHeight = iDisplayHeight
            .Format = sFormat
            .Primary = bPrimary
            .Justify = sJustify
            .Enabled = bEnabled
            .Required = bRequired
            .Locate = sLocate
            .ValueType = oValueType
            .HelpText = sHelpText
            .LabelWidth = iLabelWidth
            .Decimals = iDecimals
            .NullText = sNullText
            .Container = sContainer
        End With
    End Sub
End Class

Public Class Fields

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New FieldsCollection(Keys, Values)
    End Function

    Public Class FieldsCollection
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
    Private sName As String
    Private oType As ValidationType
    Private oValue As Object
    Private sFieldName As String

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

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal v As String)
            sName = v
        End Set
    End Property

    Public Property Type() As ValidationType
        Get
            Type = oType
        End Get
        Set(ByVal v As ValidationType)
            oType = ValidationType.EQ
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

    Public Property FieldName() As String
        Get
            FieldName = sFieldName
        End Get
        Set(ByVal v As String)
            sFieldName = v
        End Set
    End Property
End Class

Public Class FieldValidationDefns

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New FieldValidationCollection(Keys, Values)
    End Function

    Public Class FieldValidationCollection
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
