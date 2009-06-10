Option Explicit On
Option Strict On

Imports System.Windows.Forms

Public Class SQLFile
    Private sFile As String
    Private sFileType As String
    Private sSource As String
    Private bExists As Boolean
    Private sState As String = "U"
    Private sResults As String
    Private nNode As TreeNode

    Public Sub New(ByVal File As String, ByVal FileType As String, ByVal Source As String)
        sFile = File
        sFileType = FileType
        sSource = Source

        Select Case FileType
            Case "SQL", "DB"
                bExists = True
            Case "DBE"
                sFileType = "DB"
                bExists = False
                sState = "E"
            Case "SCL"
                bExists = False
                sState = "E"
            Case Else
                If Dir(File) = "" Then
                    bExists = False
                    sState = "E"
                Else
                    bExists = True
                End If
        End Select

        If Not bExists Then
            sResults = "File '" & File & "' does not exist."
        End If

        bExists = Exists
        nNode = New TreeNode(Me.Name)
        nNode.Tag = File
        SetNode()
    End Sub

    Public Sub New(ByVal File As String, ByVal Source As String, ByVal Exists As Boolean, ByVal State As String, ByVal sError As String)
        sFile = File
        sFileType = "DB"
        sSource = Source
        bExists = Exists
        sState = State
        sResults = sError

        nNode = New TreeNode(Me.Name)
        nNode.Tag = File
        SetNode()
    End Sub

    Public ReadOnly Property File() As String
        Get
            File = sFile
        End Get
    End Property

    Public ReadOnly Property FileType() As String
        Get
            FileType = sFileType
        End Get
    End Property

    Public ReadOnly Property Source() As String
        Get
            Source = sSource
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Name = System.IO.Path.GetFileName(sFile)
        End Get
    End Property

    Public ReadOnly Property Exists() As Boolean
        Get
            Exists = bExists
        End Get
    End Property

    Public Property State() As String
        Get
            State = sState
        End Get
        Set(ByVal value As String)
            sState = value
            SetNode()
        End Set
    End Property

    Public Property Results() As String
        Get
            Results = sResults
        End Get
        Set(ByVal value As String)
            sResults = value
        End Set
    End Property

    Public Property Node() As TreeNode
        Get
            Node = nNode
        End Get
        Set(ByVal value As TreeNode)
            nNode = value
        End Set
    End Property

    Private Sub SetNode()
        If nNode Is Nothing Then
            Exit Sub
        End If

        Select Case sFileType
            Case "DB"
                Select Case sState
                    Case "C"
                        nNode.ImageIndex = 1
                        nNode.SelectedImageIndex = 1
                        nNode.ForeColor = Drawing.Color.DarkGreen
                    Case "E"
                        nNode.ImageIndex = 2
                        nNode.SelectedImageIndex = 2
                        nNode.ForeColor = Drawing.Color.DarkRed
                    Case Else
                        nNode.ImageIndex = 0
                        nNode.SelectedImageIndex = 0
                        nNode.ForeColor = Drawing.Color.Navy
                End Select
            Case "SCL"
                nNode.ImageIndex = 3
                nNode.SelectedImageIndex = 3
                nNode.ForeColor = Drawing.Color.DarkRed
            Case "FILE"
                Select Case sState
                    Case "C"
                        nNode.ImageIndex = 9
                        nNode.SelectedImageIndex = 9
                        nNode.ForeColor = Drawing.Color.DarkGreen
                    Case "E"
                        nNode.ImageIndex = 10
                        nNode.SelectedImageIndex = 10
                        nNode.ForeColor = Drawing.Color.DarkRed
                    Case Else
                        nNode.ImageIndex = 8
                        nNode.SelectedImageIndex = 8
                        nNode.ForeColor = Drawing.Color.Navy
                End Select
            Case "SQL"
                Select Case sState
                    Case "C"
                        nNode.ImageIndex = 5
                        nNode.SelectedImageIndex = 5
                        nNode.ForeColor = Drawing.Color.DarkGreen
                    Case "E"
                        nNode.ImageIndex = 6
                        nNode.SelectedImageIndex = 6
                        nNode.ForeColor = Drawing.Color.DarkRed
                    Case Else
                        nNode.ImageIndex = 4
                        nNode.SelectedImageIndex = 4
                        nNode.ForeColor = Drawing.Color.DarkOrange
                End Select
            Case Else
                nNode.ImageIndex = 7
                nNode.SelectedImageIndex = 7
                nNode.ForeColor = Drawing.Color.DarkRed
        End Select
    End Sub
End Class

Public Class SQLFiles

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New PropertyEnum(Keys, Values)
    End Function

    Public Class PropertyEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Hashtable
        Dim Keys As ArrayList
        Private EnumeratorPosition As Integer = -1

        Public Sub New(ByVal aKeys As ArrayList, ByVal Hash As Hashtable)
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), SQLFile)
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
    Private Keys As New ArrayList

    Public Function Add(ByVal File As String, ByVal Type As String, ByVal Source As String) As SQLFile
        Dim parm As New SQLFile(File, Type, Source)
        Values.Add(File, parm)
        Keys.Add(File)
        Return CType(Values.Item(File), SQLFile)
    End Function

    Public Function Add(ByVal File As String, ByVal Source As String, ByVal Exists As Boolean, ByVal State As String, ByVal sError As String) As SQLFile
        Dim parm As New SQLFile(File, Source, Exists, State, sError)
        Values.Add(File, parm)
        Keys.Add(File)
        Return CType(Values.Item(File), SQLFile)
    End Function

    Public ReadOnly Property Item(ByVal File As String) As SQLFile
        Get
            Try
                Return CType(Values.Item(File), SQLFile)
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
