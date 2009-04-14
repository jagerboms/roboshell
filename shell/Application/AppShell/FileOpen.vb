Option Explicit On
Option Strict On

Public Class FileOpenDefn
    Inherits ObjectDefn

    Public Title As String
    Public TitleParameter() As String
    Public InitialDirectory As String = "c:\"
    Public Filter As String = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
    Public FilterIndex As Integer = 2
    Public Multiselect As Boolean = True
    Public OutputParameter As String

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New FileOpen(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "Title"
                Title = GetString(Value)
            Case "TitleParameters"
                TitleParameter = Split(GetString(Value), "||")
            Case "InitialDirectory"
                InitialDirectory = GetString(Value)
            Case "Filter"
                Filter = GetString(Value)
            Case "FilterIndex"
                FilterIndex = CInt(GetString(Value))
            Case "Multiselect"
                Multiselect = CType(IIf(GetString(Value) = "Y", True, False), Boolean)
            Case "OutputParameter"
                OutputParameter = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by FileOpen object")
        End Select
    End Sub
End Class

Public Class FileOpen
    Inherits ShellObject

    Private sDefn As FileOpenDefn

    Public Sub New(ByVal Defn As FileOpenDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Me.Parms.MergeValues(Parms)
            Dim s As String = ""
            Dim ss As String
            Dim openFileDialog As New System.Windows.Forms.OpenFileDialog

            openFileDialog.Title = GetTitle()
            openFileDialog.InitialDirectory = sDefn.InitialDirectory
            openFileDialog.Filter = sDefn.Filter
            openFileDialog.FilterIndex = sDefn.FilterIndex
            openFileDialog.Multiselect = sDefn.Multiselect
            openFileDialog.RestoreDirectory = True

            If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                For Each ss In openFileDialog.FileNames
                    s &= "||" & ss
                Next
                Me.parms.Item(sDefn.OutputParameter).Value = Mid(s, 3)
                Me.OnExitOkay()
            Else
                Me.OnExitFail()
                Me.SuccessFlag = False
            End If
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                Me.Messages.Add("E", ex.ToString)
            Else
                Dim ex2 As Exception = ex.InnerException
                Do While Not ex2 Is Nothing
                    Me.Messages.Add("E", ex2.ToString)
                    ex2 = ex2.InnerException
                Loop
            End If
            Me.OnExitFail()
            Me.SuccessFlag = False
        End Try
    End Sub

    Private Function GetTitle() As String
        Dim s As String
        Dim ss As String
        Dim sTemp As String

        s = sDefn.Title
        If Not sDefn.TitleParameter Is Nothing Then
            For Each ss In sDefn.TitleParameter
                sTemp = GetString(parms.Item(ss).Value)
                If sTemp <> "" Then
                    s &= " - " & sTemp
                End If
            Next
        End If
        Return s
    End Function
End Class
