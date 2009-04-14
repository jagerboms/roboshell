Option Explicit On
Option Strict On

Public Class DirectoryDefn
    Inherits ObjectDefn

    Public Title As String
    Public TitleParameter() As String
    Public InitialDirectory As String = "c:\"
    Public OutputParameter As String
    Public AllowNew As Boolean = True

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New Directory(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "Title"
                Title = GetString(Value)
            Case "TitleParameters"
                TitleParameter = Split(GetString(Value), "||")
            Case "InitialDirectory"
                InitialDirectory = GetString(Value)
            Case "OutputParameter"
                OutputParameter = GetString(Value)
            Case "AllowNew"
                AllowNew = CType(IIf(GetString(Value) = "Y", True, False), Boolean)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Directory object")
        End Select
    End Sub
End Class

Public Class Directory
    Inherits ShellObject

    Private sDefn As DirectoryDefn

    Public Sub New(ByVal Defn As DirectoryDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Me.parms.MergeValues(Parms)
            Dim FolderDialog As New System.Windows.Forms.FolderBrowserDialog

            FolderDialog.Description = GetTitle()
            FolderDialog.SelectedPath = sDefn.InitialDirectory
            FolderDialog.ShowNewFolderButton = sDefn.AllowNew
            If FolderDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.parms.Item(sDefn.OutputParameter).Value = FolderDialog.SelectedPath
                Me.OnExitOkay()
            Else
                Me.SuccessFlag = False
                Me.OnExitFail()
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
            Me.SuccessFlag = False
            Me.OnExitFail()
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
