Option Explicit On 
'Option Strict On
'ShellProcess / ShellObject

Public Class ShellProcess
    Private sProcessKey As String
    Public oOwner As Object              '' close / parameters / errors
    Public Success As Boolean = False
    Public WithEvents ProcessObject As ShellObject

    Public Sub New(ByVal ProcessKey As String, ByVal Owner As Object, _
                                                ByVal parms As ShellParameters)
        Dim p As ProcessDefn
        Me.sProcessKey = ProcessKey
        Me.oOwner = Owner
        If LCase(sProcessKey) = "null" Then
            Success = True
            Exit Sub
        End If
        p = Processes.Item(sProcessKey)
        If p Is Nothing Then
            Dim Msgs As New ShellMessages
            Msgs.Add("E", "Invalid process '" & sProcessKey & _
                                          "' no definition exists")
            oOwner.MsgOut(Msgs)
            Exit Sub
        End If
        If p.UpdateParent Then
            oOwner.Suspend(True)
        End If
        If p.ConfirmMsg <> "" Then
            If MsgBox(p.ConfirmMsg, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                If p.UpdateParent Then
                    Me.oOwner.Suspend(False)
                End If
                Success = False
                Exit Sub
            End If
        End If
        If p.LoadVariables Then
            Publics.GetVars()
        End If
        ProcessObject = Objects.Item(p.ObjectKey).Create()
        If ProcessObject Is Nothing Then
            Owner.ProcessError(ProcessObject.Messages)
            ProcessObject_ExitFail()
        Else
            ProcessObject.Parent = Owner
            Dim n As ShellProcess
            ProcessObject.Update(parms)
            If p.SuccessProcess <> "" Then
                n = New ShellProcess(p.SuccessProcess, Me.oOwner, ProcessObject.Parms)
            End If
        End If
    End Sub

    Private Sub ProcessObject_ExitOkay() Handles ProcessObject.ExitOkay
        Dim p As ProcessDefn
        Success = True
        p = Processes.Item(sProcessKey)
        If ProcessObject.Messages.count > 0 Then
            Try
                Me.oOwner.ProcessError(ProcessObject.Messages)
            Catch
                Dim s As String = ""
                For Each e As ShellMessage In ProcessObject.Messages
                    s &= e.Message & vbCrLf
                Next
                Publics.MessageOut(s)
            End Try
        End If
        For Each pp As ShellProperty In _
                        CType(Objects.Item(p.ObjectKey), ObjectDefn).Properties
            If pp.Type = "sk" Then
                For Each ll As Listener In Register.Listen
                    If ll.Key = pp.Name And ll.KeyType = "B" Then
                        Register.Item(ll.ObjectKey).pObject.Listener(ProcessObject.Parms)
                    End If
                Next
            End If
        Next
        If p.UpdateParent Then
            Me.oOwner.Update(ProcessObject.Parms)
            Me.oOwner.Suspend(False)
        End If
    End Sub

    Private Sub ProcessObject_ExitFail() Handles ProcessObject.ExitFail
        Dim p As ProcessDefn
        Dim n As ShellProcess
        p = Processes.Item(sProcessKey)
        If ProcessObject.Messages.count > 0 Then
            Try
                Me.oOwner.ProcessError(ProcessObject.Messages)
            Catch
                Dim s As String = ""
                For Each e As ShellMessage In ProcessObject.Messages
                    s &= e.Message & vbCrLf
                Next
                Publics.MessageOut(s)
            End Try
        End If
        If p.FailProcess <> "" Then
            n = New ShellProcess(p.FailProcess, Me.oOwner, ProcessObject.Parms)
        Else
            If p.UpdateParent Then
                Try
                    Me.oOwner.Suspend(False)
                Catch ex As Exception
                    Dim i As Integer
                    i = 9
                End Try
            End If
        End If
    End Sub

    Private Sub ProcessObject_Progress(ByVal iProgress As Integer) _
                                            Handles ProcessObject.ProgressChange
        Dim p As ProcessDefn
        p = Processes.Item(sProcessKey)

        Try
            Me.oOwner.Progress(iProgress)
        Catch ex As Exception
            Publics.MessageOut("Process.ProcessObject_Progress: error" & vbCrLf & ex.ToString)
        End Try
    End Sub
End Class

