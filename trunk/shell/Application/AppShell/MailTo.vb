Option Explicit On 
Option Strict On

Public Class MailToDefn
    Inherits ObjectDefn

    Public email As String
    Public cc As String
    Public bcc As String
    Public subject As String
    Public body As String

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New MailTo(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "email"
                email = GetString(Value)
            Case "cc"
                cc = GetString(Value)
            Case "bcc"
                bcc = GetString(Value)
            Case "subject"
                subject = GetString(Value)
            Case "body"
                body = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by MailTo object")
        End Select
    End Sub
End Class

Public Class MailTo
    Inherits ShellObject

    Private sDefn As MailToDefn

    Public Sub New(ByVal Defn As MailToDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim Email As String = ""
        Dim Subject As String = ""
        Dim cc As String = ""
        Dim bcc As String = ""
        Dim Body As String = ""
        Dim b As Boolean = False

        Try
            Me.Parms.MergeValues(Parms)
            For Each p As shellParameter In Me.Parms
                If p.Input Then
                    If p.Name = sDefn.email Then
                        Email = GetString(p.Value.ToString)
                    ElseIf p.Name = sDefn.subject Then
                        Subject = GetString(p.Value.ToString)
                    ElseIf p.Name = sDefn.cc Then
                        cc = GetString(p.Value.ToString)
                    ElseIf p.Name = sDefn.bcc Then
                        bcc = GetString(p.Value.ToString)
                    ElseIf p.Name = sDefn.body Then
                        Body = GetString(p.Value.ToString)
                    End If
                End If
            Next

            Dim s As String = "mailto:" & Email
            If Subject <> "" Then
                s &= "?Subject=" & System.Web.HttpUtility.UrlEncode(Subject)
                b = True
            End If
            If cc <> "" Then
                If b Then
                    s &= "&"
                Else
                    s &= "?"
                    b = True
                End If
                s &= "cc=" & System.Web.HttpUtility.UrlEncode(cc)
            End If
            If bcc <> "" Then
                If b Then
                    s &= "&"
                Else
                    s &= "?"
                    b = True
                End If
                s &= "bcc=" & System.Web.HttpUtility.UrlEncode(bcc)
            End If
            If Body <> "" Then
                If b Then
                    s &= "&"
                Else
                    s &= "?"
                    b = True
                End If
                s &= "body=" & System.Web.HttpUtility.UrlEncode(Body)
            End If

            s = s.Replace("+", "%20")
            Dim myProcess As New Process
            myProcess.StartInfo.FileName = s
            myProcess.Start()
            Me.OnExitOkay()

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
        End Try
    End Sub
End Class
