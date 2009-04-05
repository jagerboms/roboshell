Option Explicit On 
Option Strict On

Imports System.Threading 'Namespace for Thread class

Public Class CallOutDefn
    Inherits ObjectDefn

    Public ClassName As String
    Public MethodName As String
    Public ReturnParamName As String
    Public ProgressProperty As String

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New CallOut(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "ClassName"
                ClassName = GetString(Value)
            Case "MethodName"
                MethodName = GetString(Value)
            Case "ReturnParamName"
                ReturnParamName = GetString(Value)
            Case "ProgressProperty"
                ProgressProperty = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by CallOut object")
        End Select
    End Sub
End Class

Public Class CallOut
    Inherits ShellObject

    Private sDefn As CallOutDefn
    Private obj As Object
    Delegate Sub Del(ByVal d As Double)
    Private dProgress As Double

    Public Sub New(ByVal Defn As CallOutDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Me.Parms.MergeValues(Parms)
            GetData()
            'Me.OnExitOkay()
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

    Private Sub GetData()
        Dim i As Integer
        Dim inProperty(0) As Object
        Dim bAddParameter As Boolean
        Dim objOut As Object
        Dim paParams() As Object = Nothing
        Dim objType As System.Reflection.EventInfo
        Dim evtType As System.Type

        Try
            obj = CreateObject(sDefn.ClassName)

            For Each pr As ShellProperty In sDefn.Properties
                If pr.Type = "ip" Then
                    inProperty(0) = Me.Parms.Item(pr.Name).Value
                    Try
                        CallByName(obj, CType(pr.Value, String), CallType.Let, inProperty)
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
                        Exit Sub
                    End Try
                End If
            Next

            i = 0
            For Each p As shellParameter In Me.Parms
                If p.Input Then
                    bAddParameter = True

                    For Each pr As ShellProperty In sDefn.Properties
                        If pr.Type = "ip" And p.Name = pr.Name Then
                            bAddParameter = False
                            Exit For
                        End If
                    Next

                    If bAddParameter Then
                        ReDim Preserve paParams(i)
                        paParams(i) = p.Value
                        i += 1
                    End If
                End If
            Next

            Try
                dProgress = -1
                objType = obj.GetType().GetEvent("Progress")
                If Not objType Is Nothing Then
                    evtType = objType.EventHandlerType
                    objType.AddEventHandler(obj, Del.CreateDelegate(evtType, Me, "handler"))
                End If

                If paParams Is Nothing Then
                    objOut = CallByName(obj, sDefn.MethodName, CallType.Method)
                Else
                    objOut = CallByName(obj, sDefn.MethodName, CallType.Method, paParams)
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
                Exit Sub
            End Try

            If sDefn.ReturnParamName <> "" Then
                If Me.Parms.Item(sDefn.ReturnParamName).ValueType = DbType.Date Then
                    'fix return value
                    If CType(objOut, DateTime) = System.DateTime.MinValue Then
                        Me.Parms.Item(sDefn.ReturnParamName).Value = Nothing    'of date value
                    Else
                        Me.Parms.Item(sDefn.ReturnParamName).Value = objOut
                    End If
                Else
                    Me.Parms.Item(sDefn.ReturnParamName).Value = objOut
                End If
            End If

            For Each pr As ShellProperty In sDefn.Properties
                If pr.Type = "op" Then
                    Try
                        Me.Parms.Item(pr.Name).Value = _
                            CallByName(obj, CType(pr.Value, String), CallType.Get)

                        If Me.Parms.Item(pr.Name).ValueType = DbType.Date Then
                            If CType(Me.Parms.Item(pr.Name).Value, DateTime) _
                                                    = System.DateTime.MinValue Then
                                Me.Parms.Item(pr.Name).Value = Nothing
                            End If
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
                        Exit Sub
                    End Try
                End If
            Next

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

    Private Sub handler(ByVal d As Double)
        Dim j As Integer

        If dProgress < 0 Then
            If d > 0 Then
                dProgress = d
            End If
        Else
            'scale to 100
            j = CType(100 * d / dProgress, Integer)
            Me.OnProgress(j)
        End If
        Application.DoEvents()
    End Sub
End Class
