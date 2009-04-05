Option Explicit On
Option Strict On
Imports System.Data.SqlClient

Public Class StoredProcDefn
    Inherits ObjectDefn

    Public ProcName As String
    Public ConnectKey As String
    Public Mode As String
    Public DataParameter() As String
    Public Messages As Boolean = True

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New StoredProc(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case LCase(Name)
            Case "procname"
                ProcName = GetString(Value)
            Case "connectkey"
                ConnectKey = GetString(Value)
            Case "mode"
                Mode = GetString(Value)
            Case "dataparameter"
                ''DataParameter = GetString(Value)
                DataParameter = Split(GetString(Value), "||")
            Case "messages"
                Messages = (GetString(Value) = "Y")
            Case Else
                Publics.MessageOut(Name & " property is not supported by Stored Procedure object")
        End Select
    End Sub
End Class

Public Class StoredProc
    Inherits ShellObject

    Private sDefn As StoredProcDefn
    Private spParams() As SqlParameter

    Public Sub New(ByVal defn As StoredProcDefn)
        sDefn = defn
        sDefn.Parms.Clone(MyBase.parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim i As Integer
        Dim ret As Integer
        Dim pm As shellParameter
        Dim psConn As SqlConnection
        Dim DS As DataSet = Nothing
        Dim psAdapt As SqlDataAdapter

        Try
            Me.Parms.MergeValues(Parms)
            psConn = New SqlConnection(Publics.GetConnectString(sDefn.ConnectKey))
            AddHandler psConn.InfoMessage, AddressOf SQLMessages
            psConn.Open()
            psAdapt = New SqlDataAdapter(sDefn.ProcName, psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            If spParams Is Nothing Then
                SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
                ReDim spParams(psAdapt.SelectCommand.Parameters.Count)
                i = 0
                For Each p As SqlParameter In psAdapt.SelectCommand.Parameters
                    spParams(i) = p
                    i += 1
                Next
            Else
                For Each p As SqlParameter In spParams
                    psAdapt.SelectCommand.Parameters.Add(p)
                Next
            End If

            For Each p As SqlParameter In psAdapt.SelectCommand.Parameters
                If p.Direction = ParameterDirection.Input _
                Or p.Direction = ParameterDirection.InputOutput Then
                    pm = Me.parms.Item(Mid(p.ParameterName, 2))
                    If Not pm Is Nothing Then
                        If pm.Input Then
                            p.Value = pm.Value
                        End If
                    End If
                End If
            Next

            If sDefn.Mode = "X" Then
                psAdapt.SelectCommand.ExecuteNonQuery()
            Else
                DS = New DataSet
                psAdapt.Fill(DS)

                i = 0
                If Not sDefn.DataParameter Is Nothing Then
                    For Each s As String In sDefn.DataParameter
                        If Not Me.parms.Item(s) Is Nothing Then
                            Me.parms.Item(s).Value = DS.Tables(i)  'place results in output parameters
                        End If
                        i += 1
                        If DS.Tables.Count >= i Then
                            Exit For
                        End If
                    Next
                End If
            End If

            ret = CInt(psAdapt.SelectCommand.Parameters.Item("@RETURN_VALUE").Value)
            For Each p As SqlParameter In psAdapt.SelectCommand.Parameters
                If p.Direction = ParameterDirection.Output _
                Or p.Direction = ParameterDirection.InputOutput Then
                    pm = Me.Parms.Item(Mid(p.ParameterName, 2))
                    If Not pm Is Nothing Then
                        If pm.Output Then
                            pm.Value = p.Value
                        End If
                    End If
                End If
            Next

            If sDefn.Mode = "P" Then
                Dim dt As DataTable = DS.Tables(0)
                If Not dt Is Nothing Then
                    For Each p As DataColumn In dt.Columns
                        pm = Me.parms.Item(p.ColumnName)
                        If Not pm Is Nothing Then
                            If pm.Output Then
                                If dt.Rows.Count = 0 Then
                                    pm.Value = Nothing
                                Else
                                    pm.Value = dt.Rows(0).Item(p.ColumnName)
                                    If IsDBNull(pm.Value) Then
                                        pm.Value = Nothing
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            If ret = 0 Then
                Me.OnExitOkay()
            Else
                Me.OnExitFail()
            End If
            psConn.Close()

        Catch ex As SqlException
            For i = 0 To ex.Errors.Count - 1
                If ex.Number = 50999 Then
                    Me.Messages.Add("U", ex.Message)
                Else
                    Me.Messages.Add("E", ex.Message & " [" & ex.Number & "]")
                End If
            Next i
            Me.OnExitFail()

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

    Private Sub SQLMessages(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        If sDefn.Messages Then
            Me.Messages.Add("M", e.Message)
        End If
    End Sub
End Class
