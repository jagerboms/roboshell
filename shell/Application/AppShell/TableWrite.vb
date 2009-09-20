Option Explicit On 
Option Strict On

Public Class TableWriteDefn
    Inherits ObjectDefn

    Public DataParameter As String
    Public PreWriteProcess As String = ""
    Public WriteProcess As String
    Public PostWriteProcess As String = ""

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New TableWrite(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case Name
            Case "DataParameter"
                DataParameter = GetString(Value)
            Case "PreWriteProcess"
                PreWriteProcess = GetString(Value)
            Case "WriteProcess"
                WriteProcess = GetString(Value)
            Case "PostWriteProcess"
                PostWriteProcess = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by TableWrite object")
        End Select
    End Sub
End Class

Public Class TableWrite
    Inherits ShellObject

    Private sDefn As TableWriteDefn
    Private bUpdateParameters As Boolean = False

    Public Sub New(ByVal Defn As TableWriteDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Me.Parms.MergeValues(Parms)

            If Not bUpdateParameters Then
                Dim p As ShellProcess

                If sDefn.PreWriteProcess <> "" Then
                    bUpdateParameters = True
                    p = New ShellProcess(sDefn.PreWriteProcess, Me, Me.parms)
                    If Me.Messages.count > 0 Then
                        Me.OnExitFail()
                        Exit Sub
                    End If
                    bUpdateParameters = False
                End If

                Dim parm As shellParameter
                Dim dtData As DataTable = CType(Parms.Item(sDefn.DataParameter).Value, DataTable)

                For Each dr As DataRow In dtData.Rows

                    'Populate required fields with row data
                    For Each dc As DataColumn In dtData.Columns
                        parm = Me.parms.Item(dc.ColumnName)
                        If Not parm Is Nothing Then
                            parm.Value = dr.Item(dc.ColumnName)
                        End If
                    Next

                    p = New ShellProcess(sDefn.WriteProcess, Me, Me.parms)

                    If Me.Messages.count > 0 Then
                        Me.OnExitFail()
                        Exit Sub
                    End If
                Next

                If sDefn.PostWriteProcess <> "" Then
                    p = New ShellProcess(sDefn.PostWriteProcess, Me, Me.parms)

                    If Me.Messages.count > 0 Then
                        Me.OnExitFail()
                        Exit Sub
                    End If
                End If
            End If

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
