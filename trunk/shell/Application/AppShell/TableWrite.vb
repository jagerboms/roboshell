Option Explicit On 
Option Strict On

Public Class TableWriteDefn
    Inherits ObjectDefn

    Public DataParameter As String
    Public TableWritePreProcess As String = ""
    Public RowWriteProcess As String
    Public TableWritePostProcess As String = ""
    Public ErrorProcess As String = ""

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
            Case "TableWritePreProcess"
                TableWritePreProcess = GetString(Value)
            Case "RowWriteProcess"
                RowWriteProcess = GetString(Value)
            Case "TableWritePostProcess"
                TableWritePostProcess = GetString(Value)
            Case "ErrorProcess"
                ErrorProcess = GetString(Value)
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

                If sDefn.TableWritePreProcess <> "" Then
                    bUpdateParameters = True
                    p = New ShellProcess(sDefn.TableWritePreProcess, Me, Me.Parms)
                    If Me.Messages.count > 0 Then
                        ErrorFixUp()
                        Me.OnExitFail()
                        Exit Sub
                    End If
                    bUpdateParameters = False
                End If

                Dim parm As shellParameter
                Dim dtData As DataTable = CType(Parms.Item(sDefn.DataParameter).Value, DataTable)
                'Add a row counter to the parameters list
                Me.Parms.Add("RowNumber", Nothing, DbType.Int32, False, True)

                Dim iRow As Integer = 1
                For Each dr As DataRow In dtData.Rows
                    Me.Parms.Item("RowNumber").Value = iRow

                    'Populate required fields with row data
                    For Each dc As DataColumn In dtData.Columns
                        parm = Me.Parms.Item(dc.ColumnName)
                        If Not parm Is Nothing Then
                            parm.Value = dr.Item(dc.ColumnName)
                        End If
                    Next

                    p = New ShellProcess(sDefn.RowWriteProcess, Me, Me.Parms)

                    If Me.Messages.count > 0 Then
                        ErrorFixUp()
                        Me.OnExitFail()
                        Exit Sub
                    End If
                    iRow += 1
                Next

                If sDefn.TableWritePostProcess <> "" Then
                    p = New ShellProcess(sDefn.TableWritePostProcess, Me, Me.Parms)

                    If Me.Messages.count > 0 Then
                        ErrorFixUp()
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

    Private Sub ErrorFixUp()
        If sDefn.ErrorProcess <> "" Then
            Dim p As New ShellProcess(sDefn.ErrorProcess, Me, Me.Parms)
        End If
    End Sub
End Class
