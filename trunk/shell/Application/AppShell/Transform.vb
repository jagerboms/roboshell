Option Explicit On 
Option Strict On

'Transform object allows a parameter collection to be transformed
' the natural transformation is via the objects own parameter list
' which only accepts matching parameters from the input list and only
' outputs parameters marked as output.
' an output parameter (not input) will be created using a default value
' taken from the class definition.
' As well the class includes an array for parameter transforms. The two
' dimensional array identifies the source and destination parameters. 
' The transform process sets the value of the destination using the 
' value of the source parameter. Note the transformations are performed
' in the order defined by their index in the array. This mechanism is
' used to rename parameters.
'
' eg.
' transform p2 -> p4 *
'                                   --value--
'             Name Def  input out   in    out
'             P1    U     y    y    A     A
'            *P2    V     y    n    B     -
'             P3    W     y    n    C     -
'            *P4    X     n    y    -     B
'             P5    Y     n    n    -     -
'             P6    Z     n    y    -     Z

Public Class TransformDefn
    Inherits ObjectDefn

    Private sChoiceParameter As String
    Private sProcessParameter As String
    Private sNotFoundProcess As String

    Public ReadOnly Property ChoiceParameter() As String
        Get
            ChoiceParameter = sChoiceParameter
        End Get
    End Property

    Public ReadOnly Property ProcessParameter() As String
        Get
            ProcessParameter = sProcessParameter
        End Get
    End Property

    Public ReadOnly Property NotFoundProcess() As String
        Get
            NotFoundProcess = sNotFoundProcess
        End Get
    End Property

    Public Sub New(ByVal sName As String)
        MyBase.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New Transform(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)
        Select Case Name
            Case "ChoiceParameter"
                sChoiceParameter = GetString(Value)
            Case "NotFoundProcess"
                sNotFoundProcess = GetString(Value)
            Case "ProcessParameter"
                sProcessParameter = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Transform object")
        End Select
    End Sub
End Class

Public Class Transform
    Inherits ShellObject

    Private sDefn As TransformDefn

    Public Sub New(ByVal Defn As TransformDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim p As ShellProperty
        Dim ip As shellParameter
        Dim op As shellParameter
        Dim s As String
        Dim i As Integer
        Dim dt As DataTable

        Try
            Me.Parms.MergeValues(Parms)
            For Each p In sDefn.Properties
                Select Case p.Type
                    Case "tr"   '' transform parameter
                        ip = Me.Parms.Item(p.Name)
                        op = Me.parms.Item(GetString(p.Value))
                        If op.ValueType = DbType.String Then
                            op.Value = Mid(GetString(ip.Value), 1, op.Width)
                        Else
                            op.Value = ip.Value
                        End If

                    Case "rp"  '' Datatable row to new param.
                        ip = Me.Parms.Item(p.Name)
                        Dim vt As System.Data.DbType
                        If Not ip Is Nothing Then
                            dt = CType(ip.Value, DataTable)
                            If Not dt Is Nothing Then
                                For Each dr As DataRow In dt.Rows
                                    op = Me.parms.Item(GetString(dr("ParameterName")))
                                    If op Is Nothing Then
                                        op = New shellParameter
                                        op.Name = Publics.GetString(dr("ParameterName"))
                                        Me.Parms.Add(op)
                                    End If
                                    Try
                                        s = GetString(dr("ValueType"))
                                        vt = GetValueType(s)
                                        op.ValueType = vt
                                    Catch ex As Exception
                                        i = 9
                                    End Try
                                    Try
                                        i = CInt(dr("Width"))
                                        op.Width = i
                                    Catch ex As Exception
                                        i = 9
                                    End Try
                                    op.Value = dr("Value")
                                Next
                            End If
                        End If
                End Select
            Next

            If sDefn.ProcessParameter <> "" Then
                If Not MyBase.Parms.Item(sDefn.ProcessParameter) Is Nothing Then
                    s = GetString(MyBase.Parms.Item(sDefn.ProcessParameter).Value)
                    If s <> "" Then
                        Dim n As New ShellProcess(s, Me, MyBase.Parms)
                    End If
                End If
            End If

            If sDefn.ChoiceParameter <> "" Then
                If Not MyBase.Parms.Item(sDefn.ChoiceParameter) Is Nothing Then
                    s = GetString(MyBase.Parms.Item(sDefn.ChoiceParameter).Value)
                    If Not sDefn.Properties.Item(s, "cp") Is Nothing Then
                        s = GetString(sDefn.Properties.Item(s, "cp").Value)
                    Else
                        s = sDefn.NotFoundProcess
                    End If
                    If s <> "" Then
                        Dim n As New ShellProcess(s, Me, MyBase.Parms)
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
            Exit Sub
        End Try
    End Sub
End Class
