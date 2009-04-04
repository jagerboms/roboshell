Option Explicit On
Option Strict On

Imports System.Data.SqlClient

Public Class ProcedureColumn
    Public Index As Integer
    Public Name As String
    Public Type As String
    Public Length As Integer
    Public Precision As Integer
    Public Scale As Integer
    Public TestValue As String
    Public IsInput As Boolean
    Public IsOutput As Boolean
    Public TypeText As String
    Public vbType As String

    Public ReadOnly Property ConfigName() As String
        Get
            ConfigName = Mid(Name, 2)
        End Get
    End Property
End Class

Public Class StoredProcedure
    Public ProcedureName As String = ""
    Public ProcedureText As String

    Public ConfigName As String
    Public ModuleName As String
    Public ProcessName As String
    Public SuccessName As String
    Public Messages As Boolean = True
    Public ConfirmMsg As String = ""
    Public SeekKey As String = ""

    Private sMode As String = "D"
    Private PreLoad As Integer = -1
    Private Values As New Hashtable
    Private Keys(0) As String
    Private Results As New Hashtable
    Private ResultKeys(0) As String

    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal Name As String, ByVal sConnect As String)
        Dim s As String = "a"
        Dim sText As String
        Dim b As Boolean = False
        Dim sName As String
        Dim sType As String
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet
        Dim dr As DataRow

        PreLoad = 2
        psConn = New SqlConnection(sConnect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        s = "select SPECIFIC_NAME"
        s &= ",ORDINAL_POSITION"
        s &= ",PARAMETER_NAME"
        s &= ",DATA_TYPE"
        s &= ",CHARACTER_MAXIMUM_LENGTH"
        s &= ",NUMERIC_PRECISION"
        s &= ",NUMERIC_SCALE"
        s &= ",PARAMETER_MODE "
        s &= "from INFORMATION_SCHEMA.PARAMETERS "
        s &= "where SPECIFIC_NAME='" & Name
        s &= "' order by ORDINAL_POSITION"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(Details)

        If Details.Tables(0).Rows.Count = 0 Then
            PreLoad = 3
            Return
        End If

        For Each dr In Details.Tables(0).Rows        ' Columns
            If ProcedureName = "" Then
                ProcedureName = GetString(dr.Item("SPECIFIC_NAME"))
            End If
            sName = GetString(dr.Item("PARAMETER_NAME"))
            sType = GetString(dr.Item("DATA_TYPE"))
            AddParameter(sName, sType, dr.Item("CHARACTER_MAXIMUM_LENGTH"), _
                dr.Item("NUMERIC_PRECISION"), dr.Item("NUMERIC_SCALE"), _
                    LCase(GetString(dr.Item("PARAMETER_MODE"))))
        Next

        s = "select text from syscomments "
        s &= "where id = object_id('" & Name & "') order by number, colid"
        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(Details)
        psConn.Close()

        sText = ""
        For Each dr In Details.Tables(0).Rows        ' Columns
            s = GetString(dr.Item("text"))
            If Len(s) < 4000 Then s &= vbCrLf
            sText &= s ' & vbCrLf
        Next
        sText = sText.Replace(vbCrLf, Chr(13))
        sText = sText.Replace(Chr(10), Chr(13))
        sText = sText.Replace(Chr(13), vbCrLf)
        Do While 1 = 1
            Select Case Mid(sText, 1, 1)
                Case " ", Chr(9), Chr(10), Chr(13)
                    sText = Mid(sText, 2)
                Case Else
                    Exit Do
            End Select
        Loop

        Do While 1 = 1
            Select Case Right(sText, 1)
                Case " ", Chr(9), Chr(10), Chr(13)
                    sText = Mid(sText, 1, Len(sText) - 1)
                Case Else
                    Exit Do
            End Select
        Loop
        sText &= vbCrLf

        ProcedureText = sText
    End Sub

    Public ReadOnly Property State() As Integer
        Get
            Return PreLoad
        End Get
    End Property

    Public Property Mode() As String
        Get
            Mode = sMode
        End Get
        Set(ByVal value As String)
            If value = "X" Or value = "D" Or value = "P" Then
                sMode = value
            Else
                Throw New Exception("Unsupported mode")
            End If
        End Set
    End Property

    Public Function AddParameter(ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object, _
        ByVal InOut As String) As ProcedureColumn

        Dim parm As New ProcedureColumn

        If sType = "int" Then
            sType = "integer"
            parm.vbType = "integer"
        End If

        With parm
            .Index = Values.Count
            .Name = sName
            .Type = sType
            .IsInput = InStr(InOut, "in") > 0
            .IsOutput = InStr(InOut, "out") > 0
        End With

        parm.TypeText = sType
        If sType = "char" Or sType = "varchar" Or sType = "nvarchar" Then
            parm.Length = CType(oLength, Integer)
            parm.TypeText = sType & "(" & parm.Length & ")"
            parm.vbType = "string"
        End If

        If sType = "decimal" Then
            parm.Precision = CType(oPrecision, Integer)
            parm.Scale = CType(oScale, Integer)
            parm.TypeText = sType & "(" & parm.Precision & "," & parm.Scale & ")"
            parm.vbType = "double"
        End If

        If sType = "datetime" Then
            parm.vbType = "datetime"
        End If

        If sType = "smalldatetime" Then
            parm.vbType = "date"
        End If

        If sType = "sysname" Then
            parm.vbType = "string"
            parm.Length = 128
        End If

        If parm.vbType = "" Then
            parm.vbType = sType
        End If

        '"currency"
        '"date"
        '"datetime"
        '"double"
        '"integer"
        '"int64"
        '"string"
        '"object"

        If parm.Index > Keys.GetUpperBound(0) Then
            ReDim Preserve Keys(parm.Index)
        End If
        Values.Add(sName, parm)
        Keys(parm.Index) = sName
        Return CType(Values.Item(sName), ProcedureColumn)
    End Function

    Public Function AddResult(ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object) As ProcedureColumn

        Dim parm As New ProcedureColumn

        If sType = "int" Then
            sType = "integer"
            parm.vbType = "integer"
        End If

        With parm
            .Index = Values.Count
            .Name = sName
            .Type = sType
            .IsInput = False
            .IsOutput = True
        End With

        parm.TypeText = sType
        If sType = "char" Or sType = "varchar" Or sType = "nvarchar" Then
            parm.Length = CType(oLength, Integer)
            parm.TypeText = sType & "(" & parm.Length & ")"
            parm.vbType = "string"
        End If

        If sType = "decimal" Then
            parm.Precision = CType(oPrecision, Integer)
            parm.Scale = CType(oScale, Integer)
            parm.TypeText = sType & "(" & parm.Precision & "," & parm.Scale & ")"
            parm.vbType = "double"
        End If

        If sType = "datetime" Then
            parm.vbType = "datetime"
        End If

        If sType = "smalldatetime" Then
            parm.vbType = "date"
        End If

        If sType = "sysname" Then
            parm.vbType = "string"
            parm.Length = 128
        End If

        If parm.vbType = "" Then
            parm.vbType = sType
        End If

        '"currency"
        '"date"
        '"datetime"
        '"double"
        '"integer"
        '"int64"
        '"string"
        '"object"

        If parm.Index > Keys.GetUpperBound(0) Then
            ReDim Preserve Keys(parm.Index)
        End If
        Values.Add(sName, parm)
        Keys(parm.Index) = sName
        Return CType(Values.Item(sName), ProcedureColumn)
    End Function

    Public Function ConfigText() As String
        Dim sOut As String
        Dim s As String
        Dim b As Boolean = True
        Dim pc As ProcedureColumn

        sOut = "execute dbo.shlStoredProcInsert" & vbCrLf
        sOut &= "    @objectname = '" & ConfigName & "'" & vbCrLf
        sOut &= "   ,@procname = '" & ProcedureName & "'" & vbCrLf
        If sMode = "D" Then
            sOut &= "   ,@dataparameter = '" & ConfigName & "'" & vbCrLf
        Else
            sOut &= "   ,@mode = '" & sMode & "'" & vbCrLf
        End If
        If Not Messages Then
            sOut &= "   ,@messages = 'N'" & vbCrLf
        End If
        If ProcessName <> "" Then
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "---------------------------------------------------" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlProcessesInsert" & vbCrLf
            sOut &= "    @ProcessName = '" & ProcessName & "'" & vbCrLf
            sOut &= "   ,@ModuleID = '" & ModuleName & "'" & vbCrLf
            sOut &= "   ,@ObjectName = '" & ConfigName & "'" & vbCrLf

            If SuccessName <> "" Then
                sOut &= "   ,@SuccessProcess = '" & SuccessName & "'" & vbCrLf
            End If
            If sMode = "D" Then
                sOut &= "   ,@UpdateParent = 'Y'" & vbCrLf
            End If
            If ConfirmMsg <> "" Then
                sOut &= "   ,@ConfirmMsg = '" & ConfirmMsg & "'" & vbCrLf
            End If
        End If

        For Each s In Keys
            If s <> "" Then
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
                If b Then
                    sOut &= "---------------------------------------------------" & vbCrLf
                    sOut &= vbCrLf
                    b = False
                End If
                pc = CType(Values.Item(s), ProcedureColumn)
                sOut &= "execute dbo.shlParametersInsert" & vbCrLf
                sOut &= "    @ObjectName = '" & ConfigName & "'" & vbCrLf
                sOut &= "   ,@ParameterName = '" & pc.ConfigName & "'" & vbCrLf
                sOut &= "   ,@ValueType = '" & pc.vbType & "'" & vbCrLf
                If pc.vbType = "string" Then
                    sOut &= "   ,@Width = " & pc.Length & vbCrLf
                End If
                If Not pc.IsInput Then
                    sOut &= "   ,@IsInput = 'N'" & vbCrLf
                End If
                If Not pc.IsOutput Then
                    sOut &= "   ,@IsOutput = 'N'" & vbCrLf
                End If
            End If
        Next

        If SeekKey <> "" Then
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "---------------------------------------------------" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & ConfigName & "'" & vbCrLf
            sOut &= "   ,@PropertyType = 'sk'" & vbCrLf
            sOut &= "   ,@PropertyName = '" & UCase(SeekKey) & "'" & vbCrLf
            sOut &= "   ,@Value = ''" & vbCrLf
        End If

        Return sOut
    End Function

    Public Function FullText() As String
        Dim sOut As String

        sOut = "if object_id('dbo." & ProcedureName & "') is not null" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    drop procedure dbo." & ProcedureName & vbCrLf
        sOut &= "end" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= ProcedureText

        Return sOut
    End Function

#Region "private functions"
    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Console.WriteLine(e.Message)
    End Sub

    Private Function GetString(ByVal objValue As Object) As String
        If IsDBNull(objValue) Then
            Return ""
        ElseIf objValue Is Nothing Then
            Return ""
        Else
            Try
                Return CType(objValue, String).TrimEnd
            Catch ex As Exception
                Return objValue.ToString
            End Try
        End If
    End Function
#End Region
End Class
