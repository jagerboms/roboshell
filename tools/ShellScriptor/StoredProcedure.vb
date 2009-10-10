Option Explicit On
Option Strict On

Imports System.Data.SqlClient

#Region "copyright Russell Hansen, Tolbeam Pty Limited"
'dbSchema is free software issued as open source;
' you can redistribute it and/or modify it under the terms of the
' GNU General Public License version 2 as published by the Free Software Foundation.
'dbSchema is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with dbSchema;
' if not, go to the web site (http://www.gnu.org/licenses/gpl-2.0.html)
' or write to:
'   The Free Software Foundation, Inc.,
'   59 Temple Place,
'   Suite 330,
'   Boston, MA 02111-1307 USA. 
#End Region

Public Class ProcedureColumn
    Private sType As String = ""
    Private iLength As Integer = 0
    Private iPrecision As Integer = 0
    Private iScale As Integer = 0
    Private sqllib As New sql

#Region "Properties"
    Public Index As Integer
    Public Name As String
    Public TestValue As Object
    Public IsInput As Boolean
    Public IsOutput As Boolean

    Public ReadOnly Property ConfigName() As String
        Get
            ConfigName = Mid(Name, 2)
        End Get
    End Property

    Public ReadOnly Property QuotedName() As String
        Get
            QuotedName = sqllib.QuoteIdentifier(Name)
        End Get
    End Property

    Public Property Type() As String
        Get
            Type = sType
        End Get
        Set(ByVal value As String)
            If value = "int" Then
                sType = "integer"
            Else
                sType = value
            End If
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "decimal", "numeric", "datetime", "smalldatetime"
                    i = 0
                Case "sysname"
                    i = 128
                Case Else
                    i = iLength
            End Select
            Length = i
        End Get
        Set(ByVal value As Integer)
            iLength = value
        End Set
    End Property

    Public Property Precision() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "char", "varchar", "nvarchar", "datetime", "smalldatetime", "sysname"
                    i = 0
                Case Else
                    i = iPrecision
            End Select
            Precision = i
        End Get
        Set(ByVal value As Integer)
            iPrecision = value
        End Set
    End Property

    Public Property Scale() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "char", "varchar", "nvarchar", "datetime", "smalldatetime", "sysname"
                    i = 0
                Case Else
                    i = iScale
            End Select
            Scale = i
        End Get
        Set(ByVal value As Integer)
            iScale = value
        End Set
    End Property

    Public ReadOnly Property vbType() As String
        Get
            Dim s As String
            Select Case sType
                Case "char", "varchar", "nvarchar", "sysname"
                    s = "string"
                Case "decimal", "numeric"
                    s = "double"
                Case "datetime"
                    s = "datetime"
                Case "smalldatetime"
                    s = "date"
                Case "uniqueidentifier"
                    s = "Guid"
                Case Else
                    s = sType
            End Select
            vbType = s
        End Get
    End Property

    Public ReadOnly Property TypeText() As String
        Get
            Dim s As String
            s = sType
            Select Case sType
                Case "char", "varchar", "nvarchar"
                    s &= "(" & iLength & ")"
                Case "decimal", "numeric"
                    s &= "(" & iPrecision & "," & iScale & ")"
            End Select
            TypeText = s
        End Get
    End Property
#End Region
End Class

Public Class StoredProcedure
    Private pName As String = ""
    Private qName As String = ""
    Private sMode As String = "D"
    Private PreLoad As Integer = -1
    Private Values As New Hashtable
    Private Keys(0) As String
    Private Results As New Hashtable
    Private ResultKeys(0) As String
    Private slib As New sql

    Public ProcedureText As String
    Public ConfigName As String
    Public ModuleName As String
    Public ProcessName As String
    Public SuccessName As String
    Public Messages As Boolean = True
    Public ConfirmMsg As String = ""
    Public SeekKey As String = ""

    Public Property ProcedureName() As String
        Get
            ProcedureName = pName
        End Get
        Set(ByVal value As String)
            If PreLoad = 0 Then
                pName = value
                qName = slib.QuoteIdentifier(value)
            End If
        End Set
    End Property

    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal Name As String, ByVal sqllib As sql)
        Dim s As String = "a"
        Dim sText As String
        Dim b As Boolean = False
        Dim sName As String
        Dim sType As String
        Dim dt As DataTable
        Dim dr As DataRow

        PreLoad = 2
        dt = sqllib.ProcParms(Name)
        If dt.Rows.Count = 0 Then
            pName = Name
            qName = slib.QuoteIdentifier(pName)
            PreLoad = 3
            Return
        End If

        For Each dr In dt.Rows        ' Columns
            If pName = "" Then
                pName = GetString(dr.Item("SPECIFIC_NAME"))
                qName = slib.QuoteIdentifier(pName)
            End If
            sName = GetString(dr.Item("PARAMETER_NAME"))
            sType = GetString(dr.Item("DATA_TYPE"))
            AddParameter(sName, sType, dr.Item("CHARACTER_MAXIMUM_LENGTH"), _
                dr.Item("NUMERIC_PRECISION"), dr.Item("NUMERIC_SCALE"), _
                    LCase(GetString(dr.Item("PARAMETER_MODE"))))
        Next

        dt = Nothing
        dt = sqllib.ObjectText(Name)
        sText = ""
        For Each dr In dt.Rows        ' Columns
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

        With parm
            .Index = Values.Count
            .Name = sName
            .Type = sType
            If IsNumeric(oLength) Then
                .Length = CInt(oLength)
            End If
            If IsNumeric(oPrecision) Then
                .Precision = CInt(oPrecision)
            End If
            If IsNumeric(oScale) Then
                .Scale = CInt(oScale)
            End If
            .IsInput = InStr(InOut, "in") > 0
            .IsOutput = InStr(InOut, "out") > 0
        End With

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

        With parm
            .Index = Values.Count
            .Name = sName
            .Type = sType
            If IsNumeric(oLength) Then
                .Length = CInt(oLength)
            End If
            If IsNumeric(oPrecision) Then
                .Precision = CInt(oPrecision)
            End If
            If IsNumeric(oScale) Then
                .Scale = CInt(oScale)
            End If
            .IsInput = False
            .IsOutput = True
        End With

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
        sOut &= "   ,@procname = '" & pName & "'" & vbCrLf
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

        sOut = "if object_id('dbo." & pName & "') is not null" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    drop procedure dbo." & pName & vbCrLf
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
