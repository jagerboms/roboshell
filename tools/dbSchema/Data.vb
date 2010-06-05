Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections

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

Public Class Data
    Private slib As New sql
    Private cTable As TableDefn
    Private sFilter As String
    Private sFileName As String

    Public Property Filter() As String
        Get
            Filter = sFilter
        End Get
        Set(ByVal sf As String)
            sFilter = sf
        End Set
    End Property

    Public Property FileName() As String
        Get
            FileName = sFileName
        End Get
        Set(ByVal fn As String)
            sFileName = fn
        End Set
    End Property

    Public Sub New(ByVal sTableName As String, ByVal Schema As String, _
                   ByVal sqllib As sql)
        slib = sqllib
        cTable = New TableDefn(sTableName, Schema, sqllib)
    End Sub

    Public Sub DataScript()
        Dim sOut As String = ""
        Dim tc As TableColumn
        Dim dt As DataTable
        Dim sHead As String
        Dim sTail As String = ""
        Dim s As String
        Dim qN As String
        Dim i As Integer
        Dim ss As String = ""
        Dim cNDX As TableIndex = Nothing
        Dim file As New System.IO.StreamWriter(sFileName)

        If cTable.IdentityColumn <> "" Then
            file.WriteLine("set identity_insert " & cTable.QuotedSchema & "." & cTable.QuotedName & " on")
            file.WriteLine("")
        End If

        sHead = "insert into " & cTable.QuotedSchema & "." & cTable.QuotedName & vbCrLf
        sHead &= "(" & vbCrLf
        s = "    "
        For Each tc In cTable.AllColumns
            sHead &= s & tc.QuotedName
            s = ", "
        Next
        sHead &= vbCrLf
        sHead &= ")" & vbCrLf
        s = "select  x."
        For Each tc In cTable.AllColumns
            sHead &= s & tc.QuotedName & vbCrLf
            s = "       ,x."
        Next
        sHead &= "from" & vbCrLf
        sHead &= "("

        i = 0
        dt = slib.TableData(cTable.QuotedName, cTable.QuotedSchema, sFilter)
        For Each r As DataRow In dt.Rows
            sOut = ""
            If i = 0 Then
                file.WriteLine(sTail)
                file.WriteLine(sHead)
                s = "    select  "
                For Each tc In cTable.AllColumns
                    file.WriteLine(s & tc.QuotedFormat(r(tc.Name), False) & " " & tc.QuotedName)
                    s = "           ,"
                Next
                sTail = ") x" & vbCrLf
                sTail &= "left join " & cTable.QuotedSchema & "." & cTable.QuotedName & " a" & vbCrLf

                s = cTable.IKeys.PrimaryKey
                If s <> "" Then
                    cNDX = cTable.IKeys.Item(s)
                End If

                If cNDX Is Nothing Then
                    Return
                End If

                s = "on      a."
                For Each ic As IndexColumn In cNDX.Columns
                    qN = slib.QuoteIdentifier(ic.Name)
                    If ss = "" Then ss = qN
                    sTail &= s & qN & " = x." & qN & vbCrLf
                    s = "and     a."
                Next

                sTail &= "where   a." & ss & " is null" & vbCrLf
                sTail &= "go" & vbCrLf

                i += 1
            Else
                s = "    union select "
                For Each tc In cTable.AllColumns
                    sOut &= s & tc.QuotedFormat(r(tc.Name), False)
                    s = ", "
                Next
                file.WriteLine(sOut)
                i += 1
            End If
            If i = 100 Then i = 0
        Next

        file.WriteLine(sTail)
        If cTable.IdentityColumn <> "" Then
            file.WriteLine("")
            file.WriteLine("set identity_insert " & cTable.QuotedSchema & "." & cTable.QuotedName & " off")
            file.WriteLine("go")
            file.WriteLine("")
        End If

        file.Close()
    End Sub

    Public Sub DataCSV()
        Dim tc As TableColumn
        Dim dt As DataTable
        Dim sOut As String = ""
        Dim s As String = ""
        Dim file As New System.IO.StreamWriter(sFileName)

        dt = slib.TableData(cTable.QuotedName, cTable.QuotedSchema, sFilter)

        file.WriteLine(cTable.QuotedSchema & "." & cTable.QuotedName)
        s = ""
        For Each tc In cTable.AllColumns
            sOut &= s & """" & tc.QuotedName & """"
            s = ","
        Next
        file.WriteLine(sOut)

        For Each r As DataRow In dt.Rows
            sOut = ""
            s = ""
            For Each tc In cTable.AllColumns
                sOut &= s & tc.QuotedFormat(r(tc.Name), True)
                s = ","
            Next
            file.WriteLine(sOut)
        Next

        file.Close()
    End Sub
End Class
