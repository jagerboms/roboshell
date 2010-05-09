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

    Public Sub New(ByVal sTableName As String, ByVal Schema As String, _
                   ByVal sqllib As sql)
        slib = sqllib
        cTable = New TableDefn(sTableName, Schema, sqllib)
    End Sub

    Public Function DataScript(ByVal sFilter As String) As String
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

        If cTable.IdentityColumn <> "" Then
            sOut &= "set identity_insert " & cTable.QuotedSchema & "." & cTable.QuotedName & " on" & vbCrLf
            sOut &= vbCrLf
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
        sHead &= "(" & vbCrLf

        i = 0
        dt = slib.TableData(cTable.QuotedName, cTable.QuotedSchema, sFilter)
        For Each r As DataRow In dt.Rows
            If i = 0 Then
                sOut &= sTail
                sOut &= sHead
                s = "    select  "
                For Each tc In cTable.AllColumns
                    sOut &= s & tc.DataFormat(r(tc.Name)) & " " & tc.QuotedName & vbCrLf
                    s = "           ,"
                Next
                sTail = ") x" & vbCrLf
                sTail &= "left join " & cTable.QuotedSchema & "." & cTable.QuotedName & " a" & vbCrLf

                s = cTable.IKeys.PrimaryKey
                If s <> "" Then
                    cNDX = cTable.IKeys.Item(s)
                End If

                If cNDX Is Nothing Then
                    Return ""
                End If

                s = "on      a."
                For Each ic As IndexColumn In cNDX.Columns
                    qN = slib.QuoteIdentifier(ic.Name)
                    If ss = "" Then ss = qN
                    sTail &= s & qN & " = x." & qN & vbCrLf
                    s = "and     a."
                Next

                sTail &= "where   a." & ss & " is null" & vbCrLf
                sTail &= "go" & vbCrLf & vbCrLf

                i += 1
            Else
                s = "    union select "
                For Each tc In cTable.AllColumns
                    sOut &= s & tc.DataFormat(r(tc.Name))
                    s = ", "
                Next
                sOut &= vbCrLf
                i += 1
            End If
            If i = 100 Then i = 0
        Next

        sOut &= sTail
        If cTable.IdentityColumn <> "" Then
            sOut &= vbCrLf
            sOut &= "set identity_insert " & cTable.QuotedSchema & "." & cTable.QuotedName & " off" & vbCrLf
            sOut &= "go" & vbCrLf & vbCrLf
        End If

        Return sOut
    End Function
End Class
