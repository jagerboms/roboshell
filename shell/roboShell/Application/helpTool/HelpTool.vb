Option Explicit On 
Option Strict On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration

Public Class HelpTool
    Private DS As New DataSet
    Private DSH As New DataSet
    Private Copyright As String

    Public Sub UpdateHelpTables(ByVal System As String)
        Dim dt As DataTable
        Dim s As String
        Dim sObject As String

        GetSchema(System)

        ' Objects
        WriteObject(System, "MainMenu")
        dt = DS.Tables(3).Copy
        dt.DefaultView.RowFilter() = "PropertyType='df' and PropertyName='helppage'"
        For Each dr As DataRowView In dt.DefaultView
            sObject = GetString(dr.Item("ObjectName"))
            WriteObject(System, sObject)
        Next

        ' Fields
        dt = DS.Tables(7).Copy
        For Each dr As DataRowView In dt.DefaultView
            Select Case GetString(dr.Item("DisplayType"))
                Case "H", "R"
                Case Else
                    s = GetString(dr.Item("FieldName"))
                    If GetString(dr.Item("Locate")) <> "P" Then
                        sObject = GetString(dr.Item("ObjectName"))
                        WriteField(System, sObject, s)
                    End If
            End Select
        Next

        ' Actions
        dt = DS.Tables(5).Copy
        For Each dr As DataRowView In dt.DefaultView
            s = GetString(dr.Item("ImageFile"))
            If s <> "" Then
                sObject = GetString(dr.Item("ObjectName"))
                s = GetString(dr.Item("ActionName"))
                WriteAction(System, sObject, s)
            End If
        Next

        ' Colours
        WriteObject(System, "MainMenu")
        dt = DS.Tables(3).Copy
        dt.DefaultView.RowFilter() = "PropertyType='cb' or PropertyType='cl'"
        For Each dr As DataRowView In dt.DefaultView
            sObject = GetString(dr.Item("ObjectName"))
            s = GetString(dr.Item("PropertyName"))
            WriteColour(System, sObject, s)
            WriteColour(System, sObject, "")
        Next
    End Sub

    Public Sub BuildHelpPages(ByVal System As String, ByVal sPath As String)
        Dim dt As DataTable
        Dim s As String
        Dim sPage As String

        GetSchema(System)
        GetHelpSchema(System)
        Copyright = GetCopyright()

        s = doLicence()
        PutFile(sPath & "\Licence.html", s)

        s = doMainMenu()
        PutFile(sPath & "\MainMenu.html", s)

        dt = DS.Tables(3).Copy
        dt.DefaultView.RowFilter() = "PropertyType='df' and PropertyName='helppage'"
        For Each dr As DataRowView In dt.DefaultView
            s = GetString(dr.Item("ObjectName"))
            sPage = GetString(dr.Item("Value"))
            s = doPage(s)
            PutFile(sPath & "\" & sPage, s)
        Next
        dt.DefaultView.RowFilter() = ""

    End Sub

    Private Sub GetSchema(ByVal System As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString(System))
            psConn.Open()
            psAdapt = New SqlDataAdapter("shlShellGet", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            psAdapt.Fill(DS)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GetHelpSchema(ByVal System As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString("roboshell"))
            psConn.Open()
            psAdapt = New SqlDataAdapter("helpHelpGet", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
            psAdapt.SelectCommand.Parameters("@SystemID").Value = System
            psAdapt.Fill(DSH)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteObject(ByVal sSystem As String, ByVal sObject As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString("roboshell"))
            psConn.Open()
            psAdapt = New SqlDataAdapter("helpObjectsInsert", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
            psAdapt.SelectCommand.Parameters("@SystemID").Value = sSystem
            psAdapt.SelectCommand.Parameters("@ObjectName").Value = sObject
            psAdapt.Fill(DS)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteField(ByVal sSystem As String, ByVal sObject As String, ByVal sField As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString("roboshell"))
            psConn.Open()
            psAdapt = New SqlDataAdapter("helpFieldsInsert", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
            psAdapt.SelectCommand.Parameters("@SystemID").Value = sSystem
            psAdapt.SelectCommand.Parameters("@ObjectName").Value = sObject
            psAdapt.SelectCommand.Parameters("@FieldName").Value = sField
            psAdapt.Fill(DS)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteAction(ByVal sSystem As String, ByVal sObject As String, ByVal sAction As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString("roboshell"))
            psConn.Open()
            psAdapt = New SqlDataAdapter("helpActionsInsert", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
            psAdapt.SelectCommand.Parameters("@SystemID").Value = sSystem
            psAdapt.SelectCommand.Parameters("@ObjectName").Value = sObject
            psAdapt.SelectCommand.Parameters("@ActionName").Value = sAction
            psAdapt.Fill(DS)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteColour(ByVal sSystem As String, ByVal sObject As String, ByVal sColourValue As String)
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter

        Try
            psConn = New SqlConnection(GetConnectString("roboshell"))
            psConn.Open()
            psAdapt = New SqlDataAdapter("helpColoursInsert", psConn)
            psAdapt.SelectCommand.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(psAdapt.SelectCommand)
            psAdapt.SelectCommand.Parameters("@SystemID").Value = sSystem
            psAdapt.SelectCommand.Parameters("@ObjectName").Value = sObject
            psAdapt.SelectCommand.Parameters("@ColourValue").Value = sColourValue
            psAdapt.Fill(DS)
            psConn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function doPage(ByVal sObject As String) As String
        Dim s As String
        Dim ss As String
        Dim sFld As String
        Dim sOut As String
        Dim dt As DataTable
        Dim Shell As String

        Shell = GetShell(sObject)
        sOut = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>" & vbCrLf
        sOut &= "<html>" & vbCrLf
        sOut &= "  <head>" & vbCrLf
        s = GetObjectProperty(sObject, "df", "title")
        If s = "" Then s = sObject
        sOut &= "    <title>" & s & "</title>" & vbCrLf
        sOut &= "    <link rel='shortcut icon' href='favicon.ico' type='image/x-icon' />" & vbCrLf
        sOut &= "    <link rel ='stylesheet' href='help.css' type='text/css' />" & vbCrLf
        If Shell = "Y" Then
            sOut &= "    <link rel ='stylesheet' href='shell.css' type='text/css' />" & vbCrLf
        End If
        sOut &= "  </head>" & vbCrLf
        sOut &= "  <body>" & vbCrLf
        sOut &= "<div id='maincontainer'>" & vbCrLf
        sOut &= "  <div id='topsection'>" & vbCrLf

        ss = GetVariable("SystemName")
        sOut &= "    <div class='innertube'><h1>" & ss
        ss = GetVariable("Release")
        If ss <> "" Then sOut &= " (Release " & ss & ")"
        sOut &= "</h1></div>" & vbCrLf

        sOut &= "  </div>" & vbCrLf
        sOut &= "  <div id='contentwrapper'>" & vbCrLf
        sOut &= "    <div id='contentcolumn'>" & vbCrLf
        sOut &= "      <div class='innertube'>" & vbCrLf
        sOut &= "        <h2>" & s & "</h2>"
        sOut &= "        <b>Description</b><br />" & vbCrLf
        s = GetDescription(sObject)
        sOut &= s & "<br />" & vbCrLf
        sOut &= "        <b>Columns</b><br />" & vbCrLf

        dt = DS.Tables(7).Copy
        dt.DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
        For Each dr As DataRowView In dt.DefaultView
            Select Case GetString(dr.Item("DisplayType"))
                Case "H", "R"
                Case Else
                    sFld = GetString(dr.Item("FieldName"))
                    If GetString(dr.Item("Locate")) <> "P" Then
                        s = GetString(dr.Item("Label"))
                        If s = "" Then s = sFld
                        sOut &= "        <em>" & s & "</em> - "
                        s = GetFieldDescription(sObject, sFld)
                        If s = "" Then s = "..."
                        sOut &= s & "<br />" & vbCrLf
                    Else
                        s = s
                    End If
            End Select
        Next
        sOut &= getColours(sObject) & vbCrLf
        sOut &= "        <b>Actions</b><br />" & vbCrLf
        dt = DS.Tables(5).Copy
        dt.DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
        For Each dr As DataRowView In dt.DefaultView
            s = GetString(dr.Item("ImageFile"))
            If s <> "" Then
                sFld = GetString(dr.Item("ActionName"))
                sOut &= "        <img alt='" & sFld & " icon' src='../image/" & s & "' border='0' /> "
                s = GetString(dr.Item("Process"))
                s = GetProcessObject(s)
                ss = GetObjectProperty(s, "df", "helppage")
                If ss <> "" Then
                    sOut &= "<a href='" & ss & "'>"
                    s = GetString(dr.Item("ToolTip"))
                    sOut &= s & "</a> - "
                Else
                    s = GetString(dr.Item("ToolTip"))
                    sOut &= "<em>" & s & "</em> - "
                End If
                s = GetActionDescription(sObject, sFld)
                If s = "" Then s = "..."
                sOut &= s & "<br />" & vbCrLf
            End If
        Next
        sOut &= "      </div>" & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= vbCrLf
        sOut &= "  <div id='leftcolumn'>" & vbCrLf
        sOut &= "    <div class='innertube'>" & vbCrLf
        sOut &= "        <b>Contents</b><br />" & vbCrLf
        sOut &= getTOC() & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        If Shell = "Y" Then
            sOut &= "  <div id='footer'>&copy; Russell Hansen, Tolbeam Pty Limited. </div>" & vbCrLf
        Else
            sOut &= "  <div id='footer'>&copy; " & Copyright & " </div>" & vbCrLf
        End If
        sOut &= "</div>" & vbCrLf
        sOut &= "  </body>" & vbCrLf
        sOut &= "</html>" & vbCrLf

        Return sOut
    End Function

    Private Function doMainMenu() As String
        Dim s As String
        Dim ss As String
        Dim sFld As String
        Dim sOut As String
        Dim dt As DataTable

        ss = GetVariable("SystemName")
        sOut = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>" & vbCrLf
        sOut &= "<html>" & vbCrLf
        sOut &= "  <head>" & vbCrLf
        sOut &= "    <title>" & ss & " Main Menu</title>" & vbCrLf
        sOut &= "    <link rel='shortcut icon' href='favicon.ico' type='image/x-icon' />" & vbCrLf
        sOut &= "    <link rel ='stylesheet' href='help.css' type='text/css' />" & vbCrLf
        sOut &= "  </head>" & vbCrLf
        sOut &= "  <body>" & vbCrLf
        sOut &= "<div id='maincontainer'>" & vbCrLf
        sOut &= "  <div id='topsection'>" & vbCrLf
        sOut &= "    <div class='innertube'><h1>" & ss
        ss = GetVariable("Release")
        If ss <> "" Then sOut &= " (Release " & ss & ")"
        sOut &= "</h1></div>" & vbCrLf

        sOut &= "  </div>" & vbCrLf
        sOut &= "  <div id='contentwrapper'>" & vbCrLf
        sOut &= "    <div id='contentcolumn'>" & vbCrLf
        sOut &= "      <div class='innertube'>" & vbCrLf
        sOut &= "        <h2>Main Menu</h2>"
        sOut &= "        <b>Description</b><br />" & vbCrLf
        s = GetDescription("mainmenu")
        sOut &= s & "<br />" & vbCrLf

        sOut &= "        <b>Actions</b><br />" & vbCrLf
        dt = DS.Tables(5).Copy
        dt.DefaultView.RowFilter() = "ObjectName='mainmenu'"
        For Each dr As DataRowView In dt.DefaultView
            s = GetString(dr.Item("ImageFile"))
            If s <> "" Then
                sFld = GetString(dr.Item("ActionName"))
                sOut &= "        <img alt='" & sFld & " icon' src='../image/" & s & "' border='0' /> "
                s = GetString(dr.Item("Process"))
                s = GetProcessObject(s)
                ss = GetObjectProperty(s, "df", "helppage")
                If ss <> "" Then
                    sOut &= "<a href='" & ss & "'>"
                    s = GetString(dr.Item("ToolTip"))
                    sOut &= s & "</a> - "
                Else
                    s = GetString(dr.Item("ToolTip"))
                    sOut &= "<em>" & s & "</em> - "
                End If
                s = GetActionDescription("mainmenu", sFld)
                If s = "" Then s = "..."
                sOut &= s & "<br />" & vbCrLf
            End If
        Next
        '        <img src="../image/security.gif" /> <a href="Users.html">Users</a> - User permission maintenance.<br />
        sOut &= "      </div>" & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= vbCrLf
        sOut &= "  <div id='leftcolumn'>" & vbCrLf
        sOut &= "    <div class='innertube'>" & vbCrLf
        sOut &= "        <b>Contents</b><br />" & vbCrLf
        sOut &= getTOC() & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= "  <div id='footer'>&copy; " & Copyright & " </div>" & vbCrLf
        sOut &= "</div>" & vbCrLf
        sOut &= "  </body>" & vbCrLf
        sOut &= "</html>" & vbCrLf

        Return sOut
    End Function


    Private Function doLicence() As String
        Dim sOut As String

        sOut = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>" & vbCrLf
        sOut &= "<html>" & vbCrLf
        sOut &= "  <head>" & vbCrLf
        sOut &= "    <title>Roboshell Licence</title>" & vbCrLf
        sOut &= "    <link rel='shortcut icon' href='favicon.ico' type='image/x-icon' />" & vbCrLf
        sOut &= "    <link rel ='stylesheet' href='help.css' type='text/css' />" & vbCrLf
        sOut &= "    <link rel ='stylesheet' href='shell.css' type='text/css' />" & vbCrLf
        sOut &= "  </head>" & vbCrLf
        sOut &= "  <body>" & vbCrLf
        sOut &= "<div id='maincontainer'>" & vbCrLf
        sOut &= "  <div id='topsection'>" & vbCrLf
        sOut &= "    <div class='innertube'><h1>Roboshell Licence</h1></div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= "  <div id='contentwrapper'>" & vbCrLf
        sOut &= "    <div id='contentcolumn'>" & vbCrLf
        sOut &= "      <div class='innertube'>" & vbCrLf
        sOut &= "        <b>Description</b><br />" & vbCrLf
        sOut &= "Roboshell and the Roboshell tools is free software issued as open source; you can redistribute it and/or modify it under the terms of the GNU General Public License version 2 as published by the Free Software Foundation.<br /><br />" & vbCrLf
        sOut &= "Roboshell and the Roboshell tools is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; " & vbCrLf
        sOut &= "without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. " & vbCrLf
        sOut &= "See the GNU General Public License for more details.<br /><br />" & vbCrLf
        sOut &= "You should have received a copy of the GNU General Public License along with Roboshell application suite; if not, go to the <a href='http://www.gnu.org/licenses/gpl-2.0.html'>web site</a> or write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307 USA." & vbCrLf
        sOut &= "      </div>" & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= vbCrLf
        sOut &= "  <div id='leftcolumn'>" & vbCrLf
        sOut &= "    <div class='innertube'>" & vbCrLf
        sOut &= "        <b>Contents</b><br />" & vbCrLf
        sOut &= getTOC() & vbCrLf
        sOut &= "    </div>" & vbCrLf
        sOut &= "  </div>" & vbCrLf
        sOut &= "  <div id='footer'>&copy; Russell Hansen, Tolbeam Pty Limited. </div>" & vbCrLf
        sOut &= "</div>" & vbCrLf
        sOut &= "  </body>" & vbCrLf
        sOut &= "</html>" & vbCrLf

        Return sOut
    End Function

    Private Function getColours(ByVal sObject As String) As String
        Dim s As String
        Dim ss As String
        Dim scb As String
        Dim sOut As String = ""
        Dim b As Boolean = True

        s = GetObjectProperty(sObject, "df", "ColourColumn")
        If s <> "" Then
            sOut &= vbCrLf & "        <b>Colours</b><br />" & vbCrLf
            ss = GetColour(sObject)
            If ss <> "" Then
                sOut &= "        " & ss & vbCrLf
            Else
                sOut &= "        The colours used are determined by the " & s & " column as follows:" & vbCrLf
            End If
            sOut &= "        <table border='1' cellpadding='3' cellspacing='0'>" & vbCrLf
            sOut &= "         <tr style=""border:solid"">" & vbCrLf

            DSH.Tables(3).DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
            For Each dr As DataRowView In DSH.Tables(3).DefaultView
                s = GetString(dr.Item("ColourValue"))
                If s <> "" Then
                    ss = GetObjectProperty(sObject, "cl", s)
                    scb = GetObjectProperty(sObject, "cb", s)
                Else
                    ss = "blue"
                    scb = ""
                End If
                s = GetString(dr.Item("ValueDescription"))
                If s <> "" Then
                    sOut &= "           <td valign='top' style=""color:" & ss & ";"
                    If scb <> "" Then
                        sOut &= " background-color:" & scb & ";" & vbCrLf
                    End If
                    sOut &= " border-color:#000000"">" & s & "</td>" & vbCrLf
                    b = False
                End If
            Next
            DSH.Tables(3).DefaultView.RowFilter() = ""
            sOut &= "         </tr>" & vbCrLf
            sOut &= "       </table>" & vbCrLf
        End If
        If b Then sOut = ""
        Return sOut
    End Function

    Private Function getTOC() As String
        Dim sMenu As String
        Dim dt As DataTable
        Dim s As String
        Dim ss As String

        sMenu = "        <img alt='system icon' style=""height:16; width:16;"" src='favicon.ico' /> <a href='MainMenu.html'>Main Menu</a><br />"
        dt = DS.Tables(5).Copy
        dt.DefaultView.RowFilter() = "ObjectName='mainmenu'"
        For Each dr As DataRowView In dt.DefaultView
            Select Case GetString(dr.Item("MenuType"))
                Case "S"
                    s = GetString(dr.Item("ActionName"))
                    sMenu &= vbCrLf & "        "
                    ss = GetString(dr.Item("ImageFile"))
                    If ss <> "" Then
                        sMenu &= "<img alt='" & s & " icon' src='../image/" & ss & "' /> "
                    End If
                    sMenu &= s & "<br />"
                Case Else
                    s = GetString(dr.Item("Process"))
                    s = GetProcessObject(s)
                    sMenu &= vbCrLf & "        "
                    ss = GetString(dr.Item("ImageFile"))
                    If ss <> "" Then
                        sMenu &= "<img alt='" & s & " icon' src='../image/" & ss & "' />"
                    Else
                        sMenu &= "<img alt='dots' src='../image/dot.gif' />"
                    End If
                    ss = GetObjectProperty(s, "df", "helppage")
                    sMenu &= " <a href='" & ss & "'>"
                    ss = GetObjectProperty(s, "df", "title")
                    sMenu &= ss & "</a><br />"
                    sMenu &= GetObjectTOC(s, 0)
            End Select
        Next
        sMenu &= vbCrLf & "        <a href='Licence.html'>Licence</a><br />"

        Return sMenu
    End Function

    Private Function GetObjectTOC(ByVal sObject As String, ByVal Level As Integer) As String
        Dim sMenu As String = ""
        Dim dt As DataTable
        Dim sName As String
        Dim s As String
        Dim ss As String
        Dim i As Integer

        dt = DS.Tables(5).Copy
        dt.DefaultView.RowFilter() = "ObjectName='" & sObject & "'"

        For Each dr As DataRowView In dt.DefaultView
            If GetString(dr.Item("ImageFile")) <> "" Then
                sName = GetString(dr.Item("ActionName"))
                Select Case GetString(dr.Item("MenuType"))
                    Case "S"
                    Case Else
                        s = GetString(dr.Item("Process"))
                        If s <> "" Then
                            s = GetProcessObject(s)
                            ss = GetObjectProperty(s, "df", "helppage")
                            If ss <> "" Then
                                sMenu &= vbCrLf & "        "
                                For i = 1 To Level
                                    sMenu &= "<img alt='dots' src='../image/dot.gif' />"
                                Next
                                sMenu &= "<img alt='dots' src='../image/dot.gif' /> "
                                sMenu &= "<a href='" & ss & "'>"
                                ss = GetObjectProperty(s, "df", "title")
                                sMenu &= ss & "</a><br />"
                                sMenu &= GetObjectTOC(s, Level + 1)
                            End If
                        End If

                        dt = DS.Tables(10).Copy
                        dt.DefaultView.RowFilter() = "ObjectName='" & sObject & "' and ActionName='" & sName & "'"
                        For Each drv As DataRowView In dt.DefaultView
                            s = GetString(drv.Item("Process"))
                            If s <> "" Then
                                s = GetProcessObject(s)
                                ss = GetObjectProperty(s, "df", "helppage")
                                If ss <> "" Then
                                    sMenu &= vbCrLf & "        "
                                    For i = 1 To Level
                                        sMenu &= "<img alt='dots' src='../image/dot.gif' />"
                                    Next
                                    sMenu &= "<img alt='dots' src='../image/dot.gif' /> "
                                    sMenu &= "<a href='" & ss & "'>"
                                    ss = GetObjectProperty(s, "df", "title")
                                    sMenu &= ss & "</a><br />"
                                    sMenu &= GetObjectTOC(s, Level + 1)
                                End If
                            End If
                        Next
                End Select
            End If
        Next
        Return sMenu
    End Function

    Private Function GetProcessObject(ByVal sProcess As String) As String
        Dim s As String = ""
        DS.Tables(1).DefaultView.RowFilter() = "ProcessName='" & sProcess & "'"
        For Each dr As DataRowView In DS.Tables(1).DefaultView
            s = GetString(dr.Item("ObjectName"))
            If GetObjectProperty(s, "df", "helppage") = "" Then
                s = GetString(dr.Item("SuccessProcess"))
                If s <> "" Then
                    s = GetProcessObject(s)
                End If
            End If
        Next
        DS.Tables(1).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetVariable(ByVal sProperty As String) As String
        Dim s As String = ""

        DS.Tables(0).DefaultView.RowFilter() = "VariableID='" & sProperty & "'"
        For Each dr As DataRowView In DS.Tables(0).DefaultView
            s = GetString(dr.Item("VariableValue"))
            Exit For
        Next
        DS.Tables(0).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetDescription(ByVal sObject As String) As String
        Dim s As String = ""

        DSH.Tables(0).DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
        For Each dr As DataRowView In DSH.Tables(0).DefaultView
            s = GetString(dr.Item("HelpText"))
            Exit For
        Next
        DSH.Tables(0).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetShell(ByVal sObject As String) As String
        Dim s As String = ""

        DSH.Tables(0).DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
        For Each dr As DataRowView In DSH.Tables(0).DefaultView
            s = GetString(dr.Item("Shell"))
            Exit For
        Next
        DSH.Tables(0).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetColour(ByVal sObject As String) As String
        Dim s As String = ""

        DSH.Tables(0).DefaultView.RowFilter() = "ObjectName='" & sObject & "'"
        For Each dr As DataRowView In DSH.Tables(0).DefaultView
            s = GetString(dr.Item("ColourText"))
            Exit For
        Next
        DSH.Tables(0).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetCopyright() As String
        Dim s As String = ""

        For Each dr As DataRowView In DSH.Tables(4).DefaultView
            s = GetString(dr.Item("Copyright"))
            Exit For
        Next

        Return s
    End Function

    Private Function GetFieldDescription(ByVal sObject As String, ByVal sField As String) As String
        Dim s As String = ""

        DSH.Tables(1).DefaultView.RowFilter() = "ObjectName='" & sObject & "' and FieldName='" & sField & "'"
        For Each dr As DataRowView In DSH.Tables(1).DefaultView
            s = GetString(dr.Item("HelpText"))
            Exit For
        Next
        DSH.Tables(1).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetActionDescription(ByVal sObject As String, ByVal sAction As String) As String
        Dim s As String = ""

        DSH.Tables(2).DefaultView.RowFilter() = "ObjectName='" & sObject & "' and ActionName='" & sAction & "'"
        For Each dr As DataRowView In DSH.Tables(2).DefaultView
            s = GetString(dr.Item("HelpText"))
            Exit For
        Next
        DSH.Tables(2).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetObjectProperty(ByVal sObject As String, ByVal sType As String, ByVal sProperty As String) As String
        Dim s As String = ""

        DS.Tables(3).DefaultView.RowFilter() = "ObjectName='" & sObject & "' and PropertyType='" & sType & "' and PropertyName='" & sProperty & "'"
        For Each dr As DataRowView In DS.Tables(3).DefaultView
            s = GetString(dr.Item("Value"))
            Exit For
        Next
        DS.Tables(3).DefaultView.RowFilter() = ""

        Return s
    End Function

    Private Function GetCommandParameter(ByRef sSwitch As String, _
                                Optional ByRef sDefault As String = "") As String
        Dim sCommand As String
        Dim sParameter As String
        Dim i As Integer

        sCommand = Microsoft.VisualBasic.Command()
        i = InStr(1, sCommand, sSwitch, CompareMethod.Text)
        sParameter = ""
        If i > 0 Then
            sParameter = Mid(sCommand, i + 2)
            i = InStr(1, sParameter, " ", CompareMethod.Text)
            If i > 0 Then
                sParameter = Mid(sParameter, 1, i - 1)
            End If
        Else
            sParameter = sDefault
        End If
        GetCommandParameter = sParameter
    End Function

    Private Function GetConnectString(ByVal sName As String) As String
        Dim settings As System.Configuration.ConnectionStringSettingsCollection = _
            ConfigurationManager.ConnectionStrings

        If Not settings Is Nothing Then
            Return settings.Item(sName).ConnectionString
        End If
        Return ""
    End Function

    Private Function GetString(ByVal objValue As Object) As String
        Dim s As String

        If IsDBNull(objValue) Then
            Return ""
        ElseIf objValue Is Nothing Then
            Return ""
        Else
            Try
                If objValue.GetType().ToString = "System.DateTime" Then
                    If CDate(objValue) = Date.MinValue Then
                        s = ""
                    Else
                        s = Format(objValue, "dd-MMM-yyyy hh:mm:ss tt")
                        If Mid(s, 13, 11) = "12:00:00 AM" Then
                            s = Mid(s, 1, 11)
                        End If
                    End If
                    Return s
                Else
                    Return CType(objValue, String).TrimEnd
                End If

            Catch ex As Exception
                Return objValue.ToString
            End Try
        End If
    End Function

    Private Function CheckFile(ByVal sName As String) As Boolean
        Dim b As Boolean = False
        If Dir(sName) <> "" Then b = True
        Return b
    End Function

    Private Function PutFile(ByVal sName As String, ByVal sContent As String) As Boolean
        If CheckFile(sName) Then
            Return False
        Else
            Dim file As New System.IO.StreamWriter(sName)
            file.Write(sContent)
            file.Close()
            Return True
        End If
    End Function
End Class
