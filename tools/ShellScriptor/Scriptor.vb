Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Specialized

Module Scriptor
    Dim MainSCL As String = ""
    Dim SystemName As String = ""
    Dim sqllib As New sql

    ' xml field/parameter overrides
    ' xml data scripts

    Sub main()
        Dim sXML As String
        Dim i As Integer
        Dim sTable As String
        Dim sModule As String
        Dim sObject As String
        Dim sParent As String
        Dim sDescription As String
        Dim sItem As String
        Dim sSeekKey As String
        Dim sKey As String

        Try                                 ' Read the config XML into a DataSet
            sXML = GetCommandParameter("-x")
            If sXML = "" Then
                sTable = GetCommandParameter("-t")
                If sTable = "" Then
                    Console.WriteLine("Usage: SPScriptor -xfile.sdef | -tTableName [-oObjectName] [-mModule] [-pParent] [-dDescription] [-iItemName] [-sSeekKey] [-kConnectKey]")
                Else
                    sObject = GetCommandParameter("-o")
                    sModule = GetCommandParameter("-m")
                    sParent = LCase(GetCommandParameter("-p"))
                    sDescription = GetCommandParameter("-d")
                    sItem = GetCommandParameter("-i")
                    sSeekKey = GetCommandParameter("-s")
                    sKey = GetCommandParameter("-k")
                    sqllib.ConnectString = GetConnectString(sKey)

                    i = CreateXML(sTable, sObject, sModule, sParent, sDescription, sItem, sSeekKey, sKey)
                End If
            Else
                i = ProcessXML(sXML)
            End If

        Catch ex As Exception
            i = 1
            Console.WriteLine(ex.ToString)
        End Try
        If i <> 0 Then
            Console.WriteLine("press enter")
            Console.Read()
        End If
    End Sub

    Public Function CreateXML(ByVal sTable As String, ByVal sObject As String, _
            ByVal sModule As String, ByVal sParent As String, ByVal sDescription As String, _
            ByVal sItem As String, ByVal sSeekKey As String, ByVal sKey As String) As Integer
        Dim sFile As String

        Dim dsConnect As New DataSet
        Dim tc As TableColumn
        Dim node As Xml.XmlNode
        Dim node0 As Xml.XmlNode
        Dim attr As Xml.XmlAttribute
        Dim root As Xml.XmlElement
        Dim frag As Xml.XmlDocumentFragment
        Dim dom As New Xml.XmlDocument

        If sObject = "" Then
            sObject = sTable
        End If
        If sModule = "" Then
            sModule = sObject
        End If
        If sParent = "" Then
            sParent = "statics"
        End If
        If sDescription = "" Then
            sDescription = sObject
        End If
        If sItem = "" Then
            sItem = sDescription
        End If

        sFile = sObject & ".sdef"
        If CheckFile(sFile) Then
            Console.WriteLine("Error: cannot over-write existing file '" & sFile & "'.")
            Return 1
        End If

        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then
            Console.WriteLine("Error: cannot open table '" & sTable & "'.")
            Return 1
        End If

        dom.PreserveWhitespace = True
        node = dom.CreateProcessingInstruction("xml", "version='1.0'")
        dom.AppendChild(node)
        node = Nothing
        root = dom.CreateElement("roboshell")
        If sKey <> "" Then
            attr = dom.CreateAttribute("connect")
            attr.InnerText = sKey
            root.Attributes.Append(attr)
            attr = Nothing
        End If
        dom.AppendChild(root)
        root.AppendChild(dom.CreateTextNode(vbNewLine & "  "))

        ' Modules section

        frag = dom.CreateDocumentFragment

        node0 = dom.CreateElement("module")
        attr = dom.CreateAttribute("id")
        attr.InnerText = sModule
        node0.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("owner")
        attr.InnerText = sParent
        node0.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("description")
        attr.InnerText = sDescription
        node0.Attributes.Append(attr)
        attr = Nothing
        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        frag.AppendChild(node0)
        node0 = Nothing

        node0 = dom.CreateElement("module")
        attr = dom.CreateAttribute("id")
        attr.InnerText = sModule & "maintain"
        node0.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("owner")
        attr.InnerText = sModule
        node0.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("description")
        attr.InnerText = "maintain"
        node0.Attributes.Append(attr)
        attr = Nothing
        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        frag.AppendChild(node0)
        node0 = Nothing

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "  "))
        node = dom.CreateElement("modules")
        attr = dom.CreateAttribute("filename")
        attr.InnerText = "mod." & sObject & ".sql"
        node.Attributes.Append(attr)
        attr = Nothing
        node.AppendChild(frag)
        frag = Nothing
        root.AppendChild(node)
        root.AppendChild(dom.CreateTextNode(vbNewLine))
        root.AppendChild(dom.CreateTextNode(vbNewLine & "  "))
        node = Nothing

        ' Tables section

        frag = dom.CreateDocumentFragment
        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        node = dom.CreateElement("tabledefn")
        attr = dom.CreateAttribute("table")
        attr.InnerText = sTable
        node.Attributes.Append(attr)
        attr = Nothing
        frag.AppendChild(node)
        node = Nothing

        If tDefn.hasAudit Then
            frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            node = dom.CreateElement("tableauditdefn")
            attr = dom.CreateAttribute("table")
            attr.InnerText = sTable
            node.Attributes.Append(attr)
            attr = Nothing
            frag.AppendChild(node)
            node = Nothing
        End If

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "  "))
        node = dom.CreateElement("tables")
        node.AppendChild(frag)
        frag = Nothing
        root.AppendChild(node)
        node = Nothing
        root.AppendChild(dom.CreateTextNode(vbNewLine))
        frag = Nothing
        root.AppendChild(dom.CreateTextNode(vbNewLine & "  "))
        node = Nothing

        ' Objects section

        frag = dom.CreateDocumentFragment
        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        node = dom.CreateElement("procget")
        attr = dom.CreateAttribute("table")
        attr.InnerText = sTable
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("procname")
        attr.InnerText = sTable & "Get"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("objectname")
        attr.InnerText = sObject & "Get"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("module")
        attr.InnerText = sModule
        node.Attributes.Append(attr)
        attr = Nothing
        '"process"
        '"success"
        frag.AppendChild(node)
        node = Nothing

        If tDefn.PKeys.Length = 1 Then
            frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            node = dom.CreateElement("proclist")
            attr = dom.CreateAttribute("table")
            attr.InnerText = sTable
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("procname")
            attr.InnerText = sTable & "List"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("objectname")
            attr.InnerText = sObject & "List"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("module")
            attr.InnerText = "public"
            node.Attributes.Append(attr)
            attr = Nothing
            '"process"
            frag.AppendChild(node)
            node = Nothing
        End If

        If tDefn.hasAudit Then
            frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            node = dom.CreateElement("procauditget")
            attr = dom.CreateAttribute("table")
            attr.InnerText = sTable
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("procname")
            attr.InnerText = sTable & "AuditGet"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("objectname")
            attr.InnerText = sObject & "AuditGet"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("module")
            attr.InnerText = sModule
            node.Attributes.Append(attr)
            attr = Nothing
            frag.AppendChild(node)
            node = Nothing
        End If

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        node = dom.CreateElement("procinsert")
        attr = dom.CreateAttribute("table")
        attr.InnerText = sTable
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("objectname")
        attr.InnerText = sObject & "Insert"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("formobject")
        attr.InnerText = sObject & "Add"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("module")
        attr.InnerText = sModule & "maintain"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("item")
        attr.InnerText = sItem
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("mode")
        attr.InnerText = "all"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("success")
        attr.InnerText = sObject & "Get"
        node.Attributes.Append(attr)
        attr = Nothing

        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" And Not tc.Identity Then
                node.AppendChild(dom.CreateTextNode(vbNewLine & "      "))
                node0 = dom.CreateElement("field")
                attr = dom.CreateAttribute("name")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("label")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("width")
                attr.InnerText = GetFieldWidth(tc.Length)
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("helptext")
                attr.InnerText = "help text here"
                node0.Attributes.Append(attr)
                attr = Nothing
                node.AppendChild(node0)
                node0 = Nothing
            End If
        Next
        node.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        frag.AppendChild(node)
        node = Nothing

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        node = dom.CreateElement("procupdate")
        attr = dom.CreateAttribute("table")
        attr.InnerText = sTable
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("procname")
        attr.InnerText = sTable & "Update"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("objectname")
        attr.InnerText = sObject & "Update"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("formobject")
        attr.InnerText = sObject & "Edit"
        node.Attributes.Append(attr)
        attr = Nothing
        If sSeekKey <> "" Then
            attr = dom.CreateAttribute("seekkey")
            attr.InnerText = sSeekKey
            node.Attributes.Append(attr)
            attr = Nothing
        End If
        attr = dom.CreateAttribute("module")
        attr.InnerText = sModule & "maintain"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("item")
        attr.InnerText = sItem
        node.Attributes.Append(attr)
        attr = Nothing

        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" Then
                node.AppendChild(dom.CreateTextNode(vbNewLine & "      "))
                node0 = dom.CreateElement("field")
                attr = dom.CreateAttribute("name")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("label")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("width")
                attr.InnerText = GetFieldWidth(tc.Length)
                node0.Attributes.Append(attr)
                attr = Nothing
                If Not tc.Primary Then
                    attr = dom.CreateAttribute("helptext")
                    attr.InnerText = "help text here"
                    node0.Attributes.Append(attr)
                    attr = Nothing
                End If
                node.AppendChild(node0)
                node0 = Nothing
            End If
        Next
        node.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        frag.AppendChild(node)
        node = Nothing

        If tDefn.hasState Then
            frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            node = dom.CreateElement("procdisable")
            attr = dom.CreateAttribute("table")
            attr.InnerText = sTable
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("objectname")
            attr.InnerText = sObject & "Disable"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("module")
            attr.InnerText = sModule & "maintain"
            node.Attributes.Append(attr)
            attr = Nothing
            If sSeekKey <> "" Then
                attr = dom.CreateAttribute("seekkey")
                attr.InnerText = sSeekKey
                node.Attributes.Append(attr)
                attr = Nothing
            End If
            attr = dom.CreateAttribute("procname")
            attr.InnerText = sTable & "Disable"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("item")
            attr.InnerText = sItem
            node.Attributes.Append(attr)
            attr = Nothing
            frag.AppendChild(node)
            node = Nothing
        End If

        If tDefn.hasAudit Then
            frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            node = dom.CreateElement("configaudit")
            attr = dom.CreateAttribute("table")
            attr.InnerText = sTable
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("objectname")
            attr.InnerText = sObject & "Audit"
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("module")
            attr.InnerText = sModule
            node.Attributes.Append(attr)
            attr = Nothing
            attr = dom.CreateAttribute("item")
            attr.InnerText = sItem
            node.Attributes.Append(attr)
            attr = Nothing
            For Each tc In tDefn
                If tc.Name <> "AuditID" And tc.Name <> "State" Then
                    node.AppendChild(dom.CreateTextNode(vbNewLine & "      "))
                    node0 = dom.CreateElement("field")
                    attr = dom.CreateAttribute("name")
                    attr.InnerText = tc.Name
                    node0.Attributes.Append(attr)
                    attr = Nothing
                    attr = dom.CreateAttribute("label")
                    attr.InnerText = tc.Name
                    node0.Attributes.Append(attr)
                    attr = Nothing
                    attr = dom.CreateAttribute("width")
                    attr.InnerText = GetFieldWidth(tc.Length)
                    node0.Attributes.Append(attr)
                    attr = Nothing
                    node.AppendChild(node0)
                    node0 = Nothing
                End If
            Next
            node.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
            frag.AppendChild(node)
            frag.AppendChild(node)
            node = Nothing
        End If

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        node = dom.CreateElement("configgrid")
        attr = dom.CreateAttribute("table")
        attr.InnerText = sTable
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("objectname")
        attr.InnerText = sObject
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("process")
        attr.InnerText = sObject & "Get"
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("module")
        attr.InnerText = sModule
        node.Attributes.Append(attr)
        attr = Nothing
        If sSeekKey <> "" Then
            attr = dom.CreateAttribute("seekkey")
            attr.InnerText = sSeekKey
            node.Attributes.Append(attr)
            attr = Nothing
        End If
        attr = dom.CreateAttribute("item")
        attr.InnerText = sItem
        node.Attributes.Append(attr)
        attr = Nothing
        attr = dom.CreateAttribute("description")
        attr.InnerText = sDescription
        node.Attributes.Append(attr)
        attr = Nothing

        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" Then
                node.AppendChild(dom.CreateTextNode(vbNewLine & "      "))
                node0 = dom.CreateElement("field")
                attr = dom.CreateAttribute("name")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("label")
                attr.InnerText = tc.Name
                node0.Attributes.Append(attr)
                attr = Nothing
                attr = dom.CreateAttribute("width")
                attr.InnerText = GetFieldWidth(tc.Length)
                node0.Attributes.Append(attr)
                attr = Nothing
                node.AppendChild(node0)
                node0 = Nothing
            End If
        Next
        node.AppendChild(dom.CreateTextNode(vbNewLine & "    "))
        frag.AppendChild(node)
        node = Nothing

        frag.AppendChild(dom.CreateTextNode(vbNewLine & "  "))
        node = dom.CreateElement("objects")
        node.AppendChild(frag)
        frag = Nothing
        root.AppendChild(node)
        node = Nothing
        root.AppendChild(dom.CreateTextNode(vbNewLine))

        dom.Save(sFile)
        Return 0
    End Function

    Public Function ProcessXML(ByVal sXML As String) As Integer
        Dim s As String
        Dim sObject As String
        Dim sOut As String
        Dim dom As New Xml.XmlDocument
        Dim x As Xml.XmlElement
        Dim files As Xml.XmlElement

        sObject = System.IO.Path.GetFileNameWithoutExtension(sXML)
        dom.Load(sXML)
        If dom.DocumentElement.Name <> "roboshell" Then
            Console.WriteLine("Error: XML file format error")
            Return 1
        End If

        s = ""
        For Each a As Xml.XmlAttribute In dom.DocumentElement.Attributes
            Select Case a.Name
                Case "connect"
                    s = a.InnerText
            End Select
        Next
        sqllib.ConnectString = GetConnectString(s)

        For Each x In dom.DocumentElement.ChildNodes
            Select Case x.Name
                Case "objects"
                    For Each files In x.ChildNodes
                        Select Case files.Name
                            Case "storedproc"
                                SProc(files)
                            Case "storedprocconfig"
                                SProcConfig(files)

                            Case "procget"
                                TableGet(files)

                            Case "proclist"
                                TableList(files)

                            Case "procauditget"
                                TableAuditGet(files)

                            Case "procinsert"
                                TableInsert(files)

                            Case "procupdate"
                                TableUpdate(files)

                            Case "procdisable"
                                TableDisable(files)

                            Case "configaudit"
                                ConfigTableAudit(files)

                            Case "configgrid"
                                ConfigTable(files)

                            Case "scl"
                                s = ""
                                For Each a As Xml.XmlAttribute In files.Attributes
                                    Select Case a.Name
                                        Case "value"
                                            s = a.InnerText
                                    End Select
                                Next
                                If s <> "" Then MainSCL &= s & vbCrLf
                        End Select
                    Next

                Case "tables"
                    For Each files In x.ChildNodes
                        Select Case files.Name
                            Case "tabledefn"
                                Table(files)
                            Case "tableauditdefn"
                                TableAudit(files)
                            Case "staticdata"
                                Data(files)
                        End Select
                    Next

                Case "modules"
                    sOut = ""
                    For Each files In x.ChildNodes
                        If files.Name = "module" Then
                            sOut &= ModuleDefn(files)
                        End If
                    Next
                    s = sObject
                    For Each a As Xml.XmlAttribute In x.Attributes
                        Select Case a.Name
                            Case "filename"
                                s = a.InnerText
                                Exit For
                        End Select
                    Next
                    ModuleWrite(s, sOut)
            End Select
        Next
        WriteSCL(sObject)
        Return 0
    End Function

    Private Function Table(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sOut As String
        Dim i As Integer

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                    Exit For
            End Select
        Next

        If sTable = "" Then
            Return -1
        End If

        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then
            Console.WriteLine("Error: cannot open table '" & sTable & "'.")
            Return -1
        End If

        sOut = Header("") & vbCrLf & tDefn.FullTableText & Footer(True)
        PutFile("table." & tDefn.TableName & ".sql", sOut)

        i = TableIndexes(tDefn)
        If i = 0 Then i = TableFKeys(tDefn)
        Return i
    End Function

    Private Function TableIndexes(ByVal tDefn As TableColumns) As Integer
        Dim s As String
        Dim sOut As String = ""

        For Each s In tDefn.IKeys
            sOut = tDefn.IndexText(s)
            If sOut <> "" Then
                sOut = Header("") & sOut & Footer(True)
                PutFile("index." & tDefn.TableName & "." & s & ".sql", sOut)
            End If
        Next
    End Function

    Private Function TableFKeys(ByVal tDefn As TableColumns) As Integer
        Dim s As String
        Dim sOut As String = ""

        For Each s In tDefn.FKeys
            sOut = tDefn.FKeyText(s)
            If sOut <> "" Then
                sOut = Header("") & sOut & Footer(True)
                PutFile("fkey." & tDefn.TableName & "." & tDefn.LinkedTable(s) & "." & s & ".sql", sOut)
            End If
        Next
    End Function

    Private Function TableAudit(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
            End Select
        Next

        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then
            Console.WriteLine("Error: cannot open table '" & sTable & "'.")
            Return -1
        End If
        Dim aDefn As TableColumns
        Dim sName As String
        Dim sOut As String
        Dim tc As TableColumn

        If Not tDefn.hasAudit Then
            Return 0
        End If

        sName = tDefn.TableName & "Audit"
        aDefn = New TableColumns(sName, sqllib, True)
        If aDefn.State = 3 Then
            aDefn = New TableColumns()
            aDefn.TableName = sName
            For Each tc In tDefn        ' Columns
                If tc.Name = "AuditID" Or tc.Primary Then
                    tc.Identity = False
                    aDefn.AddColumn(tc)
                    aDefn.AddPKey(tc.Name, tc.Descend)
                End If
            Next
            For Each tc In tDefn        ' Columns
                If tc.Name <> "AuditID" And Not tc.Primary Then
                    aDefn.AddColumn(tc)
                End If
            Next
            aDefn.AddColumn("ActionType", "char", 1, 0, 0, "N", False, "", "")
            aDefn.AddColumn("AuditTime", "datetime", 0, 0, 0, "N", False, sName & "AuditTime", "(getdate())")
            aDefn.AddColumn("UserID", "sysname", 0, 0, 0, "N", False, sName & "UserID", "(suser_sname())")
        End If

        sOut = Header("") & vbCrLf & aDefn.FullTableText & Footer(True)

        PutFile("table." & sName & ".sql", sOut)
        Return AuditInsert(aDefn)
    End Function

    Private Function AuditInsert(ByVal tDefn As TableColumns) As Integer
        Dim sName As String
        Dim sOut As String
        Dim Comma As String
        Dim s As String
        Dim sel As String
        Dim w As String
        Dim sW As String
        Dim i As Integer
        Dim tc As TableColumn

        sName = tDefn.TableName
        sOut = Header(sName & "Insert")

        sW = ""
        w = "    where   "
        Comma = " "
        For Each s In tDefn.PKeys        ' Primary Key
            tc = tDefn.Column(s)        ' Columns

            If s <> "AuditID" Then
                sOut &= "   " & Comma & "@" & s & " " & tc.TypeText & vbCrLf
                sW &= w & "a." & s & " = @" & s & vbCrLf
                w = "    and     "
                Comma = ","
            End If
        Next

        sOut &= "   ,@AuditID integer" & vbCrLf
        sOut &= "   ,@ActionType char(1)" & vbCrLf
        sOut &= "as" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    set nocount on" & vbCrLf
        sOut &= vbCrLf
        sOut &= "    insert into dbo." & sName & vbCrLf
        sOut &= "    ("

        i = 4
        w = "    select  "
        sel = ""
        Comma = ""
        For Each tc In tDefn
            If Not tc.Primary Or tc.Name = "UserID" Or tc.Name = "AuditTime" Then
                Continue For
            End If

            i += 1
            If i > 2 Then
                sOut &= Comma & vbCrLf + "        " & tc.Name
                i = 0
            Else
                sOut &= Comma & " " & tc.Name
            End If
            sel &= w & "@" & tc.Name & vbCrLf
            Comma = ","
            w = "           ,"
        Next

        i = 4
        For Each tc In tDefn
            If tc.Primary Or tc.Name = "UserID" Or tc.Name = "AuditTime" _
                    Or tc.Name = "AuditID" Then
                Continue For
            End If

            i += 1
            If i > 2 Then
                sOut &= Comma & vbCrLf + "        " & tc.Name
                i = 0
            Else
                sOut &= Comma & " " & tc.Name
            End If
            If tc.Name = "ActionType" Then
                sel &= w & "@" & tc.Name & vbCrLf
            Else
                sel &= w & "a." & tc.Name & vbCrLf
            End If
        Next
        sOut &= vbCrLf
        sOut &= "    )" & vbCrLf
        sOut &= sel
        sOut &= "    from    dbo." & Mid(sName, 1, Len(sName) - 5) & " a" & vbCrLf
        sOut &= sW
        sOut &= vbCrLf
        sOut &= "    return @@error" & vbCrLf
        sOut &= "end" & vbCrLf
        sOut &= Footer(True)

        PutFile("proc." & sName & "Insert.sql", sOut)
        Return 0
    End Function

    Private Function Data(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sFilter As String = ""
        Dim sName As String = ""
        Dim sOut As String
        Dim ss As String = ""

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "filter"
                    sFilter = a.InnerText
                Case "filename"
                    sName = a.InnerText
            End Select
        Next

        If sTable = "" Then
            Return -1
        End If

        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then
            Console.WriteLine("Error: cannot open table '" & sTable & "'.")
            Return -1
        End If

        sOut = Header("") & vbCrLf
        sOut &= tDefn.DataScript(sFilter)
        sOut &= Footer(True)

        If sName = "" Then
            sName = "data." & tDefn.TableName & ".sql"
        End If
        PutFile(sName, sOut)
        Return 0
    End Function

    Private Function SProc(ByVal files As Xml.XmlElement) As Integer
        Dim sName As String = ""
        Dim sOut As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "name"
                    sName = a.InnerText
                    Exit For
            End Select
        Next

        If sName = "" Then
            Return -1
        End If

        Dim pDefn As New StoredProcedure(sName, sqllib)

        sOut = Header(sName)
        sOut &= vbCrLf
        sOut &= pDefn.FullText
        sOut &= Footer(True)

        PutFile("proc." & pDefn.ProcedureName & ".sql", sOut)
        Return 0
    End Function

    Private Function SProcConfig(ByVal files As Xml.XmlElement) As Integer
        Dim ProcedureName As String = ""
        Dim ConfigName As String = ""
        Dim ModuleName As String = "public"
        Dim ProcessName As String = ""
        Dim SuccessName As String = ""
        Dim Messages As String = ""
        Dim Mode As String = "D"
        Dim sOut As String

        'connectkey

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "procname"
                    ProcedureName = a.InnerText
                Case "objectname"
                    ConfigName = a.InnerText
                Case "module"
                    ModuleName = a.InnerText
                Case "process"
                    ProcessName = a.InnerText
                Case "success"
                    SuccessName = a.InnerText
                Case "messages"
                    Messages = a.InnerText
                Case "mode"
                    Mode = a.InnerText
            End Select
        Next

        If ProcedureName = "" Then
            Return -1
        End If

        Dim pDefn As New StoredProcedure(ProcedureName, sqllib)

        If ConfigName = "" Then
            ConfigName = ProcedureName
        End If
        pDefn.ConfigName = ConfigName
        pDefn.ProcessName = ProcessName
        pDefn.SuccessName = SuccessName
        If Messages = "N" Then
            pDefn.Messages = False
        End If
        pDefn.Mode = Mode

        sOut = Header("")
        sOut &= vbCrLf
        sOut &= pDefn.ConfigText
        sOut &= Footer(True)

        PutFile("config." & ConfigName & ".sql", sOut)
        Return 0
    End Function

    Private Function ModuleDefn(ByVal files As Xml.XmlElement) As String
        Dim sID As String = ""
        Dim sOwner As String = ""
        Dim sDescription As String = ""
        Dim sOut As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "id"
                    sID = a.InnerText
                Case "owner"
                    sOwner = a.InnerText
                Case "description"
                    sDescription = a.InnerText
            End Select
        Next

        If sID = "" Then Return ""
        sID = LCase(sID)
        sOwner = LCase(sOwner)

        sOut = vbCrLf
        sOut &= "execute dbo.shlModulesInsert" & vbCrLf
        sOut &= "    @ModuleID = '" & sID & "'" & vbCrLf
        sOut &= "   ,@OwnerModule = '" & sOwner & "'" & vbCrLf
        sOut &= "   ,@Description = '" & sDescription & "'" & vbCrLf
        sOut &= "go" & vbCrLf

        Return sOut
    End Function

    Private Sub ModuleWrite(ByVal sFileName As String, ByVal sText As String)
        Dim sOut As String

        If sText <> "" Then
            sOut = Header("") & sText & Footer(False)
            PutFile(sFileName, sOut)
        End If
    End Sub

    Private Function TableGet(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sProcName As String = ""
        Dim sObject As String = ""
        Dim sModule As String = ""
        Dim sProcess As String = ""
        Dim sSuccess As String = ""
        Dim sOut As String
        Dim w As String
        Dim sW As String
        Dim s As String
        Dim o As String
        Dim sO As String
        Dim tc As TableColumn
        Dim Comma As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "process"
                    sProcess = a.InnerText
                Case "success"
                    sSuccess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1

        If sProcName = "" Then sProcName = tDefn.TableName & "Get"
        If sObject = "" Then sObject = tDefn.TableName & "Get"
        If sModule = "" Then sModule = sObject
        If sProcess = "" Then sProcess = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)

        If sp.State <> 2 Then
            sp.ProcedureName = sProcName
            sOut = "create procedure dbo." & sProcName & vbCrLf

            sW = ""
            w = "    where   "
            sO = ""
            o = "    order by "
            Comma = " "
            For Each s In tDefn.PKeys        ' Primary Key
                tc = tDefn.Column(s)
                sOut &= "   " & Comma & "@p" & tc.Name & " " & tc.TypeText & " = null" & vbCrLf
                Comma = ","

                sW &= w & "a." & tc.Name & " = coalesce(@p" & tc.Name & ", a." & tc.Name & ")" & vbCrLf
                w = "    and     "
                sO &= o & "a." & tc.Name & vbCrLf
                o = "           ,"

                sp.AddParameter("@p" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
            Next
            sOut &= "as" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    set nocount on" & vbCrLf
            sOut &= vbCrLf

            w = "    select  a."
            For Each tc In tDefn
                sOut &= w & tc.Name & vbCrLf
                w = "           ,a."
                If tc.Name = "State" Then
                    sOut &= "           ,v.ValueDescription StateName" & vbCrLf
                End If
            Next

            sOut &= "    from    dbo." & tDefn.TableName & " a" & vbCrLf
            If tDefn.hasState Then
                sOut &= "    join    dbo.shlTableValues v" & vbCrLf
                sOut &= "    on      v.TableName = 'default'" & vbCrLf
                sOut &= "    and     v.ColumnName = 'State'" & vbCrLf
                sOut &= "    and     v.ColumnValue = a.State" & vbCrLf
            End If
            sOut &= sW
            sOut &= sO
            sOut &= "end" & vbCrLf

            sp.ProcedureText = sOut
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)

        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess
        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function TableList(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sProcName As String = ""
        Dim sTextColumn As String = ""
        Dim sObject As String = ""
        Dim sModule As String = "public"
        Dim sProcess As String = ""
        Dim sOut As String
        Dim sS As String = ""
        Dim sW As String = ""
        Dim s As String
        Dim sO As String = ""
        Dim tc As TableColumn

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "textcolumn"
                    sTextColumn = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "process"
                    sProcess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1
        If sProcName = "" Then sProcName = sTable & "List"
        If sObject = "" Then sObject = sProcName
        If sProcess = "" Then sProcess = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)

        If sp.State <> 2 Then
            sp.ProcedureName = sProcName

            sOut = "create procedure dbo." & sProcName & vbCrLf
            For Each s In tDefn.PKeys        ' Primary Key
                tc = tDefn.Column(s)
                sS = "    select  a." & tc.Name & vbCrLf
                If tDefn.hasState Then
                    sOut &= "    @" & tc.Name & " " & tc.TypeText & " = null" & vbCrLf
                    sW = "    where   a.State = 'ac' or a." & tc.Name & " = @" & tc.Name & vbCrLf
                    sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
                End If
            Next
            sOut &= "as" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    set nocount on" & vbCrLf
            sOut &= vbCrLf
            sOut &= sS

            For Each tc In tDefn
                If tc.Primary = False And sTextColumn = "" Then
                    sTextColumn = tc.Name
                End If

                If tc.Name = sTextColumn Then
                    If Not tc.Primary Then
                        sOut &= "           ,a." & tc.Name & vbCrLf
                    End If
                    sO &= "    order by a." & sTextColumn & vbCrLf
                    Exit For
                End If
            Next

            sOut &= "    from    dbo." & tDefn.TableName & " a" & vbCrLf
            If tDefn.hasState Then
                sOut &= sW
            End If
            sOut &= sO
            sOut &= "end" & vbCrLf

            sp.ProcedureText = sOut
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)

        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess
        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function TableAuditGet(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sProcName As String = ""
        Dim sObject As String = ""
        Dim sModule As String = ""
        Dim sProcess As String = ""

        Dim sOut As String
        Dim w As String
        Dim sW As String
        Dim s As String
        Dim tc As TableColumn
        Dim Comma As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "process"
                    sProcess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1
        If Not tDefn.hasAudit Then
            Return -1
        End If
        If sProcName = "" Then sProcName = sTable & "AuditGet"
        If sObject = "" Then sObject = sProcName
        If sProcess = "" Then sProcess = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)

        If sp.State <> 2 Then
            sp.ProcedureName = sProcName

            sOut = "create procedure dbo." & sProcName & vbCrLf
            sW = ""
            w = "    where   "
            Comma = " "
            For Each s In tDefn.PKeys        ' Primary Key
                tc = tDefn.Column(s)
                sOut &= "   " & Comma & "@" & tc.Name & " " & tc.TypeText & vbCrLf
                Comma = ","

                sW &= w & "a." & tc.Name & " = @" & tc.Name & vbCrLf
                w = "    and     "

                sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
            Next
            sOut &= "as" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    set nocount on" & vbCrLf
            sOut &= vbCrLf
            sOut &= "    select  a.AuditID" & vbCrLf
            sOut &= "           ,v.ValueDescription Action" & vbCrLf
            sOut &= "           ,a.ActionType" & vbCrLf
            sOut &= "           ,a.UserID" & vbCrLf
            sOut &= "           ,a.AuditTime" & vbCrLf

            For Each tc In tDefn
                If Not tc.Primary And tc.Name <> "AuditID" Then
                    If tc.Name = "State" Then
                        sOut &= "           ,t.ValueDescription State" & vbCrLf
                    Else
                        sOut &= "           ,a." & tc.Name & vbCrLf
                    End If
                End If
            Next

            sOut &= "    from    dbo." & tDefn.TableName & "Audit a" & vbCrLf
            sOut &= "    join    dbo.shlTableValues v" & vbCrLf
            sOut &= "    on      v.TableName = 'default'" & vbCrLf
            sOut &= "    and     v.ColumnName = 'ActionType'" & vbCrLf
            sOut &= "    and     v.ColumnValue = a.ActionType" & vbCrLf
            If tDefn.hasState Then
                sOut &= "    join    dbo.shlTableValues t" & vbCrLf
                sOut &= "    on      t.TableName = 'default'" & vbCrLf
                sOut &= "    and     t.ColumnName = 'State'" & vbCrLf
                sOut &= "    and     t.ColumnValue = a.State" & vbCrLf
            End If
            sOut &= sW
            sOut &= "    order by a.AuditID" & vbCrLf
            sOut &= "end" & vbCrLf

            sp.ProcedureText = sOut
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)

        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess
        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function TableInsert(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sProcName As String = ""
        Dim sObject As String = ""
        Dim sFormObject As String = ""
        Dim sModule As String = "public"
        Dim sProcess As String = ""
        Dim ssItem As String = ""
        Dim sSuccess As String = ""
        Dim sMode As String = ""
        Dim Fields As System.Xml.XmlNodeList = files.ChildNodes
        Dim sOut As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "formobject"
                    sFormObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "item"
                    ssItem = a.InnerText
                Case "mode"
                    sMode = LCase(a.InnerText)
                Case "process"
                    sProcess = a.InnerText
                Case "success"
                    sSuccess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1

        If sProcName = "" Then sProcName = sTable & "Insert"
        If sFormObject = "" Then sFormObject = sTable & "Add"
        If sObject = "" Then sObject = sProcName
        If sProcess = "" Then sProcess = sObject
        If sSuccess = "" Then sSuccess = sTable & "Get"
        If ssItem = "" Then ssItem = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)
        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess

        If sp.State <> 2 Then
            sp.ProcedureName = sProcName

            If tDefn.hasIdentity Then
                InsertProcIdent(sp, tDefn, ssItem, sMode)
            Else
                InsertProc(sp, tDefn, ssItem, sMode)
            End If
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)

        If sMode = "all" Then
            sp.Mode = "X"
            sp.SuccessName = sSuccess
            sp.AddResult("@" & sSuccess, "object", 0, 0, 0)
        Else
            sp.Mode = "P"
            If tDefn.hasState Then
                sp.AddResult("State", "string", 2, 0, 0)
                sp.AddResult("StateName", "string", 50, 0, 0)
            End If
            If tDefn.hasAudit Then
                sp.AddResult("AuditID", "integer", 0, 0, 0)
            End If
        End If

        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)
        Return ConfigTableAdd(tDefn, sFormObject, sObject, ssItem, sMode, Fields, sModule, sSuccess)
    End Function

    Private Function InsertProc(ByVal sp As StoredProcedure, ByVal tDefn As TableColumns, _
        ByVal ssItem As String, ByVal sMode As String) As Integer

        Dim s As String
        Dim sOut As String
        Dim Comma As String = " "
        Dim sTab As String
        Dim tc As TableColumn
        Dim b As Boolean = False
        Dim i As Integer
        Dim w As String
        Dim sW As String

        sOut = "create procedure dbo." & sp.ProcedureName & vbCrLf
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" Then
                sOut &= "   " & Comma & "@" & tc.Name & " " & tc.TypeText
                If tc.Nullable = "Y" Then
                    sOut &= " = null"
                End If
                sOut &= vbCrLf
                Comma = ","

                sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
            End If
        Next
        sOut &= "as" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    set nocount on" & vbCrLf
        sOut &= "    declare @e integer" & vbCrLf
        If tDefn.hasAudit Then
            sOut &= "           ,@AuditID integer" & vbCrLf
        End If
        If tDefn.hasState Then
            sOut &= "           ,@State char(2)" & vbCrLf
        End If
        sOut &= vbCrLf
        sOut &= "    set @e = 0" & vbCrLf
        sOut &= "    while @e = 0" & vbCrLf
        sOut &= "    begin" & vbCrLf

        For Each s In tDefn.PKeys
            tc = tDefn.Column(s)
            If tc.Type = "char" Or tc.Type = "varchar" Or tc.Type = "nvarchar" Then
                sOut &= "        set @" & tc.Name & " = upper(@" & tc.Name & ")" & vbCrLf
                b = True
            End If
        Next
        If b Then
            sOut &= vbCrLf
        End If

        If tDefn.hasAudit Then
            sOut &= "        begin transaction" & vbCrLf
            sOut &= vbCrLf
        End If

        If tDefn.hasAudit Or tDefn.hasState Then
            If tDefn.hasAudit Then
                sOut &= "        select  @AuditID = a.AuditID" & vbCrLf
                If tDefn.hasState Then
                    sOut &= "               ,@State = a.State" & vbCrLf
                End If
            ElseIf tDefn.hasState Then
                sOut &= "        select  @State = a.State" & vbCrLf
            End If
            sOut &= "        from    dbo." & tDefn.TableName & " a" & vbCrLf
            w = "        where   a."
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                w = "        and     a."
            Next
            sOut &= vbCrLf
            sOut &= "        if @@rowcount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            If Not tDefn.hasState Then
                sOut &= "            set @e = 51000" & vbCrLf
                sOut &= "            raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
                sOut &= "            break" & vbCrLf
            Else
                If tDefn.hasAudit Then
                    sOut &= "            set @AuditID = @AuditID + 1" & vbCrLf
                    sOut &= vbCrLf
                End If
                sOut &= "            if @State = 'ac'" & vbCrLf
                sOut &= "            begin" & vbCrLf
                sOut &= "                set @e = 51000" & vbCrLf
                sOut &= "                raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
                sOut &= "                break" & vbCrLf
                sOut &= "            end" & vbCrLf
                sOut &= vbCrLf
                sOut &= "            update  dbo." & tDefn.TableName & vbCrLf
                sW = ""
                w = "            where   "
                s = "            set     "
                For Each tc In tDefn
                    If tc.Name <> "State" And tc.Name <> "AuditID" Then
                        If tc.Primary Then
                            sW &= w & tc.Name & " = @" & tc.Name & vbCrLf
                            w = "            and     "
                        Else
                            sOut &= s & tc.Name & " = @" & tc.Name & vbCrLf
                            s = "                   ,"
                        End If
                    End If
                Next
                sOut &= s & "State = 'ac'" & vbCrLf
                If tDefn.hasAudit Then
                    sOut &= "                   ,AuditID = @AuditID" & vbCrLf
                End If
                sOut &= sW
                sOut &= vbCrLf
                sOut &= "            set @e = @@error" & vbCrLf
            End If
        Else
            sOut &= "        if exists" & vbCrLf
            sOut &= "        (" & vbCrLf
            sOut &= "            select  'a'" & vbCrLf
            sOut &= "            from    dbo." & tDefn.TableName & " a" & vbCrLf
            w = "            where   "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= w & "a." & tc.Name & " = @" & tc.Name & vbCrLf
                w = "            and     "
            Next
            sOut &= "        )" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            set @e = 51000" & vbCrLf
            sOut &= "            raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
            sOut &= "            break" & vbCrLf
        End If
        sOut &= "        end" & vbCrLf
        If tDefn.hasAudit Then
            sOut &= "        else" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            set @AuditID = 1" & vbCrLf
            sOut &= vbCrLf
            sTab = "    "
        Else
            sTab = ""
        End If
        sOut &= sTab & "        insert into dbo." & tDefn.TableName & vbCrLf
        sOut &= sTab & "        ("
        i = 10
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" Then
                If i <> 10 Then
                    sOut &= ","
                End If
                If i > 2 Then
                    sOut &= vbCrLf & sTab & "           "
                    i = 0
                End If
                sOut &= " " & tc.Name
            End If
        Next
        If tDefn.hasAudit Or tDefn.hasState Then
            Comma = ""
            sOut &= "," & vbCrLf & sTab & "            "
            If tDefn.hasState Then
                sOut &= "State"
                Comma = ", "
            End If
            If tDefn.hasAudit Then
                sOut &= Comma & "AuditID"
            End If
        End If
        sOut &= vbCrLf
        sOut &= sTab & "        )" & vbCrLf
        sOut &= sTab & "        values" & vbCrLf
        sOut &= sTab & "        ("
        i = 10
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" Then
                If i <> 10 Then
                    sOut &= ","
                End If
                If i > 2 Then
                    sOut &= vbCrLf & sTab & "           "
                    i = 0
                End If
                sOut &= " @" & tc.Name
            End If
        Next
        If tDefn.hasAudit Or tDefn.hasState Then
            Comma = ""
            sOut &= "," & vbCrLf & sTab & "            "
            If tDefn.hasState Then
                sOut &= "'ac'"
                Comma = ", "
            End If
            If tDefn.hasAudit Then
                sOut &= Comma & "@AuditID"
            End If
        End If
        sOut &= vbCrLf
        sOut &= sTab & "        )" & vbCrLf
        sOut &= sTab & "        set @e = @@error" & vbCrLf
        If tDefn.hasAudit Then
            sOut &= "        end" & vbCrLf
            sOut &= "        if @e <> 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            break" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= vbCrLf
            sOut &= "        execute @e = dbo." & tDefn.TableName & "AuditInsert" & vbCrLf
            Comma = " "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                Comma = ","
            Next
            sOut &= "           ,@AuditID = @AuditID" & vbCrLf
            sOut &= "           ,@ActionType = 'I'" & vbCrLf
            sOut &= "        if @e <> 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            break" & vbCrLf
            sOut &= "        end" & vbCrLf
        End If
        sOut &= "        break" & vbCrLf
        sOut &= "    end" & vbCrLf
        If tDefn.hasAudit Then
            sOut &= "    if @e <> 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            rollback transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "    else" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            commit transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            If sMode <> "all" Then
                sOut &= "        execute dbo." & tDefn.TableName & "Get" & vbCrLf
                Comma = " "
                For Each s In tDefn.PKeys
                    tc = tDefn.Column(s)
                    sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                    Comma = ","
                Next
            End If
            sOut &= "    end" & vbCrLf
        ElseIf sMode <> "all" Then
            sOut &= "    if @e = 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        execute dbo." & tDefn.TableName & "Get" & vbCrLf
            Comma = " "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                Comma = ","
            Next
            sOut &= "    end" & vbCrLf
        End If
        sOut &= "    return @e" & vbCrLf
        sOut &= "end" & vbCrLf

        sp.ProcedureText = sOut
        Return 0
    End Function

    Private Function InsertProcIdent(ByVal sp As StoredProcedure, ByVal tDefn As TableColumns, _
        ByVal ssItem As String, ByVal sMode As String) As Integer
        Dim sOut As String
        Dim s As String
        Dim b As Boolean = False
        Dim IdentKey As Boolean = True
        Dim i As Integer
        Dim tc As TableColumn
        Dim Comma As String = " "

        sOut = "create procedure dbo." & sp.ProcedureName & vbCrLf
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" And Not tc.Identity Then
                sOut &= "   " & Comma & "@" & tc.Name & " " & tc.TypeText
                If tc.Nullable = "Y" Then
                    sOut &= " = null"
                End If
                sOut &= vbCrLf
                Comma = ","

                sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
            End If
        Next
        sOut &= "as" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    set nocount on" & vbCrLf
        sOut &= "    declare @e integer" & vbCrLf

        tc = tDefn.Column(tDefn.IdentityColumn)
        sOut &= "           ,@" & tc.Name & " " & tc.TypeText & vbCrLf
        sOut &= vbCrLf
        sOut &= "    set @e = 0" & vbCrLf
        sOut &= "    while @e = 0" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        begin transaction" & vbCrLf
        sOut &= vbCrLf
        sOut &= "        insert into dbo." & tDefn.TableName & vbCrLf
        sOut &= "        ("
        i = 10
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" And Not tc.Identity Then
                If i <> 10 Then
                    sOut &= ","
                End If
                If i > 2 Then
                    sOut &= vbCrLf & "           "
                    i = 0
                End If
                sOut &= " " & tc.Name
            End If
        Next
        If tDefn.hasAudit Or tDefn.hasState Then
            Comma = ""
            sOut &= "," & vbCrLf & "            "
            If tDefn.hasState Then
                sOut &= "State"
                Comma = ", "
            End If
            If tDefn.hasAudit Then
                sOut &= Comma & "AuditID"
            End If
            sOut &= vbCrLf
        End If
        sOut &= "        )" & vbCrLf
        sOut &= "        values" & vbCrLf
        sOut &= "        ("
        i = 10
        For Each tc In tDefn
            If tc.Name <> "AuditID" And tc.Name <> "State" And Not tc.Identity Then
                If i <> 10 Then
                    sOut &= ","
                End If
                If i > 2 Then
                    sOut &= vbCrLf & "           "
                    i = 0
                End If
                sOut &= " @" & tc.Name
            End If
        Next
        If tDefn.hasAudit Or tDefn.hasState Then
            Comma = ""
            sOut &= "," & vbCrLf & "            "
            If tDefn.hasState Then
                sOut &= "'ac'"
                Comma = ", "
            End If
            If tDefn.hasAudit Then
                sOut &= Comma & "1"
            End If
            sOut &= vbCrLf
        End If
        sOut &= "        )" & vbCrLf
        sOut &= "        select  @e = @@error" & vbCrLf
        sOut &= "               ,@" & tDefn.IdentityColumn & " = @@identity" & vbCrLf
        sOut &= "        if @e <> 0" & vbCrLf
        sOut &= "        begin" & vbCrLf
        sOut &= "            break" & vbCrLf
        sOut &= "        end" & vbCrLf
        sOut &= vbCrLf
        sOut &= "        execute @e = dbo." & tDefn.TableName & "AuditInsert" & vbCrLf
        Comma = " "
        For Each s In tDefn.PKeys
            tc = tDefn.Column(s)
            sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
            Comma = ","
        Next
        sOut &= "           ,@AuditID = 1" & vbCrLf
        sOut &= "           ,@ActionType = 'I'" & vbCrLf
        sOut &= "        break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= "    if @e <> 0" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        if @@trancount > 0" & vbCrLf
        sOut &= "        begin" & vbCrLf
        sOut &= "            rollback transaction" & vbCrLf
        sOut &= "        end" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= "    else" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        if @@trancount > 0" & vbCrLf
        sOut &= "        begin" & vbCrLf
        sOut &= "            commit transaction" & vbCrLf
        sOut &= "        end" & vbCrLf
        If sMode <> "all" Then
            sOut &= "        execute dbo." & tDefn.TableName & "Get" & vbCrLf
            Comma = " "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                Comma = ","
            Next
        End If
        sOut &= "    end" & vbCrLf
        sOut &= "    return @e" & vbCrLf
        sOut &= "end" & vbCrLf

        sp.ProcedureText = sOut
        Return 0
    End Function

    Private Function ConfigTableAdd(ByVal tDefn As TableColumns, _
                ByVal sObject As String, ByVal sProcess As String, _
                ByVal ssItem As String, ByVal sMode As String, _
                ByVal Fields As System.Xml.XmlNodeList, ByVal sOwner As String, _
                ByVal sObjectGet As String) As Integer
        Dim sLabel As String
        Dim Width As Integer
        Dim sHelp As String
        Dim sOut As String
        Dim tc As TableColumn
        Dim Process As Boolean
        Dim b As Boolean = False

        sOut = Header("")
        sOut &= vbCrLf
        sOut &= "execute dbo.shlDialogFormInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@Title = 'Create New " & ssItem & "'" & vbCrLf
        sOut &= "   ,@HelpPage = '" & sObject & ".html'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlProcessesInsert" & vbCrLf
        sOut &= "    @ProcessName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ModuleID = '" & LCase(sOwner) & "'" & vbCrLf
        sOut &= "   ,@ObjectName = '" & sObject & "'" & vbCrLf
        If sMode = "all" Then
            sOut &= "   ,@UpdateParent = 'Y'" & vbCrLf
        End If
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        If sMode = "all" Then
            sOut &= "execute dbo.shlParametersInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ParameterName = '" & sObjectGet & "'" & vbCrLf
            sOut &= "   ,@ValueType = 'object'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        End If
        For Each tc In tDefn
            If tc.Name = "AuditID" Or tc.Name = "State" Or tc.Identity Then
                Continue For
            End If

            sLabel = ""
            Width = -1
            sHelp = ""
            Process = False
            For Each n As System.Xml.XmlNode In Fields
                For Each att As System.Xml.XmlAttribute In n.Attributes
                    Select Case LCase(att.Name)
                        Case "name"
                            If LCase(tc.Name) <> LCase(att.Value) Then
                                Exit For
                            End If
                            Process = True
                        Case "label"
                            sLabel = att.Value
                            If sLabel = tc.Name Then sLabel = ""
                        Case "width"
                            Width = CInt(att.Value)
                        Case "helptext"
                            sHelp = att.Value
                    End Select
                Next
                If Process Then Exit For
            Next

            sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@FieldName = '" & tc.Name & "'" & vbCrLf
            If sLabel <> "" Then
                sOut &= "   ,@Label = '" & sLabel & "'" & vbCrLf
            End If
            sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbCrLf
            If tc.vbType = "string" Then
                sOut &= "   ,@Width = " & tc.Length & vbCrLf
                If Width <> -1 Then
                    sOut &= "   ,@DisplayWidth = " & Width & vbCrLf
                Else
                    sOut &= "   ,@DisplayWidth = " & (tc.Length * 5) & vbCrLf
                End If
            Else
                If Width <> -1 Then
                    sOut &= "   ,@DisplayWidth = " & Width & vbCrLf
                Else
                    sOut &= "   ,@DisplayWidth = 60" & vbCrLf
                End If
            End If
            If tc.Primary Then
                sOut &= "   ,@IsPrimary = 'Y'" & vbCrLf
            End If
            If tc.Nullable = "N" Then
                sOut &= "   ,@Required = 'Y'" & vbCrLf
            End If
            sOut &= "   ,@IsInput = 'N'" & vbCrLf

            If sHelp <> "" Then
                sOut &= "   ,@HelpText = '" & sHelp & "'" & vbCrLf
            Else
                sOut &= "   ,@HelpText = 'Place help text here!'" & vbCrLf
            End If
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        Next
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Okay'" & vbCrLf
        sOut &= "   ,@Process = '" & sProcess & "'" & vbCrLf
        sOut &= "   ,@Validate = 'Y'" & vbCrLf
        If sMode = "all" Then
            sOut &= "   ,@CloseObject = 'O'" & vbCrLf
        Else
            sOut &= "   ,@CloseObject = 'Y'" & vbCrLf
        End If
        sOut &= "   ,@ImageFile = 'okay.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Save changes and exit'" & vbCrLf
        sOut &= "   ,@KeyCode = 13" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Cancel'" & vbCrLf
        sOut &= "   ,@CloseObject = 'Y'" & vbCrLf
        sOut &= "   ,@ImageFile = 'cancel.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Exit without saving changes'" & vbCrLf
        sOut &= "   ,@KeyCode = 27" & vbCrLf
        sOut &= Footer(True)

        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function TableUpdate(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sModule As String = "public"
        Dim ssItem As String = ""
        Dim Fields As System.Xml.XmlNodeList = files.ChildNodes
        Dim sObject As String = ""
        Dim sProcName As String = ""
        Dim sFormObject As String = ""
        Dim sProcess As String = ""
        Dim sSeekKey As String = ""
        Dim sOut As String
        Dim s As String
        Dim b As Boolean = False
        Dim w As String
        Dim tc As TableColumn
        Dim Comma As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "formobject"
                    sFormObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "seekkey"
                    sSeekKey = a.InnerText
                Case "item"
                    ssItem = a.InnerText
                Case "process"
                    sProcess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1
        If sProcName = "" Then sProcName = tDefn.TableName & "Update"
        If sFormObject = "" Then sFormObject = sTable & "Edit"
        If sObject = "" Then sObject = tDefn.TableName & "Update"
        If sProcess = "" Then sProcess = sObject
        If sSeekKey = "" Then sSeekKey = sObject
        If ssItem = "" Then ssItem = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)
        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess

        If sp.State <> 2 Then
            sp.ProcedureName = sProcName

            sOut = "create procedure dbo." & sp.ProcedureName & vbCrLf
            Comma = " "
            For Each tc In tDefn
                If tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= "   " & Comma & "@" & tc.Name & " " & tc.TypeText
                    If tc.Nullable = "Y" Then
                        sOut &= " = null"
                    End If
                    sOut &= vbCrLf
                    Comma = ","
                    sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
                End If
            Next
            If tDefn.hasAudit Then
                sOut &= "   ,@AuditID integer" & vbCrLf
                sp.AddParameter("@AuditID", "integer", 4, 0, 0, "in")
            End If
            sOut &= "as" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    set nocount on" & vbCrLf
            sOut &= "    declare @e integer" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "           ,@AudID integer" & vbCrLf
            End If

            For Each tc In tDefn
                If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= "           ,@Old" & tc.Name & " " & tc.TypeText & vbCrLf
                End If
            Next
            sOut &= vbCrLf
            sOut &= "    set @e = 0" & vbCrLf
            sOut &= "    while @e = 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            Comma = "        select  @Old"
            For Each tc In tDefn
                If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= Comma & tc.Name & " = a." & tc.Name & vbCrLf
                    Comma = "               ,@Old"
                    b = True
                End If
            Next
            If Not b Then
                Return 0    ' no editable fields to update
            End If

            If tDefn.hasAudit Then
                sOut &= "               ,@AudID = a.AuditID" & vbCrLf
            End If
            sOut &= "        from    dbo." & tDefn.TableName & " a (holdlock)" & vbCrLf
            w = "        where   a."
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                w = "        and     a."
            Next
            sOut &= vbCrLf
            sOut &= "        if @@rowcount = 0  -- not found" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            set @e = 51001" & vbCrLf
            sOut &= "            raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
            sOut &= "            break" & vbCrLf
            sOut &= "        end" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= vbCrLf
                sOut &= "        if @AudID <> @AuditID   -- already changed" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            set @AudID = -1" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
                sOut &= "        set @AudID = @AudID + 1" & vbCrLf
            End If
            Comma = "        if coalesce(@Old"
            For Each tc In tDefn
                If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= vbCrLf & Comma & tc.Name & ", '') = coalesce(@" & tc.Name & ", '')"
                    Comma = "        and coalesce(@Old"
                End If
            Next
            sOut &= vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            set @e = 51002" & vbCrLf
            sOut &= "            raiserror (@e, 16, 1)" & vbCrLf
            sOut &= "            break" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= vbCrLf
            sOut &= "        begin transaction" & vbCrLf
            sOut &= vbCrLf
            sOut &= "        update  dbo." & tDefn.TableName & vbCrLf
            Comma = "        set     "
            For Each tc In tDefn
                If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= Comma & tc.Name & " = @" & tc.Name & vbCrLf
                    Comma = "               ,"
                End If
            Next
            If tDefn.hasAudit Then
                sOut &= "               ,AuditID = @AudID" & vbCrLf
            End If
            w = "        where   "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                w = "        and     "
            Next
            sOut &= vbCrLf
            sOut &= "        set @e = @@error" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "        if @e <> 0" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
                sOut &= vbCrLf
                sOut &= "        execute @e = dbo." & tDefn.TableName & "AuditInsert" & vbCrLf
                Comma = " "
                For Each s In tDefn.PKeys
                    tc = tDefn.Column(s)
                    sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                    Comma = ","
                Next
                sOut &= "           ,@AuditID = @AudID" & vbCrLf
                sOut &= "           ,@ActionType = 'U'" & vbCrLf
            End If
            sOut &= "        break" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "    if @e <> 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            rollback transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "    else" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            commit transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= "        execute dbo." & tDefn.TableName & "Get" & "    -- return the changes" & vbCrLf
            Comma = " "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= "           " & Comma & "@p" & tc.Name & " = @" & tc.Name & vbCrLf
                Comma = ","
            Next
            If tDefn.hasAudit Then
                sOut &= vbCrLf
                sOut &= "        if @AudID = -1" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            print 'This " & ssItem & " has changed since it was retrieved'" & vbCrLf
                sOut &= "        end" & vbCrLf
            End If
            sOut &= "    end" & vbCrLf
            sOut &= "    return @e" & vbCrLf
            sOut &= "end" & vbCrLf

            sp.ProcedureText = sOut
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)
        sp.Mode = "P"
        sp.SeekKey = sSeekKey
        sp.AddResult("@StateName", "varchar", 50, 0, 0)
        sp.AddResult("@State", "varchar", 2, 0, 0)
        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)

        Return ConfigTableEdit(tDefn, sFormObject, sObject, ssItem, Fields, sModule)
    End Function

    Private Function ConfigTableEdit(ByVal tDefn As TableColumns, ByVal sObject As String, _
                    ByVal sProcess As String, ByVal ssItem As String, _
                    ByVal Fields As System.Xml.XmlNodeList, ByVal sModule As String) As Integer
        Dim sLabel As String
        Dim Width As Integer
        Dim sHelp As String
        Dim s As String
        Dim sOut As String
        Dim tc As TableColumn
        Dim Process As Boolean
        Dim b As Boolean = False

        sOut = Header("")
        sOut &= vbCrLf
        sOut &= "execute dbo.shlDialogFormInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@Title = 'Amend " & ssItem & "'" & vbCrLf
        sOut &= "   ,@TitleParameters = '"
        s = ""
        For Each tc In tDefn
            If tc.Primary Then
                sOut &= s & tc.Name
                s = "||"
            End If
        Next
        sOut &= "'" & vbCrLf
        sOut &= "   ,@HelpPage = '" & sObject & ".html'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlProcessesInsert" & vbCrLf
        sOut &= "    @ProcessName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ModuleID = '" & LCase(sModule) & "'" & vbCrLf
        sOut &= "   ,@ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        For Each tc In tDefn
            If tc.Name = "AuditID" Or tc.Name = "State" Then
                Continue For
            End If

            sLabel = ""
            Width = -1
            sHelp = ""
            Process = False
            For Each n As System.Xml.XmlNode In Fields
                For Each att As System.Xml.XmlAttribute In n.Attributes
                    Select Case LCase(att.Name)
                        Case "name"
                            If LCase(tc.Name) <> LCase(att.Value) Then
                                Exit For
                            End If
                            Process = True
                        Case "label"
                            sLabel = att.Value
                            If sLabel = tc.Name Then sLabel = ""
                        Case "width"
                            Width = CInt(att.Value)
                        Case "helptext"
                            sHelp = att.Value
                    End Select
                Next
                If Process Then Exit For
            Next

            If Process Then
                sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
                sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                sOut &= "   ,@FieldName = '" & tc.Name & "'" & vbCrLf
                If sLabel <> "" Then
                    sOut &= "   ,@Label = '" & sLabel & "'" & vbCrLf
                End If
                sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbCrLf
                If tc.vbType = "string" Then
                    sOut &= "   ,@Width = " & tc.Length & vbCrLf
                End If
                If Width > 0 Then
                    sOut &= "   ,@DisplayWidth = " & Width & vbCrLf
                Else
                    sOut &= "   ,@DisplayWidth = " & GetFieldWidth(tc.Length) & vbCrLf
                End If
                If tc.Primary Then
                    sOut &= "   ,@DisplayType = 'L'" & vbCrLf
                    sOut &= "   ,@IsPrimary = 'Y'" & vbCrLf
                Else
                    sOut &= "   ,@Enabled = 'Y'" & vbCrLf
                    If tc.Nullable = "N" Then
                        sOut &= "   ,@Required = 'Y'" & vbCrLf
                    End If
                    If sHelp <> "" Then
                        sOut &= "   ,@HelpText = '" & sHelp & "'" & vbCrLf
                    Else
                        sOut &= "   ,@HelpText = 'Place help text here!'" & vbCrLf
                    End If
                End If
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            Else
                If tc.Primary Then
                    sOut &= "execute dbo.shlParametersInsert" & vbCrLf
                    sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                    sOut &= "   ,@ParameterName = '" & tc.Name & "'" & vbCrLf
                    sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbCrLf
                    If tc.vbType = "string" Then
                        sOut &= "   ,@Width = " & tc.Length & vbCrLf
                    End If
                    sOut &= "go" & vbCrLf
                    sOut &= vbCrLf
                End If
            End If
        Next
        If tDefn.hasAudit Then
            sOut &= "execute dbo.shlParametersInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ParameterName = 'AuditID'" & vbCrLf
            sOut &= "   ,@ValueType = 'integer'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        End If
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Okay'" & vbCrLf
        sOut &= "   ,@Process = '" & sProcess & "'" & vbCrLf
        sOut &= "   ,@RowBased = 'Y'" & vbCrLf
        sOut &= "   ,@Validate = 'Y'" & vbCrLf
        sOut &= "   ,@CloseObject = 'Y'" & vbCrLf
        sOut &= "   ,@ImageFile = 'okay.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Save changes and exit'" & vbCrLf
        sOut &= "   ,@KeyCode = 13" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Cancel'" & vbCrLf
        sOut &= "   ,@CloseObject = 'Y'" & vbCrLf
        sOut &= "   ,@ImageFile = 'cancel.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Exit without saving changes'" & vbCrLf
        sOut &= "   ,@KeyCode = 27" & vbCrLf
        sOut &= Footer(True)

        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function TableDisable(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sObject As String = ""
        Dim sModule As String = "public"
        Dim sProcName As String = ""
        Dim sProcess As String = ""
        Dim sSeekKey As String = ""
        Dim ssItem As String = ""
        Dim sOut As String
        Dim s As String
        Dim b As Boolean = False
        Dim w As String
        Dim tc As TableColumn
        Dim Comma As String

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "procname"
                    sProcName = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "seekkey"
                    sSeekKey = a.InnerText
                Case "item"
                    ssItem = a.InnerText
                Case "process"
                    sProcess = a.InnerText
            End Select
        Next
        If sTable = "" Then Return -1
        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If tDefn.State <> 2 Then Return -1
        If Not tDefn.hasState Then Return -1

        If sObject = "" Then sObject = sModule
        If sProcName = "" Then sProcName = sTable & "Disable"
        If sProcess = "" Then sProcess = sObject
        If ssItem = "" Then ssItem = sObject
        If sSeekKey = "" Then sSeekKey = sObject

        Dim sp As New StoredProcedure(sProcName, sqllib)
        sp.ConfigName = sObject
        sp.ModuleName = sModule
        sp.ProcessName = sProcess

        If sp.State <> 2 Then

            sp.ProcedureName = sProcName

            sOut = "create procedure dbo." & sp.ProcedureName & vbCrLf
            Comma = " "
            For Each tc In tDefn
                If tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                    sOut &= "   " & Comma & "@" & tc.Name & " " & tc.TypeText
                    If tc.Nullable = "Y" Then
                        sOut &= " = null"
                    End If
                    sOut &= vbCrLf
                    Comma = ","
                    sp.AddParameter("@" & tc.Name, tc.Type, tc.Length, tc.Precision, tc.Scale, "in")
                End If
            Next
            If tDefn.hasAudit Then
                sOut &= "   ,@AuditID integer" & vbCrLf
                sp.AddParameter("@AuditID", "integer", 4, 0, 0, "in")
            End If
            sOut &= "as" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    set nocount on" & vbCrLf
            sOut &= "    declare @e integer" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "           ,@AudID integer" & vbCrLf
            End If
            sOut &= vbCrLf
            sOut &= "    set @e = 0" & vbCrLf
            sOut &= "    while @e = 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "        select  @AudID = a.AuditID" & vbCrLf
                sOut &= "        from    dbo." & tDefn.TableName & " a (holdlock)" & vbCrLf
                w = "        where   a."
                For Each s In tDefn.PKeys
                    tc = tDefn.Column(s)
                    sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                    w = "        and     a."
                Next
                sOut &= vbCrLf
                sOut &= "        if @@rowcount = 0" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            set @e = 51001" & vbCrLf
                sOut &= "            raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
                sOut &= vbCrLf
                sOut &= "        if @AudID <> @AuditID   -- already changed" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            set @AudID = -1" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
                sOut &= "        set @AudID = @AudID + 1" & vbCrLf
            Else
                sOut &= vbCrLf
                sOut &= "        if not exists" & vbCrLf
                sOut &= "        (" & vbCrLf
                sOut &= "            select  'a'" & vbCrLf
                sOut &= "            from    dbo." & tDefn.TableName & " a (holdlock)" & vbCrLf
                w = "            where   a."
                For Each s In tDefn.PKeys
                    tc = tDefn.Column(s)
                    sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                    w = "            and     a."
                Next
                sOut &= "        )" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            set @e = 51001" & vbCrLf
                sOut &= "            raiserror (@e, 16, 1, '" & ssItem & "')" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
            End If
            sOut &= vbCrLf
            sOut &= "        begin transaction" & vbCrLf
            sOut &= vbCrLf
            sOut &= "        update  dbo." & tDefn.TableName & vbCrLf
            sOut &= "        set     State = 'dl'" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "               ,AuditID = @AudID" & vbCrLf
            End If
            w = "        where   "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= w & tc.Name & " = @" & tc.Name & vbCrLf
                w = "        and     "
            Next
            sOut &= vbCrLf
            sOut &= "        set @e = @@error" & vbCrLf
            If tDefn.hasAudit Then
                sOut &= "        if @e <> 0" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            break" & vbCrLf
                sOut &= "        end" & vbCrLf
                sOut &= vbCrLf
                sOut &= "        execute @e = dbo." & tDefn.TableName & "AuditInsert" & vbCrLf
                Comma = " "
                For Each s In tDefn.PKeys
                    tc = tDefn.Column(s)
                    sOut &= "           " & Comma & "@" & tc.Name & " = @" & tc.Name & vbCrLf
                    Comma = ","
                Next
                sOut &= "           ,@AuditID = @AudID" & vbCrLf
                sOut &= "           ,@ActionType = 'D'" & vbCrLf
            End If
            sOut &= "        break" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "    if @e <> 0" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            rollback transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "    else" & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        if @@trancount > 0" & vbCrLf
            sOut &= "        begin" & vbCrLf
            sOut &= "            commit transaction" & vbCrLf
            sOut &= "        end" & vbCrLf
            sOut &= "        execute dbo." & tDefn.TableName & "Get" & "    -- return the changes" & vbCrLf
            Comma = " "
            For Each s In tDefn.PKeys
                tc = tDefn.Column(s)
                sOut &= "           " & Comma & "@p" & tc.Name & " = @" & tc.Name & vbCrLf
                Comma = ","
            Next
            If tDefn.hasAudit Then
                sOut &= vbCrLf
                sOut &= "        if @AudID = -1" & vbCrLf
                sOut &= "        begin" & vbCrLf
                sOut &= "            print 'This " & ssItem & " has changed since it was retrieved'" & vbCrLf
                sOut &= "        end" & vbCrLf
            End If
            sOut &= "    end" & vbCrLf
            sOut &= "    return @e" & vbCrLf
            sOut &= "end" & vbCrLf

            sp.ProcedureText = sOut
        End If

        sOut = Header("") & sp.FullText & Footer(True)
        PutFile("proc." & sProcName & ".sql", sOut)

        sp.Mode = "P"
        sp.ConfirmMsg = "Do you wish to disable this " & ssItem & "?"
        sp.SeekKey = sSeekKey
        sp.AddResult("@StateName", "varchar", 50, 0, 0)
        sp.AddResult("@State", "varchar", 2, 0, 0)
        sOut = Header("") & vbCrLf & sp.ConfigText & Footer(True)
        PutFile("config." & sObject & ".sql", sOut)

        Return 0
    End Function

    Private Function ConfigTableAudit(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sModule As String = "public"
        Dim ssItem As String = ""
        Dim sObject As String = ""
        Dim sProcess As String = ""
        Dim sOut As String
        Dim s As String
        Dim tc As TableColumn
        Dim b As Boolean = False

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "process"
                    sProcess = a.InnerText
                Case "item"
                    ssItem = a.InnerText
            End Select
        Next
        If sObject = "" Then sObject = sModule
        If sProcess = "" Then sProcess = sObject & "Get"
        If ssItem = "" Then ssItem = sObject

        Dim tDefn As New TableColumns(sTable, sqllib, True)
        If Not tDefn.hasAudit Then
            Return 0
        End If

        sOut = Header("")
        sOut &= vbCrLf
        sOut &= "execute dbo.shlGridFormInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@Title = '" & ssItem & " Change History'" & vbCrLf
        sOut &= "   ,@DataParameter = '" & sObject & "Get'" & vbCrLf
        sOut &= "   ,@ColourColumn = 'ActionType'" & vbCrLf
        sOut &= "   ,@TitleParameters = '"
        s = ""
        For Each tc In tDefn
            If tc.Primary Then
                sOut &= s & tc.Name
                s = "||"
            End If
        Next
        sOut &= "'" & vbCrLf
        sOut &= "   ,@HelpPage = '" & sObject & ".html'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlProcessesInsert" & vbCrLf
        sOut &= "    @ProcessName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ModuleID = '" & LCase(sModule) & "'" & vbCrLf
        sOut &= "   ,@ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        For Each tc In tDefn
            If tc.Primary Then
                sOut &= "execute dbo.shlParametersInsert" & vbCrLf
                sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                sOut &= "   ,@ParameterName = '" & tc.Name & "'" & vbCrLf
                sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbCrLf
                If tc.vbType = "string" Then
                    sOut &= "   ,@Width = " & tc.Length & vbCrLf
                End If
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            End If
        Next
        sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@FieldName = 'AuditID'" & vbCrLf
        sOut &= "   ,@Label = 'Sequence'" & vbCrLf
        sOut &= "   ,@ValueType = 'integer'" & vbCrLf
        sOut &= "   ,@DisplayWidth = 40" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@FieldName = 'ActionType'" & vbCrLf
        sOut &= "   ,@ValueType = 'string'" & vbCrLf
        sOut &= "   ,@Width = 1" & vbCrLf
        sOut &= "   ,@DisplayWidth = -1" & vbCrLf
        sOut &= "   ,@DisplayType = 'H'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@FieldName = 'Action'" & vbCrLf
        sOut &= "   ,@ValueType = 'string'" & vbCrLf
        sOut &= "   ,@Width = 50" & vbCrLf
        sOut &= "   ,@DisplayWidth = 60" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@FieldName = 'UserID'" & vbCrLf
        sOut &= "   ,@Label = 'Actioned By'" & vbCrLf
        sOut &= "   ,@ValueType = 'string'" & vbCrLf
        sOut &= "   ,@Width = 128" & vbCrLf
        sOut &= "   ,@DisplayWidth = 100" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@FieldName = 'AuditTime'" & vbCrLf
        sOut &= "   ,@Label = 'Actioned'" & vbCrLf
        sOut &= "   ,@ValueType = 'datetime'" & vbCrLf
        sOut &= "   ,@DisplayWidth = 80" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        For Each tc In tDefn
            If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
                sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                sOut &= "   ,@FieldName = '" & tc.Name & "'" & vbCrLf
                sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbCrLf
                If tc.vbType = "string" Then
                    sOut &= "   ,@Width = " & tc.Length & vbCrLf
                    sOut &= "   ,@DisplayWidth = " & (tc.Length * 5) & vbCrLf
                Else
                    sOut &= "   ,@DisplayWidth = 60" & vbCrLf
                End If
                sOut &= "go" & vbCrLf
                sOut &= vbCrLf
            End If
        Next
        If tDefn.hasState Then
            sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@FieldName = 'State'" & vbCrLf
            sOut &= "   ,@ValueType = 'string'" & vbCrLf
            sOut &= "   ,@Width = 50" & vbCrLf
            sOut &= "   ,@DisplayWidth = 50" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        End If
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@PropertyType = 'cl'" & vbCrLf
        sOut &= "   ,@PropertyName = 'I'" & vbCrLf
        sOut &= "   ,@Value = 'black'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@PropertyType = 'cl'" & vbCrLf
        sOut &= "   ,@PropertyName = 'D'" & vbCrLf
        sOut &= "   ,@Value = 'red'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@PropertyType = 'cb'" & vbCrLf
        sOut &= "   ,@PropertyName = 'D'" & vbCrLf
        sOut &= "   ,@Value = 'mistyrose'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Refresh'" & vbCrLf
        sOut &= "   ,@Process = '" & sProcess & "'" & vbCrLf
        sOut &= "   ,@ImageFile = 'refresh.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Refresh data'" & vbCrLf
        sOut &= "   ,@KeyCode = 120" & vbCrLf
        sOut &= Footer(True)

        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function ConfigTable(ByVal files As Xml.XmlElement) As Integer
        Dim sTable As String = ""
        Dim sObject As String = ""
        Dim sModule As String = "public"
        Dim sSeekKey As String = ""
        Dim sProcess As String = ""
        Dim ssItem As String = ""
        Dim sDescription As String = ""
        Dim Fields As System.Xml.XmlNodeList = files.ChildNodes
        Dim sOut As String
        Dim sLabel As String
        Dim Width As Integer
        Dim tc As TableColumn
        Dim Process As Boolean
        Dim b As Boolean

        For Each a As Xml.XmlAttribute In files.Attributes
            Select Case a.Name
                Case "table"
                    sTable = a.InnerText
                Case "objectname"
                    sObject = a.InnerText
                Case "process"
                    sProcess = a.InnerText
                Case "module"
                    sModule = a.InnerText
                Case "seekkey"
                    sSeekKey = a.InnerText
                Case "item"
                    ssItem = a.InnerText
                Case "description"
                    sDescription = a.InnerText
            End Select
        Next
        If sDescription = "" Then sDescription = sModule
        If sSeekKey = "" Then sSeekKey = sObject
        If ssItem = "" Then ssItem = sObject

        Dim tDefn As New TableColumns(sTable, sqllib, True)

        sOut = Header("")
        sOut &= vbCrLf
        sOut &= "execute dbo.shlGridFormInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@Title = '" & sDescription & "'" & vbCrLf
        sOut &= "   ,@DataParameter = '" & sProcess & "'" & vbCrLf
        If tDefn.hasState Then
            sOut &= "   ,@StateFilter = 'Y'" & vbCrLf
            sOut &= "   ,@ColourColumn = 'State'" & vbCrLf
        End If
        sOut &= "   ,@HelpPage = '" & sObject & ".html'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlProcessesInsert" & vbCrLf
        sOut &= "    @ProcessName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ModuleID = '" & LCase(sModule) & "'" & vbCrLf
        sOut &= "   ,@ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        If tDefn.hasAudit Then
            sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@FieldName = 'AuditID'" & vbCrLf
            sOut &= "   ,@DisplayWidth = -1" & vbCrLf
            sOut &= "   ,@ValueType = 'integer'" & vbCrLf
            sOut &= "   ,@DisplayType = 'H'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        End If
        For Each tc In tDefn
            If tc.Name = "AuditID" Or tc.Name = "State" Then
                Continue For
            End If

            sLabel = ""
            Width = -1
            Process = False
            For Each n As System.Xml.XmlNode In Fields
                For Each att As System.Xml.XmlAttribute In n.Attributes
                    Select Case LCase(att.Name)
                        Case "name"
                            If LCase(tc.Name) <> LCase(att.Value) Then
                                Exit For
                            End If
                            Process = True
                        Case "label"
                            sLabel = att.Value
                            If sLabel = tc.Name Then sLabel = ""
                        Case "width"
                            Width = CInt(att.Value)
                    End Select
                Next
                If Process Then
                    sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
                    sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                    sOut &= "   ,@FieldName = '" & tc.Name & "'" & vbCrLf
                    If sLabel <> "" Then
                        sOut &= "   ,@Label = '" & sLabel & "'" & vbCrLf
                    End If
                    sOut &= "   ,@ValueType = '" & tc.vbType & "'" & vbLf
                    If tc.vbType = "string" Then
                        sOut &= "   ,@Width = " & tc.Length & vbCrLf
                    End If
                    If Width > 0 Then
                        sOut &= "   ,@DisplayWidth = " & Width & vbCrLf
                    ElseIf Width = 0 Then
                        sOut &= "   ,@DisplayWidth = -1" & vbCrLf
                        sOut &= "   ,@DisplayType = 'H'" & vbCrLf
                    Else
                        sOut &= "   ,@DisplayWidth = " & GetFieldWidth(tc.Length) & vbCrLf
                    End If
                    If tc.Primary Then
                        sOut &= "   ,@IsPrimary = 'Y'" & vbCrLf
                    End If
                    sOut &= "go" & vbCrLf
                    sOut &= vbCrLf

                    Exit For
                End If
            Next
        Next
        If tDefn.hasState Then
            sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@FieldName = 'StateName'" & vbCrLf
            sOut &= "   ,@Label = 'State'" & vbCrLf
            sOut &= "   ,@ValueType = 'string'" & vbCrLf
            sOut &= "   ,@Width = 50" & vbCrLf
            sOut &= "   ,@DisplayWidth = 50" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlFieldParamInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@FieldName = 'State'" & vbCrLf
            sOut &= "   ,@ValueType = 'string'" & vbCrLf
            sOut &= "   ,@Width = 2" & vbCrLf
            sOut &= "   ,@DisplayWidth = -1" & vbCrLf
            sOut &= "   ,@DisplayType = 'H'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "---------------------------------------------------" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@PropertyType = 'cl'" & vbCrLf
            sOut &= "   ,@PropertyName = 'dl'" & vbCrLf
            sOut &= "   ,@Value = 'red'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@PropertyType = 'cb'" & vbCrLf
            sOut &= "   ,@PropertyName = 'dl'" & vbCrLf
            sOut &= "   ,@Value = 'mistyrose'" & vbCrLf
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
        Else
            sOut &= "---------------------------------------------------" & vbCrLf
            sOut &= vbCrLf
        End If
        sOut &= "execute dbo.shlPropertiesInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@PropertyType = 'lk'" & vbCrLf
        sOut &= "   ,@PropertyName = '" & UCase(sSeekKey) & "'" & vbCrLf
        sOut &= "   ,@Value = ''" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "---------------------------------------------------" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Refresh'" & vbCrLf
        sOut &= "   ,@Process = '" & sObject & "Get'" & vbCrLf
        sOut &= "   ,@ImageFile = 'refresh.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Refresh data'" & vbCrLf
        sOut &= "   ,@KeyCode = 120" & vbCrLf
        sOut &= "go" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute dbo.shlActionsInsert" & vbCrLf
        sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
        sOut &= "   ,@ActionName = 'Add'" & vbCrLf
        sOut &= "   ,@Process = '" & sObject & "Add'" & vbCrLf
        sOut &= "   ,@ImageFile = 'add.gif'" & vbCrLf
        sOut &= "   ,@ToolTip = 'Create new " & ssItem & "'" & vbCrLf

        b = False
        For Each tc In tDefn
            If Not tc.Primary And tc.Name <> "AuditID" And tc.Name <> "State" Then
                b = True
                Exit For
            End If
        Next
        If b Then
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlActionsInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ActionName = 'Update'" & vbCrLf
            sOut &= "   ,@Process = '" & sObject & "Edit'" & vbCrLf
            sOut &= "   ,@RowBased = 'Y'" & vbCrLf
            sOut &= "   ,@Validate = 'Y'" & vbCrLf
            sOut &= "   ,@ImageFile = 'edit.gif'" & vbCrLf
            sOut &= "   ,@ToolTip = 'Amend " & ssItem & "'" & vbCrLf
            If tDefn.hasState Then
                sOut &= vbCrLf
                sOut &= "execute dbo.shlActionRulesInsert" & vbCrLf
                sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
                sOut &= "   ,@ActionName = 'Update'" & vbCrLf
                sOut &= "   ,@RuleName = 'R1'" & vbCrLf
                sOut &= "   ,@FieldName = 'State'" & vbCrLf
                sOut &= "   ,@Value = 'ac'" & vbCrLf
            End If
        End If
        If tDefn.hasState Then
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlActionsInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ActionName = 'Disable'" & vbCrLf
            sOut &= "   ,@Process = '" & sObject & "Disable'" & vbCrLf
            sOut &= "   ,@RowBased = 'Y'" & vbCrLf
            sOut &= "   ,@Validate = 'Y'" & vbCrLf
            sOut &= "   ,@ImageFile = 'delete.gif'" & vbCrLf
            sOut &= "   ,@ToolTip = 'Disable " & ssItem & "'" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlActionRulesInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ActionName = 'Disable'" & vbCrLf
            sOut &= "   ,@RuleName = 'R1'" & vbCrLf
            sOut &= "   ,@FieldName = 'State'" & vbCrLf
            sOut &= "   ,@Value = 'ac'" & vbCrLf
        End If
        If tDefn.hasAudit Then
            sOut &= "go" & vbCrLf
            sOut &= vbCrLf
            sOut &= "execute dbo.shlActionsInsert" & vbCrLf
            sOut &= "    @ObjectName = '" & sObject & "'" & vbCrLf
            sOut &= "   ,@ActionName = 'History'" & vbCrLf
            sOut &= "   ,@Process = '" & sObject & "Audit'" & vbCrLf
            sOut &= "   ,@RowBased = 'Y'" & vbCrLf
            sOut &= "   ,@ImageFile = 'history.gif'" & vbCrLf
            sOut &= "   ,@ToolTip = 'Change history'" & vbCrLf
        End If
        sOut &= Footer(True)

        PutFile("config." & sObject & ".sql", sOut)
        Return 0
    End Function

    Private Function WriteSCL(ByVal sFileName As String) As Integer
        If PutFile(sFileName & ".scl", MainSCL & vbCrLf & "#end" & vbCrLf) Then
            Return 0
        Else
            Return -1
        End If
    End Function

#Region "private functions"
    Private Function GetFieldWidth(ByVal Width As Integer) As String
        Dim s As String = "200"
        If Width < 21 Then
            s = "50"
        ElseIf Width < 51 Then
            s = "80"
        ElseIf Width < 101 Then
            s = "150"
        ElseIf Width < 201 Then
            s = "150"
        End If
        Return s
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
            MainSCL &= sName & vbCrLf

            Return True
        End If
    End Function

    Private Function GetCommandParameter(ByRef sSwitch As String) As String
        Dim sCommand As String
        Dim sParameter As String
        Dim i As Integer

        sCommand = Microsoft.VisualBasic.Command()
        i = InStr(1, sCommand, sSwitch, CompareMethod.Text)
        sParameter = ""
        If i > 0 Then
            sParameter = LTrim(Mid(sCommand, i + 2))
            If Mid(sParameter, 1, 1) = "-" Then
                sParameter = ""
            ElseIf Mid(sParameter, 1, 1) = """" Then
                sParameter = Mid(sParameter, 2)
                i = InStr(1, sParameter, """", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            Else
                i = InStr(1, sParameter, " ", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            End If
        End If
        GetCommandParameter = sParameter
    End Function

    Private Function Header(ByVal sName As String) As String
        Dim s As String
        Dim sh As String

        If SystemName = "" Then
            SystemName = sqllib.ShellName()
        End If
        sh = Mid("------------------------------------------------------------------", 1, Len(SystemName) + 6)

        s = "print '" & sh & "'" & vbCrLf
        s &= "print '-- " & SystemName & " --'" & vbCrLf
        s &= "print '" & sh & "'" & vbCrLf
        s &= "set nocount on" & vbCrLf
        s &= "go" & vbCrLf
        If sName <> "" Then
            s &= "if object_id('dbo." & sName & "') is not null" & vbCrLf
            s &= "begin" & vbCrLf
            s &= "    drop procedure dbo." & sName & vbCrLf
            s &= "end" & vbCrLf
            s &= "go" & vbCrLf
            s &= "create procedure dbo." & sName & vbCrLf
        End If
        Return s
    End Function

    Private Function Footer(ByVal bGo As Boolean) As String
        Dim sOut As String = ""

        If bGo Then
            sOut = "go" & vbCrLf
        End If
        sOut &= vbCrLf
        sOut &= "print '.oOo.'" & vbCrLf
        sOut &= "go" & vbCrLf

        Return sOut
    End Function

    Private Function GetConnectString(ByVal key As String) As String
        Try
            Dim instance As New System.Configuration.AppSettingsReader
            Dim settings As System.Configuration.ConnectionStringSettingsCollection = _
                ConfigurationManager.ConnectionStrings
            Dim s As String = ""

            If key = "default" Or key = "" Then
                instance.GetValue("default", key.GetType).ToString()
                s = instance.GetValue("default", key.GetType).ToString()
            Else
                s = key
            End If

            If Not settings Is Nothing Then
                Return settings.Item(s).ConnectionString
            End If
            Return ""

        Catch ex As Exception
            Console.WriteLine("GetConnectString error:")
            Console.WriteLine(ex.ToString)
            Return Nothing
        End Try
    End Function
#End Region
End Module
