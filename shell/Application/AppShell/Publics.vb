Option Explicit On 
Option Strict On

Imports System.IO
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Specialized

Module Publics
    Private sSystems As String = ""
    Private SystemKey As String
    Private ImageKey As String = ""
    Private Variables As New Hashtable
    Private sImagePath As String = ""
    Private sHelpPath As String = ""
    Private Ab As System.Windows.Forms.Form
    Private Missing As Image
    Private PropParm As ShellParameters
    Private oIcon As System.Drawing.Icon
    Private bMDI As Boolean
    Private bInit As Boolean = True
    Private dBusiness As Date = Nothing
    Private oParent As Form

    Public Processes As New ProcessDefns
    Public Objects As New ObjectDefns
    Public Register As New ObjectRegisters
    Public Styles As New shellStyles

    Public Property MDIParent() As Form
        Get
            MDIParent = oParent
        End Get
        Set(ByVal value As Form)
            oParent = value
        End Set
    End Property

    Public ReadOnly Property BusinessDate() As Date
        Get
            BusinessDate = dBusiness
        End Get
    End Property

    Public ReadOnly Property IsMDI() As Boolean
        Get
            IsMDI = bMDI
        End Get
    End Property

    Public ReadOnly Property ShellIcon() As System.Drawing.Icon
        Get
            ShellIcon = oIcon
        End Get
    End Property

    Public Property inInit() As Boolean
        Get
            inInit = bInit
        End Get
        Set(ByVal value As Boolean)
            bInit = value
        End Set
    End Property

    Public Function InitialiseApp(ByRef objStartup As MainMenu) As Boolean
        Dim b As Boolean = False
        Dim i As Integer
        Dim s As String
        Dim sFile As String
        Dim sDefault As String = ""
        Dim ar As New System.Configuration.AppSettingsReader

        SystemKey = GetCommandParameter("-k")

        Dim names As String() = ConfigurationManager.AppSettings.AllKeys
        Dim appStgs As NameValueCollection = ConfigurationManager.AppSettings
        For i = 0 To appStgs.Count - 1
            Select Case LCase(names(i))
                Case "systems"
                    sSystems = appStgs(i)
                Case "imagepath"
                    sImagePath = appStgs(i)
                Case "helppath"
                    sHelpPath = appStgs(i)
                Case "mdi"
                    If LCase(appStgs(i)) = "y" Then
                        bMDI = True
                    End If
            End Select
        Next

        If sSystems <> "" Then
            getSystemFileParms()
        Else
            sDefault = GetString(ar.GetValue("default", SystemKey.GetType))
            If SystemKey = "" Or LCase(SystemKey) = "default" Then
                SystemKey = sDefault
            End If

            For i = 0 To appStgs.Count - 1
                Select Case LCase(names(i))
                    Case LCase(SystemKey) & "imagekey"
                        ImageKey = appStgs(i)
                    Case LCase(SystemKey) & "imagepath"
                        sImagePath = appStgs(i)
                    Case LCase(SystemKey) & "helppath"
                        sHelpPath = appStgs(i)
                End Select
            Next
        End If

        s = GetCommandParameter("-c")
        If s <> "" Then ImageKey = s
        s = GetCommandParameter("-i")
        If s <> "" Then sImagePath = s
        s = GetCommandParameter("-h")
        If s <> "" Then sHelpPath = s

        If ImageKey = "" Then
            ImageKey = SystemKey
        End If
        If sImagePath = "" Then
            sImagePath = "image"
        End If
        If sHelpPath = "" Then
            sHelpPath = "help"
        End If

        Missing = Image.FromStream(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("AppShell.missing.gif"))
        dBusiness = Today()
        About()
        Application.DoEvents()
        If oIcon Is Nothing Then
            sFile = GetImagePath(ImageKey & ".ico")
            Try
                Dim ico As New System.Drawing.Icon(sFile)
                oIcon = ico
            Catch
                oIcon = Ab.Icon
            End Try
        End If

        If GetSchema(objStartup) Then
            b = Init(objStartup)
        End If
        Return b
    End Function

    Public Function getSystemFileParms() As Boolean
        Dim dom As New Xml.XmlDocument
        Dim x As Xml.XmlElement
        Dim x1 As Xml.XmlElement
        Dim sDefault As String = ""

        dom = GetConfigXML(sSystems)
        If dom Is Nothing Then
            MessageOut("System configuration file, '" & sSystems & "' not found or invalid.", "C")
            Return False
        End If

        For Each x In dom.DocumentElement.ChildNodes
            If x.Name = "defaults" Then
                For Each x1 In x.ChildNodes
                    Select Case x1.Name
                        Case "default"
                            sDefault = GetAttribute(x1.Attributes, "value")
                            If SystemKey = "" Or LCase(SystemKey) = "default" Then
                                SystemKey = sDefault
                            End If
                        Case "imagepath"
                            sImagePath = GetAttribute(x1.Attributes, "value")
                        Case "helppath"
                            sHelpPath = GetAttribute(x1.Attributes, "value")
                    End Select
                Next
            End If
        Next
        If SystemKey = "" Then
            MessageOut("System key not defined.", "C")
            Return False
        End If

        For Each x In dom.DocumentElement.ChildNodes
            If x.Name = "systems" Then
                For Each x1 In x.ChildNodes
                    If GetAttribute(x1.Attributes, "name") = SystemKey Then
                        For Each a As Xml.XmlAttribute In x1.Attributes
                            Select Case LCase(a.Name)
                                Case "imagekey"
                                    ImageKey = a.InnerText
                                Case "imagepath"
                                    sImagePath = a.InnerText
                                Case "helppath"
                                    sHelpPath = a.InnerText
                                Case "mdi"
                                    If LCase(a.InnerText) = "y" Then
                                        bMDI = True
                                    End If
                            End Select
                        Next
                    End If
                Next
            End If
        Next
        Return True
    End Function

    Private Function GetConfigXML(ByVal sFile As String) As Xml.XmlDocument
        Dim dom As New Xml.XmlDocument

        Try
            dom.Load(sFile)
            If dom.DocumentElement.Name = "roboshell" Then
                Return dom
            End If
        Catch ex As Exception
            dom = Nothing
        End Try
        Return Nothing
    End Function

    Public Function GetAttribute(ByVal Attribs As Xml.XmlAttributeCollection, ByVal Name As String) As String
        Dim s As String = ""
        Dim x As Xml.XmlAttribute

        For Each x In Attribs
            If x.Name = Name Then
                Return x.InnerText
            End If
        Next
        Return s
    End Function

    Public Function GetCommandParameter(ByRef sSwitch As String, _
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

    Private Function GetSchema(ByRef objStartup As MainMenu) As Boolean
        Dim dt As DataTable
        Dim Version As Double = 1
        Dim f As Field
        Dim p As shellParameter
        Dim dr As DataRow
        Dim objd As ObjectDefn
        Dim Act As ActionDefns
        Dim iKeyCode As Integer
        Dim ad As ActionDefn
        Dim vd As ValidationDefn
        Dim i As Integer
        Dim s As String
        Dim ProcessObject As ShellObject

        Try
            Dim xob As New StoredProcDefn("shlShellGet")
            xob.SetProperty("procname", "shlShellGet")
            xob.SetProperty("connectkey", SystemKey)
            xob.SetProperty("mode", "D")
            xob.SetProperty("dataparameter", "Var||Proc||Obj||Prop||Parm||Act||ActR||Fld||Val||ValR||APR||Ver||Style")
            xob.Parms.Add("Var", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Proc", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Obj", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Prop", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Parm", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Act", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("ActR", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Fld", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Val", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("ValR", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("APR", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Ver", Nothing, DbType.Object, False, True, 0)
            xob.Parms.Add("Style", Nothing, DbType.Object, False, True, 0)
            ProcessObject = xob.Create()
            If Not ProcessObject Is Nothing Then
                ProcessObject.Update(Nothing)
                If Not ProcessObject.SuccessFlag Then
                    Return False
                End If

                s = ""
                For Each ms As ShellMessage In ProcessObject.Messages
                    If ms.Type = "E" Then
                        s &= ms.Message & vbCrLf
                    End If
                Next
                If s <> "" Then
                    MessageOut(s)
                    Return False
                End If
            End If

            dt = CType(ProcessObject.parms.Item("Ver").Value, DataTable)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Version = CDbl(dt.Rows(0)(0))
                End If
            End If

            dt = CType(ProcessObject.parms.Item("Var").Value, DataTable)
            For Each dr In dt.Rows        ' Variables
                s = LCase(GetString(dr("VariableID")))
                If s = "businessdate" Then
                    dBusiness = CDate(GetString(dr("VariableValue")))
                Else
                    Variables.Add(LCase(GetString(dr("VariableID"))), GetString(dr("VariableValue")))
                End If
            Next
            Application.DoEvents()

            dt = CType(ProcessObject.parms.Item("Proc").Value, DataTable)
            For Each dr In dt.Rows        ' Processes
                Processes.Add(GetString(dr.Item("ProcessName")), _
                              GetString(dr.Item("SuccessProcess")), _
                              GetString(dr.Item("FailProcess")), _
                              GetString(dr.Item("ConfirmMsg")), _
                              GetString(dr.Item("UpdateParent")), _
                              GetString(dr.Item("ObjectName")), _
                              (CType(dr.Item("LoadVariables"), String) = "Y"))
                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Obj").Value, DataTable)
            For Each dr In dt.Rows        ' Objects
                Select Case GetString(dr.Item("ObjectType"))
                    Case "StoredProc"
                        Dim ob As New StoredProcDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Grid"
                        Dim ob As New GridDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Dialog"
                        Dim ob As New DialogDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Tree"
                        Dim ob As New TreeDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "CallOut"
                        Dim ob As New CallOutDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "CallAsm"
                        Dim ob As New CallAsmDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Transform"
                        Dim ob As New TransformDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Report"
                        Dim ob As New ReportDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Monitor"
                        Dim ob As New MonitorDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "TableWrite"
                        Dim ob As New TableWriteDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "MailTo"
                        Dim ob As New MailToDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Menu"
                        Dim ob As New MenuDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "FileOpen"
                        Dim ob As New FileOpenDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                    Case "Directory"
                        Dim ob As New DirectoryDefn(GetString(dr.Item("ObjectName")))
                        Objects.Add(ob)
                End Select

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Prop").Value, DataTable)
            For Each dr In dt.Rows        ' Properties
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)
                If objd Is Nothing Then
                    MessageOut("Property: Error finding object: " & _
                                            GetString(dr.Item("ObjectName")))
                Else
                    If GetString(dr.Item("PropertyType")) = "df" Then
                        objd.SetProperty(GetString(dr.Item("PropertyName")), _
                                         dr.Item("Value"))
                    Else
                        objd.Properties.Add(GetString(dr.Item("PropertyName")), _
                                            GetString(dr.Item("PropertyType")), _
                                            GetString(dr.Item("UserSpecific")) = "Y", _
                                            dr.Item("Value"))
                    End If
                End If

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Parm").Value, DataTable)
            For Each dr In dt.Rows        ' Parameters
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)
                If objd Is Nothing Then
                    MessageOut("Error loading parameter " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        GetString(dr.Item("ParameterName")), "C")
                Else
                    Dim vt As System.Data.DbType
                    vt = GetValueType(GetString(dr.Item("ValueType")))

                    Dim objValue As Object
                    Select Case CType(dr.Item("Type"), String)
                        Case "U"
                            objValue = dr.Item("Value")
                        Case "C"
                            objValue = GetConnectString(GetString(dr.Item("Value")))
                        Case Else
                            objValue = dr.Item("Value")
                    End Select

                    p = objd.Parms.Add(GetString(dr.Item("ParameterName")), objValue, _
                                    vt, (CType(dr.Item("Input"), String) = "Y"), _
                                (CType(dr.Item("Output"), String) = "Y"), _
                                    CType(dr.Item("Width"), Int32))
                    If Version > 1.1 Then
                        s = GetString(dr.Item("Field"))
                        If s <> "" Then
                            p.Field = s
                        End If
                    End If
                End If

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Fld").Value, DataTable)
            For Each dr In dt.Rows        ' Fields
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)
                If objd Is Nothing Then
                    MessageOut("Error loading field " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        GetString(dr.Item("FieldName")), "C")
                Else
                    Dim vt As System.Data.DbType
                    vt = GetValueType(GetString(dr.Item("ValueType")))

                    Try
                        i = CType(dr.Item("DisplayHeight"), Int32)
                    Catch ex As Exception
                        i = 1
                    End Try

                    f = objd.Fields.Add(GetString(dr.Item("FieldName")), _
                                    GetString(dr.Item("Label")), _
                                    CType(dr.Item("DisplayType"), String), _
                                    CType(dr.Item("DisplayWidth"), Int32), _
                                    i, _
                                    GetString(dr.Item("Format")), _
                                    (CType(dr.Item("Primary"), String) = "Y"), _
                                    CType(dr.Item("Justify"), String), _
                                    (CType(dr.Item("Required"), String) = "Y"), _
                                    GetString(dr.Item("Locate")), _
                                    vt, _
                                    CType(dr.Item("Width"), Int32), _
                                    CType(dr.Item("LabelWidth"), Int32), _
                                    CType(dr.Item("Decimal"), Int32), _
                                    GetString(dr.Item("NullText")), _
                                    GetString(dr.Item("HelpText")))

                    f.FillProcess = GetString(dr.Item("FillProcess"))
                    f.TextField = GetString(dr.Item("TextField"))
                    f.ValueField = GetString(dr.Item("ValueField"))
                    Try
                        f.LinkColumn = GetString(dr.Item("LinkColumn"))
                    Catch ex As Exception
                    End Try
                    Try
                        f.LinkField = GetString(dr.Item("LinkField"))
                    Catch ex As Exception
                    End Try

                    If Version > 1.0 Then
                        s = GetString(dr.Item("Container"))
                        If s <> "" Then
                            f.Container = s
                        End If
                    End If
                End If
                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Act").Value, DataTable)
            For Each dr In dt.Rows        ' Actions
                s = GetString(dr.Item("ObjectName"))
                If (CType(dr.Item("IsKey"), String) = "Y") Then
                    iKeyCode = CType(dr.Item("KeyCode"), Int32)
                Else
                    iKeyCode = 0
                End If

                If s = "MainMenu" Then
                    Act = objStartup.Actions
                Else
                    objd = CType(Objects.Item(s), ObjectDefn)
                    If objd Is Nothing Then
                        Act = Nothing
                    Else
                        Act = objd.Actions
                    End If
                End If
                If Act Is Nothing Then
                    MessageOut("Error loading action " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        GetString(dr.Item("ActionName")), "C")
                Else
                    Act.Add(GetString(dr.Item("ActionName")), _
                            GetString(dr.Item("Process")), _
                           (GetString(dr.Item("Validate")) = "Y"), _
                           (GetString(dr.Item("RowBased")) = "Y"), _
                            GetString(dr.Item("CloseObject")), _
                           (GetString(dr.Item("IsDblClick")) = "Y"), _
                            GetString(dr.Item("ImageFile")), _
                            GetString(dr.Item("ToolTip")), _
                            GetString(dr.Item("MenuType")), _
                            GetString(dr.Item("MenuText")), _
                            GetString(dr.Item("Parent")), _
                            Nothing, iKeyCode, _
                            GetString(dr.Item("FieldName")), _
                            GetString(dr.Item("ProcessField")), _
                            GetString(dr.Item("LinkedParam")), _
                            GetString(dr.Item("ParamValue")))
                End If
                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("ActR").Value, DataTable)
            For Each dr In dt.Rows        ' Action Rules
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)

                If objd Is Nothing Then
                    MessageOut("Error loading action rule " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        RTrim(GetString(dr.Item("ActionName"))) & "." & _
                        GetString(dr.Item("RuleName")), "C")
                Else
                    ad = CType(objd.Actions.Item(GetString(dr.Item("ActionName"))), _
                               ActionDefn)
                    If ad Is Nothing Then
                        MessageOut("Error loading action rule " & _
                            RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                            RTrim(GetString(dr.Item("ActionName"))) & "." & _
                            GetString(dr.Item("RuleName")), "C")
                    Else

                        Dim vt As ActionRule.ValidationType
                        Select Case GetString(dr.Item("ValidationType"))
                            Case "EQ"
                                vt = ActionRule.ValidationType.EQ
                            Case "GE"
                                vt = ActionRule.ValidationType.GE
                            Case "GT"
                                vt = ActionRule.ValidationType.GT
                            Case "LE"
                                vt = ActionRule.ValidationType.LE
                            Case "LT"
                                vt = ActionRule.ValidationType.LT
                            Case "NE"
                                vt = ActionRule.ValidationType.NE
                            Case "NN"
                                vt = ActionRule.ValidationType.NN
                            Case "VL"
                                vt = ActionRule.ValidationType.VL
                            Case "NV"
                                vt = ActionRule.ValidationType.NV
                            Case Else
                                vt = ActionRule.ValidationType.EQ
                                MessageOut("Unknown parameter value type in database")
                        End Select

                        If ad.Rules Is Nothing Then
                            ad.Rules = New ActionRuleDefns
                        End If

                        ad.Rules.Add(CType(dr.Item("RuleID"), Integer), _
                                     GetString(dr.Item("RuleName")), _
                                     GetString(dr.Item("FieldName")), _
                                     vt, _
                                     dr.Item("Value"))
                    End If
                End If

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("APR").Value, DataTable)
            For Each dr In dt.Rows        ' Action Process Rules
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)

                If objd Is Nothing Then
                    MessageOut("Error loading action process rule " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        RTrim(GetString(dr.Item("ActionName"))), "C")
                Else
                    ad = CType(objd.Actions.Item(GetString(dr.Item("ActionName"))), _
                               ActionDefn)
                    If ad Is Nothing Then
                        MessageOut("Error loading action process rule " & _
                            RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                            RTrim(GetString(dr.Item("ActionName"))), "C")
                    Else
                        If ad.Processes Is Nothing Then
                            ad.Processes = New ActionProcessRuleDefns
                        End If

                        ad.Processes.Add(GetString(dr.Item("Process")), _
                                     dr.Item("Value"))
                    End If
                End If

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("Val").Value, DataTable)
            For Each dr In dt.Rows        ' Validations
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)
                If objd Is Nothing Then
                    MessageOut("Error loading validation " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        GetString(dr.Item("ValidationName")), "C")
                Else
                    Dim vt As ValidationDefn.ValidationType
                    Select Case GetString(dr.Item("ValidationType"))
                        Case "EQ"
                            vt = ValidationDefn.ValidationType.EQ
                        Case "GE"
                            vt = ValidationDefn.ValidationType.GE
                        Case "GT"
                            vt = ValidationDefn.ValidationType.GT
                        Case "LE"
                            vt = ValidationDefn.ValidationType.LE
                        Case "LT"
                            vt = ValidationDefn.ValidationType.LT
                        Case "NE"
                            vt = ValidationDefn.ValidationType.NE
                        Case Else
                            vt = ValidationDefn.ValidationType.EQ
                            MessageOut("Unknown parameter value type in database")
                    End Select

                    Dim vl As ValidationDefn.ValType
                    Select Case GetString(dr.Item("ValueType"))
                        Case "C"
                            vl = ValidationDefn.ValType.Constant
                        Case "P"
                            vl = ValidationDefn.ValType.Process
                        Case "F"
                            vl = ValidationDefn.ValType.Field
                        Case Else
                            vl = ValidationDefn.ValType.Constant
                            MessageOut("Unknown validation value type in database")
                    End Select

                    objd.Validations.Add( _
                        GetString(dr.Item("ValidationName")), _
                        GetString(dr.Item("FieldName")), _
                        GetString(dr.Item("Process")), _
                        vt, _
                        vl, _
                        dr.Item("Value"), _
                        GetString(dr.Item("Message")), _
                        GetString(dr.Item("ReturnParameter")))
                End If

                Application.DoEvents()
            Next

            dt = CType(ProcessObject.parms.Item("ValR").Value, DataTable)
            For Each dr In dt.Rows        ' ValidationRules
                objd = CType(Objects.Item(GetString(dr.Item("ObjectName"))), _
                                                                    ObjectDefn)
                If objd Is Nothing Then
                    MessageOut("Error loading validation " & _
                        RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                        GetString(dr.Item("ValidationName")), "C")
                Else

                    vd = CType(objd.Validations.Item(GetString(dr.Item("ValidationName"))), _
                                ValidationDefn)
                    If vd Is Nothing Then
                        MessageOut("Error loading validation rule " & _
                            RTrim(GetString(dr.Item("ObjectName"))) & "." & _
                            RTrim(GetString(dr.Item("ValidationName"))) & "." & _
                            GetString(dr.Item("FieldNameName")), "C")
                    Else
                        If vd.AssociatedFields Is Nothing Then
                            i = 0
                        Else
                            i = vd.AssociatedFields.GetUpperBound(0) + 1
                        End If
                        ReDim Preserve vd.AssociatedFields(i)
                        vd.AssociatedFields(i) = GetString(dr.Item("FieldName"))
                    End If
                End If

                Application.DoEvents()
            Next

            If Version > 1.2 Then
                dt = CType(ProcessObject.parms.Item("Style").Value, DataTable)
                For Each dr In dt.Rows        ' Styles
                    Styles.Add(GetString(dr.Item("StyleID")), _
                               GetString(dr.Item("RowForeColor")), _
                               GetString(dr.Item("RowBackColor")), _
                               GetString(dr.Item("SelForeColor")), _
                               GetString(dr.Item("SelBackColor")))
                    Application.DoEvents()
                Next
            End If
            Return True

        Catch ex As Exception
            MessageOut(ex.Message, "C")
            Return False
        End Try
    End Function

    Public Function GetVariable(ByVal VariableID As String) As String
        Dim s As String
        Try
            s = Variables.Item(LCase(VariableID)).ToString
        Catch ex As Exception
            s = ""
        End Try
        Return s
    End Function

    Public Function GetVars() As Boolean
        Dim i As Integer = 9
    End Function

    Private Function Init(ByVal objStartup As Object) As Boolean
        Try
            ' set global values from user profile.

            Dim td As ObjectDefn = CType(Objects.Item("MainMenu"), ObjectDefn)
            If Not td Is Nothing Then
                Dim p As ShellProperty = td.Properties.Item("IsMDI", "u")
                If Not p Is Nothing Then
                    If UCase(GetString(p.Value)) = "Y" Then
                        bMDI = True
                    Else
                        bMDI = False
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            MessageOut(ex.Message, "C")
            Return False
        End Try
    End Function

    Public Sub RaiseHelp(ByVal bShift As Boolean, ByVal sPage As String)
        Dim s As String
        If bShift Then
            Dim myProcess As New Process
            If sPage = "" Or sPage = "MainMenu" Then
                s = Path.Combine(sHelpPath, "MainMenu.html")
            Else
                s = Path.Combine(sHelpPath, sPage)
            End If
            myProcess.StartInfo.FileName = s
            Try
                myProcess.Start()
            Catch ex As Exception
                About()
            End Try
        Else
            About()
        End If
    End Sub

    Public Sub About()
        Dim lbl As Label
        Dim fv As System.Diagnostics.FileVersionInfo
        Dim w As Integer
        Dim h As Integer

        If Ab Is Nothing Then
            Ab = New System.Windows.Forms.Form

            Ab.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Ab.ControlBox = False
            Ab.MaximizeBox = False
            Ab.MinimizeBox = False
            Ab.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Ab.ShowInTaskbar = False
            Ab.KeyPreview = True

            Ab.BackgroundImage = GetImage(ImageKey & ".about.bmp")
            Ab.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center

            w = Ab.BackgroundImage.Width
            h = Ab.BackgroundImage.Height
            Ab.ClientSize = New System.Drawing.Size(w, h)

            lbl = New Label
            lbl.BackColor = System.Drawing.Color.Transparent
            lbl.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            lbl.ForeColor = System.Drawing.Color.DarkGreen
            lbl.Top = h - 28
            lbl.Left = 10
            lbl.Name = "lblVersion"
            lbl.Size = New System.Drawing.Size(150, 18)
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            lbl.Visible = True
            Ab.Controls.Add(lbl)

            lbl = New Label
            lbl.BackColor = System.Drawing.Color.Transparent
            lbl.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            lbl.ForeColor = System.Drawing.Color.DarkGreen
            lbl.Top = h - 28
            lbl.Left = w - 310
            lbl.Name = "Copyright"
            lbl.RightToLeft = System.Windows.Forms.RightToLeft.No
            lbl.Size = New System.Drawing.Size(300, 18)

            FV = System.Diagnostics.FileVersionInfo.GetVersionInfo( _
                    System.Reflection.Assembly.GetExecutingAssembly.Location)
            lbl.Text = "Copyright: " & Replace(fv.LegalCopyright, "&", "&&")
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            lbl.Visible = True
            Ab.Controls.Add(lbl)

            AddHandler Ab.Click, AddressOf ab_Click
            AddHandler Ab.KeyPress, AddressOf ab_Press
        End If
        If Not bInit Then
            lbl = CType(Ab.Controls.Item(0), Label)
            lbl.Text = "Version " & Publics.GetVariable("Release")
            lbl.Visible = True
        End If
        Ab.Show()
    End Sub

    Private Sub ab_Press(ByVal Sender As System.Object, _
                        ByVal e As System.Windows.Forms.KeyPressEventArgs)
        AboutClose()
    End Sub

    Private Sub ab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AboutClose()
    End Sub

    Public Sub AboutClose()
        Try
            Ab.Hide()
        Catch ex As Exception
        End Try
    End Sub

    Public Function GetConnectString(ByVal sName As String) As String
        Dim s As String
        Dim dom As New Xml.XmlDocument
        Dim x As Xml.XmlElement
        Dim x1 As Xml.XmlElement

        If LCase(sName) = "default" Or sName = "" Then
            s = SystemKey
        Else
            s = sName
        End If

        dom = GetConfigXML(sSystems)
        If Not dom Is Nothing Then
            For Each x In dom.DocumentElement.ChildNodes
                If x.Name = "connections" Then
                    For Each x1 In x.ChildNodes
                        If GetAttribute(x1.Attributes, "name") = s Then
                            Return GetAttribute(x1.Attributes, "connect")
                        End If
                    Next
                End If
            Next
        Else
            Dim settings As System.Configuration.ConnectionStringSettingsCollection = _
                ConfigurationManager.ConnectionStrings

            If Not settings Is Nothing Then
                Return settings.Item(s).ConnectionString
            End If
        End If

        Return ""
    End Function

    Public Function GetString(ByVal objValue As Object) As String
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

    Public Function GetValueType(ByVal Name As String) As System.Data.DbType

        Dim vt As System.Data.DbType
        Select Case LCase(Name)
            Case "currency"
                vt = DbType.Currency
            Case "date"
                vt = DbType.Date
            Case "datetime"
                vt = DbType.DateTime
            Case "double"
                vt = DbType.Double
            Case "numeric", "decimal"
                vt = DbType.Decimal
            Case "int32", "integer"
                vt = DbType.Int32
            Case "int64"
                vt = DbType.Int64
            Case "object"
                vt = DbType.Object
            Case "string"
                vt = DbType.String
            Case Else
                vt = DbType.String
                MessageOut("Unknown parameter value type in database")
        End Select

        Return vt

    End Function

    Public Sub SetFormPosition(ByRef f As Form, ByVal sObjectName As String)
        Dim td As ObjectDefn = CType(Objects.Item(sObjectName), ObjectDefn)
        If td Is Nothing Then
            Exit Sub
        End If
        Dim p As ShellProperty = td.Properties.Item("Position", "u")
        If p Is Nothing Then
            Exit Sub
        End If
        Dim i As Integer
        Dim j As Integer
        Dim a() As String

        a = Split(GetString(p.Value), "||", 4, CompareMethod.Binary)
        With f
            i = CType(a(0), Integer)
            j = CType(a(1), Integer)
            If Screen.PrimaryScreen.WorkingArea.Height > i Then
                If i + j > 0 Then
                    .Top = i
                End If
            End If
            .Height = j
            i = CType(a(2), Integer)
            j = CType(a(3), Integer)
            If Screen.PrimaryScreen.WorkingArea.Width > i Then
                If i + j > 0 Then
                    .Left = i
                End If
            End If
            .Width = CType(a(3), Integer)
        End With
    End Sub

    Public Sub SaveFormPosition(ByRef f As Form, ByVal sObjectName As String)
        Dim s As String

        Try
            If f.WindowState = FormWindowState.Normal Then
                With f
                    s = .Top & "||" & .Height & "||" & .Left & "||" & .Width
                End With
                SaveProperty(sObjectName, "Position", "u", s)
            End If
        Catch ex As Exception
            MessageOut(ex.Message, "C")
        End Try
    End Sub

    Public Sub SaveProperty(ByVal ObjectName As String, _
                    ByVal PropertyName As String, ByVal PropertyType As String, ByVal Value As String)

        Try
            Dim td As ObjectDefn = CType(Objects.Item(ObjectName), ObjectDefn)
            If td Is Nothing Then
                MessageOut("Invalid object parameter specified (SaveProperty)")
                Exit Sub
            End If

            Dim p As ShellProperty = td.Properties.Item(PropertyName, "u")

            If p Is Nothing Then
                td.Properties.Add(PropertyName, "u", True, Value)
            Else
                p.Value = Value
            End If

            If PropParm Is Nothing Then
                Dim xd As ObjectDefn = CType(Objects.Item("shlUserPropertyAlter"), ObjectDefn)
                PropParm = New ShellParameters
                xd.Parms.Clone(PropParm)
            End If

            PropParm.Item("ObjectName").Value = ObjectName
            PropParm.Item("PropertyName").Value = PropertyName
            PropParm.Item("PropertyType").Value = PropertyType
            PropParm.Item("Value").Value = Value
            Dim pr As New ShellProcess("shluserpropertyalter", Nothing, PropParm)

        Catch ex As Exception
            MessageOut(ex.Message, "C")
        End Try
    End Sub

    Public Function ValueFormat(ByVal dec As Integer) As String
        Dim s As String = "#,##0"
        If dec > 0 Then
            s &= Mid(".0000000000000000000", 1, dec + 1)
        End If
        Return s
    End Function

    Public Function Round(ByVal Value As Double, ByVal Decs As Integer) As Double
        Return Math.Round(Value + (0.1 * 10 ^ (Decs * -1)), Decs)
    End Function

    Public Function Round(ByVal Value As Decimal, ByVal Decs As Integer) As Decimal
        Return CType(Math.Round(Value + (0.1 * 10 ^ (Decs * -1)), Decs), Decimal)
    End Function

    Public Sub MessageOut(ByVal sMessage As String, Optional ByVal sType As String = "")
        Select Case sType
            Case "C"
                MsgBox(sMessage, MsgBoxStyle.Critical)
            Case "E"
                MsgBox(sMessage, MsgBoxStyle.Exclamation)
            Case Else
                MsgBox(sMessage, MsgBoxStyle.Information)
        End Select

    End Sub

    Public Function GetImage(ByVal sFile As String) As Image
        Dim s As String
        Dim img As Image

        s = GetImagePath(sFile)

        If s = "" Then
            Return Nothing
        End If
        Try
            img = Image.FromFile(s)
        Catch
            img = Missing
        End Try
        Return img
    End Function

    Private Function GetImagePath(ByVal sFile As String) As String
        Dim s As String
        If sFile = "" Then
            s = ""
        ElseIf sImagePath <> "" Then
            s = System.IO.Path.Combine(sImagePath, sFile)
        Else
            s = sFile
        End If
        Return s
    End Function

    Public Function GetBackColour() As System.Drawing.Color
        If LCase(Publics.GetVariable("Production")) = "y" Then
            Return System.Drawing.Color.GhostWhite
        Else
            Return System.Drawing.Color.MistyRose
        End If
    End Function
End Module
