Option Explicit On
Option Strict On

'sp_add_jobstep [ @job_id = ] job_id | [ @job_name = ] 'job_name' 
'     [ , [ @step_id = ] step_id ] 
'     { , [ @step_name = ] 'step_name' } 
'     [ , [ @subsystem = ] 'subsystem' ] 
'     [ , [ @command = ] 'command' ] 
'     [ , [ @additional_parameters = ] 'parameters' ] 
'     [ , [ @cmdexec_success_code = ] code ] 
'     [ , [ @on_success_action = ] success_action ] 
'     [ , [ @on_success_step_id = ] success_step_id ] 
'     [ , [ @on_fail_action = ] fail_action ] 
'     [ , [ @on_fail_step_id = ] fail_step_id ] 
'     [ , [ @server = ] 'server' ] 
'     [ , [ @database_name = ] 'database' ] 
'     [ , [ @database_user_name = ] 'user' ] 
'     [ , [ @retry_attempts = ] retry_attempts ] 
'     [ , [ @retry_interval = ] retry_interval ] 
'     [ , [ @os_run_priority = ] run_priority ] 
'     [ , [ @output_file_name = ] 'file_name' ] 
'     [ , [ @flags = ] flags ] 
'     [ , { [ @proxy_id = ] proxy_id 
'         | [ @proxy_name = ] 'proxy_name' } ]

Imports System
Imports System.Collections

Public Enum JobStepAction
    ExitOK = 1
    ExitFail = 2
    GoToNext = 3
    GoToStep = 4
End Enum

Public Enum JobStepFlag
    OverwriteOutput = 0
    AppendToOutput = 2
    WriteStepHistory = 4
    OverwriteTable = 8
    AppendTable = 16
End Enum

Public Class JobStep
    Private slib As sql

    Private iID As Integer
    Private sName As String
    Private sSubSystem As String = "TSQL"
    Private sCommand As String
    Private sParameters As String = ""
    Private iExecSuccessCode As Integer = 0
    Private eOnSuccessAction As JobStepAction = JobStepAction.ExitOK
    Private iOnSuccessStepID As Integer = 0     ' used when eOnSuccessAction = GoToStep
    Private eOnFailAction As JobStepAction = JobStepAction.ExitFail
    Private iOnFailStepID As Integer = 0        ' used when eOnSuccessAction = GoToStep
    Private sServer As String = ""
    Private sDatabaseName As String = ""
    Private sDatabaseUserName As String = ""
    Private iRetryAttempts As Integer = 0
    Private iRetryInterval As Integer = 0
    Private iOSRunPriority As Integer = 0
    Private sOutputFileName As String = ""
    Private eFlags As JobStepFlag = JobStepFlag.OverwriteOutput
    Private sProxyName As String = ""
    Private iProxyEnabled As Integer = 0
    Private sProxyDescription As String = ""
    Private sProxyCredential As String = ""

#Region "Properties"
    Public Property ID() As Integer
        Get
            ID = iID
        End Get
        Set(ByVal iD As Integer)
            iID = iD
        End Set
    End Property

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal nm As String)
            sName = nm
        End Set
    End Property

    Public Property SubSystem() As String
        Get
            SubSystem = sSubSystem
        End Get
        Set(ByVal ss As String)

            Select Case UCase(ss)
                Case "ACTIVESCRIPTING", "CMDEXEC", "DISTRIBUTION", _
                     "SNAPSHOT", "LOGREADER", "MERGE", "QUEUEREADER", _
                     "ANALYSISQUERY", "ANALYSISCOMMAND", "DTS", "TSQL"
                    sSubSystem = ss

                Case Else
                    sSubSystem = "TSQL"

            End Select
        End Set
    End Property

    Public Property Command() As String
        Get
            Command = sCommand
        End Get
        Set(ByVal cmd As String)
            sCommand = cmd
        End Set
    End Property

    Public Property Parameters() As String
        Get
            Parameters = sParameters
        End Get
        Set(ByVal sp As String)
            sParameters = sp
        End Set
    End Property

    Public Property ExecSuccessCode() As Integer
        Get
            ExecSuccessCode = iExecSuccessCode
        End Get
        Set(ByVal esc As Integer)
            iExecSuccessCode = esc
        End Set
    End Property

    Public Property OnSuccessAction() As JobStepAction
        Get
            OnSuccessAction = eOnSuccessAction
        End Get
        Set(ByVal sa As JobStepAction)
            eOnSuccessAction = sa
        End Set
    End Property

    Public Property iOnSuccessAction() As Integer
        Get
            iOnSuccessAction = iGetJobStepAction(eOnSuccessAction)
        End Get
        Set(ByVal sa As Integer)
            eOnSuccessAction = Me.GetJobStepAction(sa, 1)
        End Set
    End Property

    Public Property OnSuccessStepID() As Integer
        Get
            OnSuccessStepID = iOnSuccessStepID
        End Get
        Set(ByVal os As Integer)
            iOnSuccessStepID = os
        End Set
    End Property

    Public Property OnFailAction() As JobStepAction
        Get
            OnFailAction = eOnFailAction
        End Get
        Set(ByVal ofa As JobStepAction)
            eOnFailAction = ofa
        End Set
    End Property

    Public Property iOnFailAction() As Integer
        Get
            iOnFailAction = iGetJobStepAction(eOnSuccessAction)
        End Get
        Set(ByVal sa As Integer)
            eOnFailAction = Me.GetJobStepAction(sa, 2)
        End Set
    End Property

    Public Property OnFailStepID() As Integer
        Get
            OnFailStepID = iOnFailStepID
        End Get
        Set(ByVal ofs As Integer)
            iOnFailStepID = ofs
        End Set
    End Property

    Public Property Server() As String
        Get
            Server = sServer
        End Get
        Set(ByVal sv As String)
            sServer = sv
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            DatabaseName = sDatabaseName
        End Get
        Set(ByVal db As String)
            sDatabaseName = db
        End Set
    End Property

    Public Property DatabaseUserName() As String
        Get
            DatabaseUserName = sDatabaseUserName
        End Get
        Set(ByVal dun As String)
            sDatabaseUserName = dun
        End Set
    End Property

    Public Property RetryAttempts() As Integer
        Get
            RetryAttempts = iRetryAttempts
        End Get
        Set(ByVal ra As Integer)
            iRetryAttempts = ra
        End Set
    End Property

    Public Property RetryInterval() As Integer
        Get
            RetryInterval = iRetryInterval
        End Get
        Set(ByVal ri As Integer)
            iRetryInterval = ri
        End Set
    End Property

    Public Property OSRunPriority() As Integer
        Get
            OSRunPriority = iOSRunPriority
        End Get
        Set(ByVal rp As Integer)
            iOSRunPriority = rp
        End Set
    End Property

    Public Property OutputFileName() As String
        Get
            OutputFileName = sOutputFileName
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Flags() As JobStepFlag
        Get
            Flags = eFlags
        End Get
        Set(ByVal jsf As JobStepFlag)
            eFlags = jsf
        End Set
    End Property

    Public Property iFlags() As Integer
        Get
            Select Case eFlags
                Case JobStepFlag.OverwriteOutput
                    iFlags = 0
                Case JobStepFlag.AppendToOutput
                    iFlags = 2
                Case JobStepFlag.WriteStepHistory
                    iFlags = 4
                Case JobStepFlag.OverwriteTable
                    iFlags = 8
                Case JobStepFlag.AppendTable
                    iFlags = 16
                Case Else
                    iFlags = 0
            End Select
        End Get
        Set(ByVal ijf As Integer)
            Select Case ijf
                Case 0
                    eFlags = JobStepFlag.OverwriteOutput
                Case 2
                    eFlags = JobStepFlag.AppendToOutput
                Case 4
                    eFlags = JobStepFlag.WriteStepHistory
                Case 8
                    eFlags = JobStepFlag.OverwriteTable
                Case 16
                    eFlags = JobStepFlag.AppendTable
                Case Else
                    eFlags = JobStepFlag.OverwriteOutput
            End Select
        End Set
    End Property

    Public Property ProxyName() As String
        Get
            ProxyName = sProxyName
        End Get
        Set(ByVal pn As String)
            sProxyName = pn
        End Set
    End Property

    Public Property ProxyEnabled() As Integer
        Get
            ProxyEnabled = iProxyEnabled
        End Get
        Set(ByVal pe As Integer)
            iProxyEnabled = pe
        End Set
    End Property

    Public Property ProxyDescription() As String
        Get
            ProxyDescription = sProxyDescription
        End Get
        Set(ByVal pd As String)
            sProxyDescription = pd
        End Set
    End Property

    Public Property ProxyCredential() As String
        Get
            ProxyCredential = sProxyCredential
        End Get
        Set(ByVal pc As String)
            sProxyCredential = pc
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal pID As Integer, ByVal pStepName As String, ByVal pSubSystem As String, ByVal pCommand As String, ByVal sqllib As sql)
        slib = sqllib
        iID = pID
        sName = pStepName
        sSubSystem = pSubSystem
        sCommand = pCommand
    End Sub

    Public Function XMLText(ByVal sTab As String) As String
        Dim sOut As String = ""
        Dim s As String = vbCrLf & sTab & "      "

        sOut &= sTab & "<step id='" & iID & "'" & vbCrLf
        sOut &= sTab & "      name='" & slib.GetXMLString(sName) & "'" & vbCrLf
        sOut &= sTab & "      subsystem='" & slib.GetXMLString(sSubSystem) & "'"
        If sParameters <> "" Then
            sOut &= s & "parameters='" & slib.GetXMLString(sParameters) & "'"
        End If
        If iExecSuccessCode <> 0 Then
            sOut &= s & "cmdsuccesscode='" & iExecSuccessCode & "'"
        End If
        If eOnSuccessAction <> JobStepAction.ExitOK Then
            sOut &= s & "successaction='" & eOnSuccessAction & "'"
        End If
        If iOnSuccessStepID <> 0 Then
            sOut &= s & "successstep='" & iOnSuccessStepID & "'"
        End If
        If eOnFailAction <> JobStepAction.ExitFail Then
            sOut &= s & "failaction='" & eOnFailAction & "'"
        End If
        If iOnFailStepID <> 0 Then
            sOut &= s & "failstep='" & iOnFailStepID & "'"
        End If
        If sServer <> "" Then
            sOut &= s & "server='" & slib.GetXMLString(sServer) & "'"
        End If
        If sDatabaseName <> "" Then
            sOut &= s & "database='" & slib.GetXMLString(sDatabaseName) & "'"
        End If
        If sDatabaseUserName <> "" Then
            sOut &= s & "databaseuser='" & slib.GetXMLString(sDatabaseUserName) & "'"
        End If
        If iRetryAttempts <> 0 Then
            sOut &= s & "retryattempts='" & iRetryAttempts & "'"
        End If
        If iRetryInterval <> 0 Then
            sOut &= s & "retryinterval='" & iRetryInterval & "'"
        End If
        If iOSRunPriority <> 0 Then
            sOut &= s & "osrunpriority='" & iOSRunPriority & "'"
        End If
        If sOutputFileName <> "" Then
            sOut &= s & "outputfile='" & slib.GetXMLString(sOutputFileName) & "'"
        End If
        sOut &= s & "flags='" & eFlags & "'"
        If sProxyName <> "" Then
            sOut &= s & "proxyname='" & slib.GetXMLString(sProxyName) & "'"
            sOut &= s & "proxyenabled='" & iProxyEnabled & "'"
            sOut &= s & "proxydescription='" & slib.GetXMLString(sProxyDescription) & "'"
            sOut &= s & "proxycredential='" & slib.GetXMLString(sProxyCredential) & "'"
        End If
        sOut &= ">" & vbCrLf & sTab & "  <![CDATA[" & sCommand & "]]>" & vbCrLf
        sOut &= sTab & "</step>" & vbCrLf

        Return sOut
    End Function
#End Region

    Private Function GetJobStepAction(ByVal i As Integer, ByVal iDefault As Integer) As JobStepAction
        If i < 1 Or i > 4 Then i = iDefault
        Select Case i
            Case 1
                Return JobStepAction.ExitOK
            Case 2
                Return JobStepAction.ExitFail
            Case 3
                Return JobStepAction.GoToNext
            Case 4
                Return JobStepAction.GoToStep
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function iGetJobStepAction(ByVal jsa As JobStepAction) As Integer
        Select Case jsa
            Case JobStepAction.ExitOK
                Return 1
            Case JobStepAction.ExitFail
                Return 2
            Case JobStepAction.GoToNext
                Return 3
            Case JobStepAction.GoToStep
                Return 4
            Case Else
                Return 0
        End Select
    End Function
End Class

Public Class JobSteps
    Inherits CollectionBase

    Private slib As sql

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal StepName As String) As JobStep
        Get
            For Each js As JobStep In Me
                If js.Name = StepName Then
                    Return js
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByRef sqllib As sql)
        slib = sqllib
    End Sub

    Public Function Add(ByVal js As JobStep) As Integer
        Return List.Add(js)
    End Function

    Public Function XMLText(ByVal sTab As String) As String
        Dim ss As String = ""
        Dim sOut As String = ""
        Dim cJS As JobStep

        For Each cJS In Me
            ss &= cJS.XMLText(sTab & "  ")
        Next
        If ss <> "" Then
            sOut &= sTab & "<steps>" & vbCrLf
            sOut &= ss
            sOut &= sTab & "</steps>" & vbCrLf
        End If
        Return sOut
    End Function
#End Region
End Class
