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

' sp_add_job [ @job_name = ] 'job_name'
'     [ , [ @enabled = ] enabled ] 
'     [ , [ @description = ] 'description' ] 
'     [ , [ @start_step_id = ] step_id ] 
'     [ , [ @category_name = ] 'category' ] 
'     [ , [ @category_id = ] category_id ] 
'     [ , [ @owner_login_name = ] 'login' ] 
'     [ , [ @notify_level_eventlog = ] eventlog_level ] 
'     [ , [ @notify_level_email = ] email_level ] 
'     [ , [ @notify_level_netsend = ] netsend_level ] 
'     [ , [ @notify_level_page = ] page_level ] 
'     [ , [ @notify_email_operator_name = ] 'email_name' ] 
'     [ , [ @notify_netsend_operator_name = ] 'netsend_name' ] 
'     [ , [ @notify_page_operator_name = ] 'page_name' ] 
'     [ , [ @delete_level = ] delete_level ] 
'     [ , [ @job_id = ] job_id OUTPUT ] 

Public Class Job
    Private slib As sql
    Private PreLoad As Integer = -1

    Public Enum JobLogLevel
        Never = 0
        OnSuccess = 1
        OnFailure = 2
        Always = 3
    End Enum

    Private sID As String = ""
    Private sName As String = ""
    Private iEnabled As Integer = 1
    Private sDescription As String
    Private iStartStepID As Integer = 1
    Private sCategory As String
    Private sCategoryType As String = "LOCAL"
    Private sOwner As String = "sa"
    Private eNotifyEventLog As JobLogLevel = JobLogLevel.OnFailure
    Private eNotifyEmail As JobLogLevel = JobLogLevel.Never
    Private eNotifyNetSend As JobLogLevel = JobLogLevel.Never
    Private eNotifyPage As JobLogLevel = JobLogLevel.Never
    Private sEmail As String = ""
    Private sNetSend As String = ""
    Private sPage As String = ""
    Private eDeleteLevel As JobLogLevel = JobLogLevel.Never

    Private sServers As ArrayList
    Private cJobSteps As JobSteps
    Private cJobSchedules As JobSchedules

#Region "Properties"

    Public Property JobName() As String
        Get
            JobName = sName
        End Get
        Set(ByVal jn As String)
            sName = jn
        End Set
    End Property

    Public Property Enabled() As Integer
        Get
            Enabled = iEnabled
        End Get
        Set(ByVal en As Integer)
            iEnabled = en
        End Set
    End Property

    Public Property Description() As String
        Get
            Description = sDescription
        End Get
        Set(ByVal dn As String)
            sDescription = dn
        End Set
    End Property

    Public Property StartStepID() As Integer
        Get
            StartStepID = iStartStepID
        End Get
        Set(ByVal ssi As Integer)
            iStartStepID = ssi
        End Set
    End Property

    Public Property Category() As String
        Get
            Category = sCategory
        End Get
        Set(ByVal cat As String)
            sCategory = cat
        End Set
    End Property

    Public Property CategoryType() As String
        Get
            CategoryType = sCategoryType
        End Get
        Set(ByVal ct As String)
            sCategoryType = ct
        End Set
    End Property

    Public Property Owner() As String
        Get
            Owner = sOwner
        End Get
        Set(ByVal ow As String)
            sOwner = ow
        End Set
    End Property

    Public Property NotifyEventLog() As JobLogLevel
        Get
            NotifyEventLog = eNotifyEventLog
        End Get
        Set(ByVal nel As JobLogLevel)
            eNotifyEventLog = nel
        End Set
    End Property

    Public Property NotifyEmail() As JobLogLevel
        Get
            NotifyEmail = eNotifyEmail
        End Get
        Set(ByVal ne As JobLogLevel)
            eNotifyEmail = ne
        End Set
    End Property

    Public Property NotifyNetSend() As JobLogLevel
        Get
            NotifyNetSend = eNotifyNetSend
        End Get
        Set(ByVal nns As JobLogLevel)
            eNotifyNetSend = nns
        End Set
    End Property

    Public Property NotifyPage() As JobLogLevel
        Get
            NotifyPage = eNotifyPage
        End Get
        Set(ByVal np As JobLogLevel)
            eNotifyPage = np
        End Set
    End Property

    Public Property Email() As String
        Get
            Email = sEmail
        End Get
        Set(ByVal eml As String)
            sEmail = eml
        End Set
    End Property

    Public Property NetSend() As String
        Get
            NetSend = sNetSend
        End Get
        Set(ByVal ns As String)
            sNetSend = ns
        End Set
    End Property

    Public Property Page() As String
        Get
            Page = sPage
        End Get
        Set(ByVal pg As String)
            sPage = pg
        End Set
    End Property

    Private Property DeleteLevel() As JobLogLevel
        Get
            DeleteLevel = eDeleteLevel
        End Get
        Set(ByVal dl As JobLogLevel)
            eDeleteLevel = dl
        End Set
    End Property

    Public ReadOnly Property StepSQL() As String
        Get
            Dim sOut As String = ""

            If PreLoad = 3 Then Return ""

            For Each js As JobStep In cJobSteps
                If js.SubSystem = "TSQL" Then
                    If js.Command <> "" Then
                        sOut &= "-- " & js.Server & " " & js.DatabaseName & vbCrLf
                        sOut &= "-- Step: " & js.ID & vbCrLf
                        sOut &= js.Command & vbCrLf
                        sOut &= "go" & vbCrLf
                        sOut &= vbCrLf
                    End If
                End If
            Next

            StepSQL = sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByRef sqllib As sql)
        PreLoad = 0
        slib = sqllib
    End Sub

    Public Sub New(ByVal JobID As String, ByRef sqllib As sql)
        Dim dr As DataRow
        Dim dt As DataTable
        Dim i As Integer
        Dim s As String
        Dim sSubSystem As String
        Dim sCommand As String

        PreLoad = 2
        slib = sqllib

        dr = sqllib.JobObject(JobID)
        If dr Is Nothing Then
            PreLoad = 3
            Return
        End If

        sName = slib.GetSQLString(dr("name"))
        iEnabled = slib.GetInteger(dr("enabled"), 0)
        sDescription = slib.GetSQLString(dr("description"))
        iStartStepID = slib.GetInteger(dr("start_step_id"), 0)
        sCategory = slib.GetSQLString(dr("category"))
        sCategoryType = slib.GetSQLString(dr("category_type"))
        sOwner = slib.GetSQLString(dr("owner"))
        eNotifyEventLog = GetJobLogLevel(dr("notify_level_eventlog"))
        eNotifyEmail = GetJobLogLevel(dr("notify_level_email"))
        eNotifyNetSend = GetJobLogLevel(dr("notify_level_netsend"))
        eNotifyPage = GetJobLogLevel(dr("notify_level_page"))
        sEmail = slib.GetSQLString(dr("email"))
        sNetSend = slib.GetSQLString(dr("netsend"))
        sPage = slib.GetSQLString(dr("page"))
        eDeleteLevel = GetJobLogLevel(dr("delete_level"))

        dt = sqllib.JobServer(JobID)
        sServers = New ArrayList
        For Each dr In dt.Rows
            sServers.Add(sqllib.GetString(dr("server")))
        Next

        cJobSteps = New JobSteps(sqllib)
        dt = sqllib.JobStep(JobID)
        For Each dr In dt.Rows
            i = slib.GetInteger(dr("step_id"), 0)
            s = slib.GetString(dr("step_name"))
            sSubSystem = slib.GetString(dr("subsystem"))
            sCommand = slib.GetString(dr("command"))

            Dim js As New JobStep(i, s, sSubSystem, sCommand, slib)

            js.Parameters = slib.GetString(dr("additional_parameters"))
            js.ExecSuccessCode = slib.GetInteger(dr("cmdexec_success_code"), 0)
            js.iOnSuccessAction = slib.GetInteger(dr("on_success_action"), 0)
            js.OnSuccessStepID = slib.GetInteger(dr("on_success_step_id"), 0)
            js.iOnFailAction = slib.GetInteger(dr("on_fail_action"), 0)
            js.OnFailStepID = slib.GetInteger(dr("on_fail_step_id"), 0)
            js.Server = slib.GetString(dr("server"))
            js.DatabaseName = slib.GetString(dr("database_name"))
            js.DatabaseUserName = slib.GetString(dr("database_user_name"))
            js.RetryAttempts = slib.GetInteger(dr("retry_attempts"), 0)
            js.RetryInterval = slib.GetInteger(dr("retry_interval"), 0)
            js.OSRunPriority = slib.GetInteger(dr("os_run_priority"), 0)
            js.OutputFileName = slib.GetString(dr("output_file_name"))
            js.iFlags = slib.GetInteger(dr("flags"), 0)
            js.ProxyName = slib.GetString(dr("proxy"))
            js.ProxyEnabled = slib.GetInteger(dr("proxyenabled"), 0)
            js.ProxyDescription = slib.GetString(dr("proxydescription"))
            js.ProxyCredential = slib.GetString(dr("proxycredential"))

            cJobSteps.Add(js)
        Next

        cJobSchedules = New JobSchedules(sqllib)
        dt = sqllib.JobSchedule(JobID)
        For Each dr In dt.Rows
            s = slib.GetString(dr("name"))
            i = slib.GetInteger(dr("freq_type"), 1)

            Dim js As New JobSchedule(s, i, slib)

            js.Enabled = slib.GetInteger(dr("enabled"), 1)
            js.FrequencyInterval = slib.GetInteger(dr("freq_interval"), 1)
            js.iFrequencySubdayType = slib.GetInteger(dr("freq_subday_type"), 0)
            js.FrequencySubdayInterval = slib.GetInteger(dr("freq_subday_interval"), 0)
            js.iFrequencyRelativeInterval = slib.GetInteger(dr("freq_relative_interval"), 0)
            js.FrequencyRecurrenceFactor = slib.GetInteger(dr("freq_recurrence_factor"), 0)
            js.ActiveStartDate = slib.GetInteger(dr("active_start_date"), 0)
            js.ActiveEndDate = slib.GetInteger(dr("active_end_date"), 0)
            js.ActiveStartTime = slib.GetInteger(dr("active_start_time"), 0)
            js.ActiveEndTime = slib.GetInteger(dr("active_end_time"), 0)
            js.OwnerLoginName = slib.GetString(dr("OwnerLogin"))

            cJobSchedules.Add(js)
        Next
    End Sub

    '<?xml version='1.0'?>
    '<sqldef>
    '  <job name='ReconcileEquityPositions'
    '       enabled='0'
    '       description='Reconcile Equity Positions SUMMIT -> EVE.'
    '       startstep='1'
    '       category='[Uncategorized (Local)]'
    '       categorytype='LOCAL'
    '       ownerlogin='sa'
    '       notifyeventlog='2'
    '       notifyemail='0'
    '       notifynetsend='0'
    '       notifypage='0'
    '       emailoperator=''
    '       netsendoperator=''
    '       pageoperator=''
    '       deletelevel='0'>
    '    <servers>
    '      <server>(local)</server>
    '    </server>
    '    <steps>
    '      <step id='1'
    '            name='main'
    '            subsystem='TSQL'
    '            database='roboshell'
    '            retryinterval='1'
    '            flags='0'
    '        <![CDATA['execute ReconcileEqPositions']]>
    '      </step>
    '    </steps>
    '    <schedules>
    '      <schedule name='main'
    '                enabled='1'
    '                freqtype='4'
    '                freqinterval='1'
    '                freqsubday='1'
    '                freqsubdayinterval='0'
    '                freqrelativeinterval='0'
    '                freqrecurrencefactor='0'
    '                activestartdate='20090423'
    '                activeenddate='99991231'
    '                activestarttime='230000'
    '                activeendtime='235959'
    '                ownerloginname='sa' />
    '    </schedules>
    '  </job>
    '</sqldef>

    Public Sub New(ByVal sqllib As sql, ByVal x As Xml.XmlElement)
        slib = sqllib
        Dim att As Xml.XmlAttribute
        Dim ele As Xml.XmlElement
        Dim ele0 As Xml.XmlElement

        Dim i As Integer
        Dim s As String
        Dim sSubSystem As String
        Dim sCommand As String

        If x.Name = "job" Then
            For Each att In x.Attributes
                Select Case att.Name
                    Case "name"
                        sName = att.InnerText

                    Case "enabled"
                        iEnabled = GetInteger(att.InnerText)

                    Case "description"
                        sDescription = att.InnerText

                    Case "startstep"
                        iStartStepID = GetInteger(att.InnerText)

                    Case "category"
                        sCategory = att.InnerText

                    Case "categorytype"
                        sCategoryType = att.InnerText

                    Case "ownerlogin"
                        sOwner = att.InnerText

                    Case "notifyeventlog"
                        eNotifyEventLog = GetJobLogLevel(att.InnerText)

                    Case "notifyemail"
                        eNotifyEmail = GetJobLogLevel(att.InnerText)

                    Case "notifynetsend"
                        eNotifyNetSend = GetJobLogLevel(att.InnerText)

                    Case "notifypage"
                        eNotifyPage = GetJobLogLevel(att.InnerText)

                    Case "emailoperator"
                        sEmail = att.InnerText

                    Case "netsendoperator"
                        sNetSend = att.InnerText

                    Case "pageoperator"
                        sPage = att.InnerText

                    Case "deletelevel"
                        eDeleteLevel = GetJobLogLevel(att.InnerText)
                End Select
            Next
        End If

        sServers = New ArrayList
        cJobSteps = New JobSteps(sqllib)
        cJobSchedules = New JobSchedules(sqllib)

        For Each ele In x.ChildNodes
            Select Case ele.Name
                Case "servers"
                    For Each ele0 In ele.ChildNodes
                        If ele0.Name = "server" Then
                            sServers.Add(ele0.InnerText)
                        End If
                    Next

                Case "steps"
                    For Each ele0 In ele.ChildNodes
                        If ele0.Name = "step" Then
                            i = 0
                            s = ""
                            sSubSystem = ""
                            sCommand = ele0.InnerText

                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "id"
                                        i = GetInteger(att.InnerText)

                                    Case "name"
                                        s = att.InnerText

                                    Case "subsystem"
                                        sSubSystem = att.InnerText
                                End Select
                            Next

                            Dim js As New JobStep(i, s, sSubSystem, sCommand, slib)

                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "parameters"
                                        js.Parameters = att.InnerText

                                    Case "cmdsuccesscode"
                                        js.ExecSuccessCode = GetInteger(att.InnerText)

                                    Case "successaction"
                                        js.iOnSuccessAction = GetInteger(att.InnerText)

                                    Case "successstep"
                                        js.OnSuccessStepID = GetInteger(att.InnerText)

                                    Case "failaction"
                                        js.iOnFailAction = GetInteger(att.InnerText)

                                    Case "failstep"
                                        js.OnFailStepID = GetInteger(att.InnerText)

                                    Case "server"
                                        js.Server = att.InnerText

                                    Case "database"
                                        js.DatabaseName = att.InnerText

                                    Case "databaseuser"
                                        js.DatabaseUserName = att.InnerText

                                    Case "retryattempts"
                                        js.RetryAttempts = GetInteger(att.InnerText)

                                    Case "retryinterval"
                                        js.RetryInterval = GetInteger(att.InnerText)

                                    Case "osrunpriority"
                                        js.OSRunPriority = GetInteger(att.InnerText)

                                    Case "outputfile"
                                        js.OutputFileName = att.InnerText

                                    Case "flags"
                                        js.iFlags = GetInteger(att.InnerText)

                                    Case "proxyname"
                                        js.ProxyName = att.InnerText

                                    Case "proxyenabled"
                                        js.ProxyEnabled = GetInteger(att.InnerText)

                                    Case "proxydescription"
                                        js.ProxyDescription = att.InnerText

                                    Case "proxycredential"
                                        js.ProxyCredential = att.InnerText
                                End Select
                            Next
                            cJobSteps.Add(js)
                        End If
                    Next

                Case "schedules"
                    For Each ele0 In ele.ChildNodes
                        i = 0
                        s = ""

                        If ele0.Name = "schedule" Then
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "name"
                                        s = att.InnerText

                                    Case "freqtype"
                                        i = GetInteger(att.InnerText)

                                End Select
                            Next

                            Dim js As New JobSchedule(s, i, slib)

                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "enabled"
                                        js.Enabled = GetInteger(att.InnerText)

                                    Case "freqinterval"
                                        js.FrequencyInterval = GetInteger(att.InnerText)

                                    Case "freqsubday"
                                        js.iFrequencySubdayType = GetInteger(att.InnerText)

                                    Case "freqsubdayinterval"
                                        js.FrequencySubdayInterval = GetInteger(att.InnerText)

                                    Case "freqrelativeinterval"
                                        js.iFrequencyRelativeInterval = GetInteger(att.InnerText)

                                    Case "freqrecurrencefactor"
                                        js.FrequencyRecurrenceFactor = GetInteger(att.InnerText)

                                    Case "activestartdate"
                                        js.ActiveStartDate = GetInteger(att.InnerText)

                                    Case "active_end_date"
                                        js.ActiveEndDate = GetInteger(att.InnerText)

                                    Case "activestarttime"
                                        js.ActiveStartTime = GetInteger(att.InnerText)

                                    Case "activeendtime"
                                        js.ActiveEndTime = GetInteger(att.InnerText)

                                    Case "ownerloginname"
                                        js.OwnerLoginName = att.InnerText
                                End Select
                            Next
                            cJobSchedules.Add(js)
                        End If
                    Next
            End Select
        Next
    End Sub

    Public Function JobText(ByVal opt As ScriptOptions) As String
        Dim sOut As String = ""

        If PreLoad = 3 Then Return ""

        sOut = "execute @rc = msdb.dbo.sp_add_category" & vbCrLf
        sOut &= "    @class  = N'JOB'" & vbCrLf
        sOut &= "   ,@type = N'" & sCategoryType & "'" & vbCrLf
        sOut &= "   ,@name = N'" & sCategory & "'" & vbCrLf
        sOut &= vbCrLf
        sOut &= "execute @rc = msdb.dbo.sp_add_job" & vbCrLf
        sOut &= "    @job_name = '" & sName & "'" & vbCrLf
        sOut &= "   ,@enabled = 0" & vbCrLf
        If eDeleteLevel <> JobLogLevel.Never Then
            sOut &= "   ,@delete_level = " & eDeleteLevel & vbCrLf
        End If
        sOut &= "   ,@description = '" & sDescription & "'" & vbCrLf
        sOut &= "   ,@category_name = '" & sCategory & "'" & vbCrLf
        sOut &= "   ,@owner_login_name = '" & sOwner & "'" & vbCrLf
        sOut &= "   ,@job_id = @JobID output" & vbCrLf
        sOut &= vbCrLf

        For Each js As JobStep In cJobSteps
            sOut &= "execute @rc = msdb.dbo.sp_add_jobstep" & vbCrLf
            sOut &= "    @job_id = @JobID" & vbCrLf
            sOut &= "   ,@step_id = " & js.ID & vbCrLf
            sOut &= "   ,@step_name = '" & js.Name & "'" & vbCrLf
            If js.SubSystem <> "TSQL" Then
                sOut &= "   ,@subsystem = '" & js.SubSystem & "'" & vbCrLf
            End If
            If js.Command <> "" Then
                sOut &= "   ,@command = '" & js.Command & "'" & vbCrLf
            End If
            If js.Parameters <> "" Then
                sOut &= "   ,@additional_parameters = '" & js.Parameters & "'" & vbCrLf
            End If
            If js.ExecSuccessCode <> 0 Then
                sOut &= "   ,@cmdexec_success_code = " & js.ExecSuccessCode & vbCrLf
            End If
            If js.iOnSuccessAction <> 1 Then
                sOut &= "   ,@on_success_action = " & js.iOnSuccessAction & vbCrLf
            End If
            If js.OnSuccessStepID <> 0 Then
                sOut &= "   ,@on_success_step_id = " & js.OnSuccessStepID & vbCrLf
            End If
            If js.iOnFailAction <> 2 Then
                sOut &= "   ,@on_fail_action = " & js.iOnFailAction & vbCrLf
            End If
            If js.OnFailStepID <> 0 Then
                sOut &= "   ,@on_fail_step_id = " & js.OnFailStepID & vbCrLf
            End If
            If js.Server <> "" Then
                sOut &= "   ,@server = '" & js.Server & "'" & vbCrLf
            End If
            If js.DatabaseName <> "" Then
                sOut &= "   ,@database_name = '" & js.DatabaseName & "'" & vbCrLf
            End If
            If js.DatabaseUserName <> "" Then
                sOut &= "   ,@database_user_name = '" & js.DatabaseUserName & "'" & vbCrLf
            End If
            If js.RetryAttempts <> 0 Then
                sOut &= "   ,@retry_attempts = " & js.RetryAttempts & vbCrLf
            End If
            If js.RetryInterval <> 0 Then
                sOut &= "   ,@retry_interval = " & js.RetryInterval & vbCrLf
            End If
            If js.OSRunPriority <> 0 Then
                sOut &= "   ,@os_run_priority = " & js.OSRunPriority & vbCrLf
            End If
            If js.OutputFileName <> "" Then
                sOut &= "   ,@output_file_name = '" & js.OutputFileName & "'" & vbCrLf
            End If
            If js.iFlags <> 0 Then
                sOut &= "   ,@flags = " & js.iFlags & vbCrLf
            End If
            sOut &= vbCrLf
        Next
        JobText = sOut
    End Function

    Public Function FullText(ByVal opt As ScriptOptions) As String
        Dim sOut As String = ""
        Dim s As String
        Dim i As Integer
        Dim b As Boolean

        If PreLoad = 3 Then Return ""

        sOut = "declare" & vbCrLf
        sOut &= "    @rc integer," & vbCrLf
        sOut &= "    @JobID binary(16)," & vbCrLf
        sOut &= "    @db sysname" & vbCrLf
        sOut &= vbCrLf
        sOut &= "set @rc = 0" & vbCrLf
        sOut &= "while @rc = 0" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    select  @JobID = job_id" & vbCrLf
        sOut &= "    from    msdb.dbo.sysjobs" & vbCrLf
        sOut &= "    where   name = '" & sName & "'" & vbCrLf
        sOut &= vbCrLf
        sOut &= "    if @@rowcount = 0" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        if not exists" & vbCrLf
        sOut &= "        (" & vbCrLf
        sOut &= "            select  'a'" & vbCrLf
        sOut &= "            from    msdb.dbo.syscategories" & vbCrLf
        sOut &= "            where   name = N'" & sCategory & "'" & vbCrLf
        sOut &= "            and     category_class = 1" & vbCrLf
        sOut &= "        )" & vbCrLf
        sOut &= "        begin" & vbCrLf
        sOut &= "            print 'creating job category ''" & sCategory & "''.'" & vbCrLf
        sOut &= "            execute @rc = msdb.dbo.sp_add_category" & vbCrLf
        sOut &= "                @class  = N'JOB'" & vbCrLf
        sOut &= "               ,@type = N'" & sCategoryType & "'" & vbCrLf
        sOut &= "               ,@name = N'" & sCategory & "'" & vbCrLf
        sOut &= vbCrLf
        sOut &= "            if @rc <> 0" & vbCrLf
        sOut &= "            begin" & vbCrLf
        sOut &= "                break" & vbCrLf
        sOut &= "            end" & vbCrLf
        sOut &= "        end" & vbCrLf
        sOut &= vbCrLf
        sOut &= "        print 'creating new job ''" & sName & "''.'" & vbCrLf
        sOut &= "        execute @rc = msdb.dbo.sp_add_job" & vbCrLf
        sOut &= "            @job_name = N'" & sName & "'" & vbCrLf
        sOut &= "           ,@enabled = 0" & vbCrLf
        If eNotifyEventLog <> JobLogLevel.OnFailure Then
            sOut &= "           ,@notify_level_eventlog = " & eNotifyEventLog & vbCrLf
        End If
        If eNotifyEmail <> JobLogLevel.Never Then
            sOut &= "           ,@notify_level_email = " & eNotifyEmail & vbCrLf
        End If
        If eNotifyNetSend <> JobLogLevel.Never Then
            sOut &= "           ,@notify_level_netsend = " & eNotifyNetSend & vbCrLf
        End If
        If eNotifyPage <> JobLogLevel.Never Then
            sOut &= "           ,@notify_level_page = " & eNotifyPage & vbCrLf
        End If
        If eDeleteLevel <> JobLogLevel.Never Then
            sOut &= "           ,@delete_level = " & eDeleteLevel & vbCrLf
        End If
        If sEmail <> "" Then
            sOut &= "           ,@notify_email_operator_name = '" & sEmail & "'" & vbCrLf
        End If
        If sNetSend <> "" Then
            sOut &= "           ,@notify_netsend_operator_name = '" & sNetSend & "'" & vbCrLf
        End If
        If sPage <> "" Then
            sOut &= "           ,@notify_page_operator_name = '" & sPage & "'" & vbCrLf
        End If
        sOut &= "           ,@description = N'" & sDescription & "'" & vbCrLf
        sOut &= "           ,@category_name = N'" & sCategory & "'" & vbCrLf
        sOut &= "           ,@owner_login_name = N'" & sOwner & "'" & vbCrLf
        sOut &= "           ,@job_id = @JobID output" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= vbCrLf
        sOut &= "	    execute @rc = msdb.dbo.sp_add_jobserver" & vbCrLf
        sOut &= "            @job_id = @JobID" & vbCrLf
        s = ""
        If sServers.Count > 0 Then
            s = sServers.Item(0).ToString
        End If
        If s = "" Then s = "(local)"
        sOut &= "           ,@server_name = N'" & s & "'" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= "    else" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        print 'updating job ''" & sName & "''.'" & vbCrLf
        sOut &= "        execute @rc = msdb.dbo.sp_delete_jobstep" & vbCrLf
        sOut &= "            @job_id = @JobID" & vbCrLf
        sOut &= "           ,@step_id = 0" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= vbCrLf

        b = True
        For Each js As JobStep In cJobSteps
            sOut &= "    execute @rc = msdb.dbo.sp_add_jobstep" & vbCrLf
            sOut &= "        @job_id = @JobID" & vbCrLf
            sOut &= "       ,@step_id = " & js.ID & vbCrLf
            sOut &= "       ,@step_name = N'" & js.Name & "'" & vbCrLf
            If js.SubSystem <> "TSQL" Then
                sOut &= "       ,@subsystem = N'" & js.SubSystem & "'" & vbCrLf
            End If
            If js.Command <> "" Then
                sOut &= "       ,@command = N'" & js.Command & "'" & vbCrLf
            End If
            If js.Parameters <> "" Then
                sOut &= "       ,@additional_parameters = N'" & js.Parameters & "'" & vbCrLf
            End If
            If js.ExecSuccessCode <> 0 Then
                sOut &= "       ,@cmdexec_success_code = " & js.ExecSuccessCode & vbCrLf
            End If
            If js.iOnSuccessAction <> 1 Then
                sOut &= "       ,@on_success_action = " & js.iOnSuccessAction & vbCrLf
            End If
            If js.OnSuccessStepID <> 0 Then
                sOut &= "       ,@on_success_step_id = " & js.OnSuccessStepID & vbCrLf
            End If
            If js.iOnFailAction <> 2 Then
                sOut &= "       ,@on_fail_action = " & js.iOnFailAction & vbCrLf
            End If
            If js.OnFailStepID <> 0 Then
                sOut &= "       ,@on_fail_step_id = " & js.OnFailStepID & vbCrLf
            End If
            If js.Server <> "" Then
                sOut &= "       ,@server = N'" & js.Server & "'" & vbCrLf
            End If
            If js.DatabaseName <> "" Then
                sOut &= "       ,@database_name = N'" & js.DatabaseName & "'" & vbCrLf
            End If
            If js.DatabaseUserName <> "" Then
                sOut &= "       ,@database_user_name = N'" & js.DatabaseUserName & "'" & vbCrLf
            End If
            If js.RetryAttempts <> 0 Then
                sOut &= "       ,@retry_attempts = " & js.RetryAttempts & vbCrLf
            End If
            If js.RetryInterval <> 0 Then
                sOut &= "       ,@retry_interval = " & js.RetryInterval & vbCrLf
            End If
            If js.OSRunPriority <> 0 Then
                sOut &= "       ,@os_run_priority = " & js.OSRunPriority & vbCrLf
            End If
            If js.OutputFileName <> "" Then
                sOut &= "       ,@output_file_name = N'" & js.OutputFileName & "'" & vbCrLf
            End If
            If js.iFlags <> 0 Then
                sOut &= "       ,@flags = " & js.iFlags & vbCrLf
            End If
            If js.ProxyName <> "" Then
                sOut &= "       ,@proxy_name = N'" & js.ProxyName & "'" & vbCrLf
            End If
            sOut &= "    if @rc <> 0 break" & vbCrLf
            sOut &= vbCrLf

            b = False
        Next

        If Not b Then
            sOut &= "    execute @rc = msdb.dbo.sp_update_job" & vbCrLf
            sOut &= "        @job_id = @jobId" & vbCrLf
            sOut &= "       ,@start_step_id = " & iStartStepID & vbCrLf
            sOut &= "    if @rc <> 0 break" & vbCrLf
            sOut &= vbCrLf
        End If

        b = True
        For Each js As JobSchedule In cJobSchedules
            If b Then
                b = False
                sOut &= "    if not exists" & vbCrLf
                sOut &= "    (" & vbCrLf
                sOut &= "        select  'a'" & vbCrLf
                sOut &= "        from    msdb.dbo.sysjobschedules" & vbCrLf
                sOut &= "        where   job_id = @JobID" & vbCrLf
                sOut &= "    )" & vbCrLf
                sOut &= "    begin" & vbCrLf
                sOut &= "        print 'creating schedule.'"
            End If
            sOut &= vbCrLf
            sOut &= "        execute @rc = msdb.dbo.sp_add_schedule" & vbCrLf
            sOut &= "            @schedule_name = N'" & js.Name & "'" & vbCrLf
            If js.Enabled = 0 Then
                sOut &= "           ,@enabled = 0" & vbCrLf
            End If
            i = js.FrequencyType
            If i <> 1 Then
                sOut &= "           ,@freq_type = " & i & vbCrLf
            End If
            i = js.FrequencyInterval
            If i <> 0 Then
                sOut &= "           ,@freq_interval = " & i & vbCrLf
            End If
            i = js.FrequencySubdayType
            If i <> 0 Then
                sOut &= "           ,@freq_subday_type = " & i & vbCrLf
            End If
            i = js.FrequencySubdayInterval
            If i <> 0 Then
                sOut &= "           ,@freq_subday_interval = " & i & vbCrLf
            End If
            i = js.FrequencyRelativeInterval
            If i <> 0 Then
                sOut &= "           ,@freq_relative_interval = " & i & vbCrLf
            End If
            i = js.FrequencyRecurrenceFactor
            If i <> 0 Then
                sOut &= "           ,@freq_recurrence_factor = " & i & vbCrLf
            End If
            i = js.ActiveStartDate
            If i > 19900101 And i <= 99991231 Then
                sOut &= "           ,@active_start_date = " & i & vbCrLf
            End If
            i = js.ActiveEndDate
            If i > 0 And i < 99991231 Then
                sOut &= "           ,@active_end_date = " & i & vbCrLf
            End If
            i = js.ActiveStartTime
            If i > 0 And i < 235960 Then
                sOut &= "           ,@active_start_time = " & i & vbCrLf
            End If
            i = js.ActiveEndTime
            If i > 0 And i < 235959 Then
                sOut &= "           ,@active_end_time = " & i & vbCrLf
            End If
            sOut &= "           ,@owner_login_name = N'" & js.OwnerLoginName & "'" & vbCrLf
            sOut &= "        if @rc <> 0 break" & vbCrLf
            sOut &= vbCrLf
            sOut &= "        print 'attaching schedule to job.'" & vbCrLf
            sOut &= "        execute @rc = msdb.dbo.sp_attach_schedule" & vbCrLf
            sOut &= "            @job_id = @JobID" & vbCrLf
            sOut &= "           ,@schedule_name = N'" & js.Name & "'" & vbCrLf
            sOut &= "        if @rc <> 0 break" & vbCrLf
        Next
        If Not b Then
            sOut &= "    end" & vbCrLf
        End If
        sOut &= "    break" & vbCrLf
        sOut &= "end" & vbCrLf

        FullText = sOut
    End Function

    Public Function XML(ByVal opt As ScriptOptions) As String
        Dim sOut As String
        Dim b As Boolean

        If PreLoad = 3 Then Return ""

        sOut = "<?xml version='1.0'?>" & vbCrLf
        sOut &= "<sqldef>" & vbCrLf
        sOut &= "  <job name='" & slib.GetXMLString(sName) & "'" & vbCrLf
        sOut &= "       enabled='" & iEnabled & "'" & vbCrLf
        sOut &= "       description='" & slib.GetXMLString(sDescription) & "'" & vbCrLf
        sOut &= "       startstep='" & iStartStepID & "'" & vbCrLf
        sOut &= "       category='" & slib.GetXMLString(sCategory) & "'" & vbCrLf
        sOut &= "       categorytype='" & slib.GetXMLString(sCategoryType) & "'" & vbCrLf
        sOut &= "       ownerlogin='" & slib.GetXMLString(sOwner) & "'" & vbCrLf
        sOut &= "       notifyeventlog='" & eNotifyEventLog & "'" & vbCrLf
        sOut &= "       notifyemail='" & eNotifyEmail & "'" & vbCrLf
        sOut &= "       notifynetsend='" & eNotifyNetSend & "'" & vbCrLf
        sOut &= "       notifypage='" & eNotifyPage & "'" & vbCrLf
        sOut &= "       emailoperator='" & slib.GetXMLString(sEmail) & "'" & vbCrLf
        sOut &= "       netsendoperator='" & slib.GetXMLString(sNetSend) & "'" & vbCrLf
        sOut &= "       pageoperator='" & slib.GetXMLString(sPage) & "'" & vbCrLf
        sOut &= "       deletelevel='" & eDeleteLevel & "'>" & vbCrLf

        sOut &= "    <servers>" & vbCrLf
        b = True
        For Each o As Object In sServers
            sOut &= "      <server>" & o.ToString & "</server>" & vbCrLf
            b = False
        Next
        If b Then
            sOut &= "      <server>(local)</server>" & vbCrLf
        End If
        sOut &= "    </servers>" & vbCrLf

        sOut &= cJobSteps.XMLText("    ")

        sOut &= cJobSchedules.XMLText("    ")

        sOut &= "  </job>" & vbCrLf
        sOut &= "</sqldef>" & vbCrLf
        XML = sOut
    End Function

    Private Function GetJobLogLevel(ByVal o As Object) As JobLogLevel
        Select Case slib.GetInteger(o, 0)
            Case 0
                Return JobLogLevel.Never
            Case 1
                Return JobLogLevel.OnSuccess
            Case 2
                Return JobLogLevel.OnFailure
            Case 3
                Return JobLogLevel.Always
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function GetInteger(ByVal Value As String) As Integer
        Try
            Return CInt(Value)
        Catch ex As Exception
            Return 0
        End Try
    End Function
#End Region
End Class
