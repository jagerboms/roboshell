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

Public Class Job
    Private dtJobs As DataTable
    Private dtServers As DataTable
    Private dtSteps As DataTable
    Private dtSchedules As DataTable

    Public ReadOnly Property JobName() As String
        Get
            Dim s As String = ""
            If Not dtJobs Is Nothing Then
                If dtJobs.Rows.Count > 0 Then
                    s = GetString(dtJobs.Rows(0)("name"))
                End If
            End If
            JobName = s
        End Get
    End Property

    Public Sub New(ByVal sJobID As String, ByVal sConnect As String)
        Dim psConn As SqlConnection

        psConn = New SqlConnection(sConnect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        dtJobs = GetJob(sJobID, psConn)
        dtServers = GetServers(sJobID, psConn)
        dtSteps = GetSteps(sJobID, psConn)
        dtSchedules = GetSchedules(sJobID, psConn)

        psConn.Close()
    End Sub

    Public Function FullText() As String
        Dim sOut As String = ""
        Dim dr As DataRow = Nothing
        Dim s As String
        Dim i As Integer
        Dim b As Boolean
        Dim sName As String = ""

        If Not dtJobs Is Nothing Then
            If dtJobs.Rows.Count > 0 Then
                dr = dtJobs.Rows(0)
            End If
        End If

        If dr Is Nothing Then Return ""
        sName = GetSQLString(dr("name"))

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
        sOut &= "        print 'creating new job ''" & sName & "''.'" & vbCrLf
        sOut &= "	    execute @rc = msdb.dbo.sp_add_job" & vbCrLf
        sOut &= "            @job_name = '" & sName & "'," & vbCrLf
        sOut &= "            @enabled = 0," & vbCrLf
        i = CInt(dr("notify_level_eventlog"))
        If i <> 2 Then
            sOut &= "            @notify_level_eventlog = " & i & "," & vbCrLf
        End If
        i = CInt(dr("notify_level_email"))
        If i <> 0 Then
            sOut &= "            @notify_level_email = " & i & "," & vbCrLf
        End If
        i = CInt(dr("notify_level_netsend"))
        If i <> 0 Then
            sOut &= "            @notify_level_netsend = " & i & "," & vbCrLf
        End If
        i = CInt(dr("notify_level_page"))
        If i <> 0 Then
            sOut &= "            @notify_level_page = " & i & "," & vbCrLf
        End If
        i = CInt(dr("delete_level"))
        If i <> 0 Then
            sOut &= "            @delete_level = " & i & "," & vbCrLf
        End If
        s = GetString(dr("email"))
        If s <> "" Then
            sOut &= "            @notify_email_operator_name = '" & s & "'," & vbCrLf
        End If
        s = GetString(dr("netsend"))
        If s <> "" Then
            sOut &= "            @notify_netsend_operator_name = '" & s & "'," & vbCrLf
        End If
        s = GetString(dr("page"))
        If s <> "" Then
            sOut &= "            @notify_page_operator_name = '" & s & "'," & vbCrLf
        End If
        sOut &= "            @description = '" & GetSQLString(dr("description")) & "'," & vbCrLf
        sOut &= "            @category_name = '" & GetSQLString(dr("category")) & "'," & vbCrLf
        sOut &= "            @owner_login_name = '" & GetSQLString(dr("owner")) & "'," & vbCrLf
        sOut &= "            @job_id = @JobID output" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= vbCrLf
        sOut &= "	    execute @rc = msdb.dbo.sp_add_jobserver" & vbCrLf
        sOut &= "            @job_id = @JobID," & vbCrLf
        s = ""
        If Not dtServers Is Nothing Then
            If dtServers.Rows.Count > 0 Then
                If CInt(dtServers.Rows(0)("server_id")) <> 0 Then
                    s = GetString(dtServers.Rows(0)("server"))
                End If
            End If
        End If
        If s = "" Then s = "(local)"
        sOut &= "            @server_name = '" & s & "'" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= "    else" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        print 'updating job ''" & sName & "''.'" & vbCrLf
        sOut &= "        execute @rc = msdb.dbo.sp_delete_jobstep" & vbCrLf
        sOut &= "            @job_id = @JobID," & vbCrLf
        sOut &= "            @step_id = 0" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= vbCrLf

        For Each ds As DataRow In dtSteps.Rows
            sOut &= "    execute @rc = msdb.dbo.sp_add_jobstep" & vbCrLf
            sOut &= "        @job_id = @JobID," & vbCrLf
            sOut &= "        @step_id = " & CInt(ds("step_id")) & "," & vbCrLf
            sOut &= "        @step_name = '" & GetSQLString(ds("step_name")) & "'"
            s = GetString(ds("subsystem"))
            If s <> "TSQL" Then
                sOut &= "," & vbCrLf
                sOut &= "        @subsystem = '" & s & "'"
            End If
            s = GetSQLString(ds("command"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @command = '" & s & "'"
            End If
            s = GetSQLString(ds("additional_parameters"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @additional_parameters = '" & s & "'"
            End If
            i = CInt(ds("cmdexec_success_code"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @cmdexec_success_code = " & i
            End If
            i = CInt(ds("on_success_action"))
            If i <> 1 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_success_action = " & i
            End If
            i = CInt(ds("on_success_step_id"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_success_step_id = " & i
            End If
            i = CInt(ds("on_fail_action"))
            If i <> 2 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_fail_action = " & i
            End If
            i = CInt(ds("on_fail_step_id"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_fail_step_id = " & i
            End If
            s = GetSQLString(ds("server"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @server = '" & s & "'"
            End If
            s = GetSQLString(ds("database_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @database_name = '" & s & "'"
            End If
            s = GetSQLString(ds("database_user_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @database_user_name = '" & s & "'"
            End If
            i = CInt(ds("retry_attempts"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @retry_attempts = " & i
            End If
            i = CInt(ds("retry_interval"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @retry_interval = " & i
            End If
            i = CInt(ds("os_run_priority"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @os_run_priority = " & i
            End If
            s = GetSQLString(ds("output_file_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @output_file_name = '" & s & "'"
            End If
            i = CInt(ds("flags"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @flags = " & i
            End If
            s = GetSQLString(ds("proxy"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @proxy_name = '" & s & "'"
            End If

            sOut &= vbCrLf
            sOut &= "    if @rc <> 0 break" & vbCrLf
            sOut &= vbCrLf
        Next

        b = True
        For Each ds As DataRow In dtSchedules.Rows
            If b Then
                b = False
                sOut &= "    if not exists" & vbCrLf
                sOut &= "    (" & vbCrLf
                sOut &= "        select  'a'" & vbCrLf
                sOut &= "        from    msdb.dbo.sysjobschedules" & vbCrLf
                sOut &= "        where   job_id = @JobID" & vbCrLf
                sOut &= "    )" & vbCrLf
                sOut &= "    begin" & vbCrLf
                sOut &= "        print 'creating job schedule.'" & vbCrLf
            Else
                sOut &= vbCrLf
            End If
            sOut &= "        execute @rc = msdb.dbo.sp_add_jobschedule" & vbCrLf
            sOut &= "            @job_id = @JobID," & vbCrLf
            sOut &= "            @name = '" & GetSQLString(ds("name")) & "'"
            i = CInt(ds("freq_type"))
            If i <> 1 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_type = " & i
            End If
            i = CInt(ds("freq_interval"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_interval = " & i
            End If
            i = CInt(ds("freq_subday_type"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_subday_type = " & i
            End If
            i = CInt(ds("freq_subday_interval"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_subday_interval = " & i
            End If
            i = CInt(ds("freq_relative_interval"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_relative_interval = " & i
            End If
            i = CInt(ds("freq_recurrence_factor"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @freq_recurrence_factor = " & i
            End If
            i = CInt(ds("active_end_date"))
            If i <> 99991231 Then
                sOut &= "," & vbCrLf
                sOut &= "            @active_end_date = " & i
            End If
            i = CInt(ds("active_start_time"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "            @active_start_time = " & i
            End If
            i = CInt(ds("active_end_time"))
            If i <> 235959 Then
                sOut &= "," & vbCrLf
                sOut &= "            @active_end_time = " & i
            End If

            sOut &= vbCrLf
            sOut &= "        if @rc <> 0 break" & vbCrLf
        Next
        If Not b Then
            sOut &= "    end" & vbCrLf
        End If
        sOut &= "    break" & vbCrLf
        sOut &= "end" & vbCrLf

        Return sOut
    End Function

    Public Function CommonText() As String
        Dim sOut As String = ""
        Dim dr As DataRow = Nothing
        Dim s As String
        Dim i As Integer
        Dim sName As String = ""

        If Not dtJobs Is Nothing Then
            If dtJobs.Rows.Count > 0 Then
                dr = dtJobs.Rows(0)
            End If
        End If

        If dr Is Nothing Then Return ""
        sName = GetSQLString(dr("name"))

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
        sOut &= "        print 'creating new job ''" & sName & "''.'" & vbCrLf
        sOut &= "	    execute @rc = msdb.dbo.sp_add_job" & vbCrLf
        sOut &= "            @job_name = '" & sName & "'," & vbCrLf
        sOut &= "            @enabled = 0," & vbCrLf
        i = CInt(dr("delete_level"))
        If i <> 0 Then
            sOut &= "            @delete_level = " & i & "," & vbCrLf
        End If
        sOut &= "            @description = '" & GetSQLString(dr("description")) & "'," & vbCrLf
        sOut &= "            @category_name = '" & GetSQLString(dr("category")) & "'," & vbCrLf
        sOut &= "            @owner_login_name = '" & GetSQLString(dr("owner")) & "'," & vbCrLf
        sOut &= "            @job_id = @JobID output" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= "    else" & vbCrLf
        sOut &= "    begin" & vbCrLf
        sOut &= "        print 'updating job ''" & sName & "''.'" & vbCrLf
        sOut &= "        execute @rc = msdb.dbo.sp_delete_jobstep" & vbCrLf
        sOut &= "            @job_id = @JobID," & vbCrLf
        sOut &= "            @step_id = 0" & vbCrLf
        sOut &= "        if @rc <> 0 break" & vbCrLf
        sOut &= "    end" & vbCrLf
        sOut &= vbCrLf

        For Each ds As DataRow In dtSteps.Rows
            sOut &= "    execute @rc = msdb.dbo.sp_add_jobstep" & vbCrLf
            sOut &= "        @job_id = @JobID," & vbCrLf
            sOut &= "        @step_id = " & CInt(ds("step_id")) & "," & vbCrLf
            sOut &= "        @step_name = '" & GetSQLString(ds("step_name")) & "'"
            s = GetString(ds("subsystem"))
            If s <> "TSQL" Then
                sOut &= "," & vbCrLf
                sOut &= "        @subsystem = '" & s & "'"
            End If
            s = GetSQLString(ds("command"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @command = '" & s & "'"
            End If
            s = GetSQLString(ds("additional_parameters"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @additional_parameters = '" & s & "'"
            End If
            i = CInt(ds("cmdexec_success_code"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @cmdexec_success_code = " & i
            End If
            i = CInt(ds("on_success_action"))
            If i <> 1 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_success_action = " & i
            End If
            i = CInt(ds("on_success_step_id"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_success_step_id = " & i
            End If
            i = CInt(ds("on_fail_action"))
            If i <> 2 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_fail_action = " & i
            End If
            i = CInt(ds("on_fail_step_id"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @on_fail_step_id = " & i
            End If
            s = GetSQLString(ds("server"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @server = '" & s & "'"
            End If
            s = GetSQLString(ds("database_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @database_name = '" & s & "'"
            End If
            s = GetSQLString(ds("database_user_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @database_user_name = '" & s & "'"
            End If
            i = CInt(ds("retry_attempts"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @retry_attempts = " & i
            End If
            i = CInt(ds("retry_interval"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @retry_interval = " & i
            End If
            i = CInt(ds("os_run_priority"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @os_run_priority = " & i
            End If
            s = GetSQLString(ds("output_file_name"))
            If s <> "" Then
                sOut &= "," & vbCrLf
                sOut &= "        @output_file_name = '" & s & "'"
            End If
            i = CInt(ds("flags"))
            If i <> 0 Then
                sOut &= "," & vbCrLf
                sOut &= "        @flags = " & i
            End If

            sOut &= vbCrLf
            sOut &= "    if @rc <> 0 break" & vbCrLf
            sOut &= vbCrLf
        Next

        sOut &= "    break" & vbCrLf
        sOut &= "end" & vbCrLf

        Return sOut
    End Function

    Private Function GetJob(ByVal sJobID As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim JobDetails As New DataSet

        s = "select j.*,suser_sname(j.owner_sid) owner,c.name category,"
        s &= "o1.name email,o2.name netsend,o3.name page "
        s &= "from dbo.sysjobs j "
        s &= "join dbo.syscategories c "
        s &= "on c.category_id=j.category_id "
        s &= "left join dbo.sysoperators o1 "
        s &= "on o1.id = j.notify_email_operator_id "
        s &= "left join dbo.sysoperators o2 "
        s &= "on o2.id = j.notify_netsend_operator_id "
        s &= "left join dbo.sysoperators o3 "
        s &= "on o3.id = j.notify_page_operator_id "
        s &= "where job_id='" & sJobID & "'"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(JobDetails)

        Return JobDetails.Tables(0)
    End Function

    Private Function GetServers(ByVal JobID As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim JobDetails As New DataSet

        s = "select j.*,s.name server "
        s &= "from dbo.sysjobservers j "
        s &= "join master.sys.servers s on j.server_id=s.server_id "
        s &= "where job_id='" & JobID & "'"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(JobDetails)

        Return JobDetails.Tables(0)
    End Function

    Private Function GetSteps(ByVal JobID As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim JobDetails As New DataSet

        s = "select j.*,p.name proxy from dbo.sysjobsteps j "
        s &= "left join dbo.sysproxies p on p.proxy_id = j.proxy_id "
        s &= "where job_id='" & JobID & "'"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(JobDetails)

        Return JobDetails.Tables(0)
    End Function

    Private Function GetSchedules(ByVal JobID As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim JobDetails As New DataSet

        s = "select * from dbo.sysjobschedules j "
        s &= "join dbo.sysschedules s on s.Schedule_id = j.schedule_id "
        s &= "where j.job_id='" & JobID & "' and s.enabled=1"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(JobDetails)

        Return JobDetails.Tables(0)
    End Function

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
            ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Console.WriteLine(e.Message)
    End Sub

    Public Function GetSQLString(ByVal objValue As Object) As String
        Return Replace(GetString(objValue), "'", "''")
    End Function

    Public Function GetString(ByVal objValue As Object) As String
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
End Class
