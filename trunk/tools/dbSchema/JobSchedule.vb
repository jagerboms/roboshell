Option Explicit On
Option Strict On

'sp_add_schedule [ @schedule_name = ] 'schedule_name' 
'    [ , [ @enabled = ] enabled ]
'    [ , [ @freq_type = ] freq_type ]
'    [ , [ @freq_interval = ] freq_interval ] 
'    [ , [ @freq_subday_type = ] freq_subday_type ] 
'    [ , [ @freq_subday_interval = ] freq_subday_interval ] 
'    [ , [ @freq_relative_interval = ] freq_relative_interval ] 
'    [ , [ @freq_recurrence_factor = ] freq_recurrence_factor ] 
'    [ , [ @active_start_date = ] active_start_date ] 
'    [ , [ @active_end_date = ] active_end_date ] 
'    [ , [ @active_start_time = ] active_start_time ] 
'    [ , [ @active_end_time = ] active_end_time ] 
'    [ , [ @owner_login_name = ] 'owner_login_name' ]
'    [ , [ @schedule_uid = ] schedule_uid OUTPUT ]
'    [ , [ @schedule_id = ] schedule_id OUTPUT ]
'    [ , [ @originating_server = ] server_name ] /* internal */

'sp_attach_schedule
'     { [ @job_id = ] job_id | [ @job_name = ] 'job_name' } , 
'     { [ @schedule_id = ] schedule_id 
'     | [ @schedule_name = ] 'schedule_name' }

Imports System
Imports System.Collections

Public Enum JobScheduleFrequency
    Once = 1
    Daily = 4
    Weekly = 8
    Monthly = 16
    MonthlyRelative = 32
    AgentStart = 64
    WhenIdle = 128
End Enum

Public Enum JobScheduleFrequencySubday
    Specified = 1
    Minutes = 4
    Hours = 8
End Enum

Public Enum JobScheduleFrequencyRelInt
    First = 1
    Second = 2
    Third = 4
    Fourth = 8
    Last = 16
End Enum

Public Class JobSchedule
    Private sName As String
    Private iEnabled As Integer = 1
    Private eFrequencyType As JobScheduleFrequency = Nothing
    Private iFrequencyInterval As Integer = 1

    '4 (daily)    Every freq_interval days.
    '8 (weekly)   freq_interval is one or more of the following (combined with an OR logical operator): 
    '                1 = Sunday 
    '                2 = Monday 
    '                4 = Tuesday 
    '                8 = Wednesday 
    '               16 = Thursday 
    '               32 = Friday 
    '               64 = Saturday
    '16 (monthly) On the freq_interval day of the month.
    '32 (monthly relative) freq_interval is one of the following: 
    '                1 = Sunday
    '                2 = Monday
    '                3 = Tuesday
    '                4 = Wednesday
    '                5 = Thursday
    '                6 = Friday
    '                7 = Saturday
    '                8 = Day
    '                9 = Weekday
    '               10 = Weekend day

    Private eFrequencySubdayType As JobScheduleFrequencySubday = Nothing
    Private iFrequencySubdayInterval As Integer = 0
    Private eFrequencyRelativeInterval As JobScheduleFrequencyRelInt = Nothing
    Private iFrequencyRecurrenceFactor As Integer = 0
    Private iActiveStartDate As Integer = 0
    Private iActiveEndDate As Integer = 0
    Private iActiveStartTime As Integer = 0
    Private iActiveEndTime As Integer = 0
    Private sOwnerLoginName As String = ""

    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal nm As String)
            sName = nm
        End Set
    End Property

    Public Property Enabled() As Integer
        Get
            Enabled = iEnabled
        End Get
        Set(ByVal pe As Integer)
            iEnabled = pe
        End Set
    End Property

    Public Property FrequencyType() As JobScheduleFrequency
        Get
            FrequencyType = eFrequencyType
        End Get
        Set(ByVal ofa As JobScheduleFrequency)
            eFrequencyType = ofa
        End Set
    End Property

    Public Property iFrequencyType() As Integer
        Get
            iFrequencyType = Me.iGetFrequencyType(eFrequencyType)
        End Get
        Set(ByVal sa As Integer)
            eFrequencyType = Me.GetFrequencyType(sa, 0)
        End Set
    End Property

    Public Property FrequencyInterval() As Integer
        Get
            FrequencyInterval = iFrequencyInterval
        End Get
        Set(ByVal fi As Integer)
            iFrequencyInterval = fi
        End Set
    End Property

    Public Property FrequencySubdayType() As JobScheduleFrequencySubday
        Get
            FrequencySubdayType = eFrequencySubdayType
        End Get
        Set(ByVal fsd As JobScheduleFrequencySubday)
            eFrequencySubdayType = fsd
        End Set
    End Property

    Public Property iFrequencySubdayType() As Integer
        Get
            iFrequencySubdayType = eFrequencySubdayType
        End Get
        Set(ByVal fsd As Integer)
            eFrequencySubdayType = Me.GetFrequencySubday(fsd, 0)
        End Set
    End Property

    Public Property FrequencySubdayInterval() As Integer
        Get
            FrequencySubdayInterval = iFrequencySubdayInterval
        End Get
        Set(ByVal fsi As Integer)
            iFrequencySubdayInterval = fsi
        End Set
    End Property

    Public Property FrequencyRelativeInterval() As JobScheduleFrequencyRelInt
        Get
            FrequencyRelativeInterval = eFrequencyRelativeInterval
        End Get
        Set(ByVal fri As JobScheduleFrequencyRelInt)
            eFrequencyRelativeInterval = fri
        End Set
    End Property

    Public Property iFrequencyRelativeInterval() As Integer
        Get
            iFrequencyRelativeInterval = Me.iGetFrequencyRelInt(eFrequencyRelativeInterval)
        End Get
        Set(ByVal fri As Integer)
            eFrequencyRelativeInterval = Me.GetFrequencyRelInt(fri, 0)
        End Set
    End Property

    Public Property FrequencyRecurrenceFactor() As Integer
        Get
            FrequencyRecurrenceFactor = iFrequencyRecurrenceFactor
        End Get
        Set(ByVal frf As Integer)
            iFrequencyRecurrenceFactor = frf
        End Set
    End Property

    Public Property ActiveStartDate() As Integer
        Get
            ActiveStartDate = iActiveStartDate
        End Get
        Set(ByVal asd As Integer)
            iActiveStartDate = asd
        End Set
    End Property

    Public Property ActiveEndDate() As Integer
        Get
            ActiveEndDate = iActiveEndDate
        End Get
        Set(ByVal aed As Integer)
            iActiveEndDate = aed
        End Set
    End Property

    Public Property ActiveStartTime() As Integer
        Get
            ActiveStartTime = iActiveStartTime
        End Get
        Set(ByVal ast As Integer)
            iActiveStartTime = ast
        End Set
    End Property

    Public Property ActiveEndTime() As Integer
        Get
            ActiveEndTime = iActiveEndTime
        End Get
        Set(ByVal aet As Integer)
            iActiveEndTime = aet
        End Set
    End Property

    Public Property OwnerLoginName() As String
        Get
            OwnerLoginName = sOwnerLoginName
        End Get
        Set(ByVal oln As String)
            sOwnerLoginName = oln
        End Set
    End Property

#Region "Methods"
    Public Sub New(ByVal pScheduleName As String, ByVal iFreqType As Integer)
        sName = pScheduleName
        eFrequencyType = Me.GetFrequencyType(iFreqType, 1)
    End Sub

    Public Function XMLText(ByVal sTab As String) As String
        Dim sOut As String = ""
        Dim s As String = vbCrLf & sTab & "          "
        Dim i As Integer
        Dim ft As Integer

        sOut &= sTab & "<schedule name='" & sName & "'"
        If iEnabled <> 1 Then
            sOut &= s & "enabled='0'"
        End If
        ft = Me.GetFrequencyType(eFrequencyType, 0)
        sOut &= s & "freqtype='" & ft & "'"
        If ft = 4 Or ft = 8 Or ft = 16 Or ft = 32 Then
            sOut &= s & "freqinterval='" & iFrequencyInterval & "'"
        End If
        i = Me.iGetFrequencySubday(eFrequencySubdayType)
        If i = 1 Or i = 4 Or i = 8 Then
            sOut &= s & "freqsubday='" & i & "'"
            If i <> 1 Then
                sOut &= s & "freqsubdayinterval='" & iFrequencySubdayInterval & "'"
            End If
        End If
        If ft = 32 Then
            i = Me.iGetFrequencyRelInt(eFrequencyRelativeInterval)
            If i = 1 Or i = 2 Or i = 4 Or i = 8 Or i = 16 Then
                sOut &= s & "freqrelativeinterval='" & i & "'"
            End If
        End If
        If ft = 8 Or ft = 16 Or ft = 32 Then
            i = iFrequencyRecurrenceFactor
            If i > 0 Then
                sOut &= s & "freqrecurrencefactor='" & i & "'"
            End If
        End If
        i = iActiveStartDate
        If i > 1990010 Then
            sOut &= s & "activestartdate='" & i & "'"
        End If
        i = iActiveEndDate
        If i > 1990010 And i < 99991231 Then
            sOut &= s & "activeenddate='" & i & "'"
        End If
        i = iActiveStartTime
        If i > 0 And i < 235960 Then
            sOut &= s & "activestarttime='" & i & "'"
        End If
        i = ActiveEndTime
        If i > -1 And i < 235959 Then
            sOut &= s & "activeendtime='" & i & "'"
        End If
        If sOwnerLoginName <> "" Then
            sOut &= s & "ownerloginname='" & sOwnerLoginName & "'"
        End If
        sOut &= " />" & vbCrLf
        Return sOut
    End Function
#End Region

    Private Function GetFrequencyType(ByVal i As Integer, ByVal iDefault As Integer) As JobScheduleFrequency
        If i < 1 Or i > 4 Then i = iDefault
        Select Case i
            Case 1
                Return JobScheduleFrequency.Once
            Case 4
                Return JobScheduleFrequency.Daily
            Case 8
                Return JobScheduleFrequency.Weekly
            Case 16
                Return JobScheduleFrequency.Monthly
            Case 32
                Return JobScheduleFrequency.MonthlyRelative
            Case 64
                Return JobScheduleFrequency.AgentStart
            Case 128
                Return JobScheduleFrequency.WhenIdle
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function iGetFrequencyType(ByVal jsa As JobScheduleFrequency) As Integer
        Select Case jsa
            Case JobScheduleFrequency.Once
                Return 1
            Case JobScheduleFrequency.Daily
                Return 4
            Case JobScheduleFrequency.Weekly
                Return 8
            Case JobScheduleFrequency.Monthly
                Return 16
            Case JobScheduleFrequency.MonthlyRelative
                Return 32
            Case JobScheduleFrequency.AgentStart
                Return 64
            Case JobScheduleFrequency.WhenIdle
                Return 128
            Case Else
                Return 0
        End Select
    End Function

    Private Function GetFrequencySubday(ByVal i As Integer, ByVal iDefault As Integer) As JobScheduleFrequencySubday
        If i <> 1 And i <> 4 And i <> 8 Then i = iDefault
        if i = 1
            Return JobScheduleFrequencySubday.Specified
        ElseIf i = 4 Then
            Return JobScheduleFrequencySubday.Minutes
        ElseIf i = 8 Then
            Return JobScheduleFrequencySubday.Hours
        Else
            Return Nothing
        End If
    End Function

    Private Function iGetFrequencySubday(ByVal jsd As JobScheduleFrequencySubday) As Integer
        Select Case jsd
            Case JobScheduleFrequencySubday.Specified
                Return 1
            Case JobScheduleFrequencySubday.Minutes
                Return 4
            Case JobScheduleFrequencySubday.Hours
                Return 8
            Case Else
                Return 0
        End Select
    End Function

    Private Function GetFrequencyRelInt(ByVal i As Integer, ByVal iDefault As Integer) As JobScheduleFrequencyRelInt
        If i <> 1 And i <> 2 And i <> 4 And i <> 8 Then i = iDefault
        If i = 1 Then
            Return JobScheduleFrequencyRelInt.First
        ElseIf i = 2 Then
            Return JobScheduleFrequencyRelInt.Second
        ElseIf i = 4 Then
            Return JobScheduleFrequencyRelInt.Third
        ElseIf i = 8 Then
            Return JobScheduleFrequencyRelInt.Fourth
        ElseIf i = 16 Then
            Return JobScheduleFrequencyRelInt.Last
        Else
            Return Nothing
        End If
    End Function

    Private Function iGetFrequencyRelInt(ByVal fri As JobScheduleFrequencyRelInt) As Integer
        Select Case fri
            Case JobScheduleFrequencyRelInt.First
                Return 1
            Case JobScheduleFrequencyRelInt.Second
                Return 2
            Case JobScheduleFrequencyRelInt.Third
                Return 4
            Case JobScheduleFrequencyRelInt.Fourth
                Return 8
            Case JobScheduleFrequencyRelInt.Last
                Return 16
            Case Else
                Return 0
        End Select
    End Function
End Class

Public Class JobSchedules
    Inherits CollectionBase

    Private slib As sql

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal ScheduleName As String) As JobSchedule
        Get
            For Each js As JobSchedule In Me
                If js.Name = ScheduleName Then
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

    Public Function Add(ByVal js As JobSchedule) As Integer
        Return List.Add(js)
    End Function

    Public Function XMLText(ByVal sTab As String) As String
        Dim ss As String = ""
        Dim sOut As String = ""
        Dim cJS As JobSchedule

        For Each cJS In Me
            ss &= cJS.XMLText(sTab & "  ")
        Next
        If ss <> "" Then
            sOut &= sTab & "<schedules>" & vbCrLf
            sOut &= ss
            sOut &= sTab & "</schedules>" & vbCrLf
        End If
        Return sOut
    End Function
#End Region
End Class
