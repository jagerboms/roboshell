Option Explicit On 
Option Strict On

Imports ActiveDs

Public Class Members
    Private sGroup As String
    Private sName As String
    Private dt As DataTable

    Public Property LoginName() As String
        Get
            LoginName = sGroup
        End Get
        Set(ByVal Value As String)
            sGroup = Value
            InitGroup()
        End Set
    End Property

    Public ReadOnly Property data() As DataTable
        Get
            data = dt
        End Get
    End Property

    Public ReadOnly Property GroupName() As String
        Get
            GroupName = sName
        End Get
    End Property

    Public Sub New(ByVal LoginName As String)
        sGroup = LoginName
        InitGroup()
    End Sub

    Private Sub InitGroup()
        dt = New DataTable("Data")
        dt.Columns.Add("MemberName", Type.GetType("System.String"))
        If Mid(Publics.Path, 1, 6) = "WinNT:" Then
            InitGroupNT()
        Else
            InitGroupAD()
        End If
    End Sub

    Private Sub InitGroupAD()
        Dim s As String
        Dim entry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        Dim ent As System.DirectoryServices.DirectoryEntry
        Dim search As New System.DirectoryServices.DirectorySearcher(entry)
        Dim searchResult As System.DirectoryServices.SearchResult

        dt = New DataTable("Data")
        dt.Columns.Add("MemberName", Type.GetType("System.String"))

        search.Filter = "(name=" & Replace(sGroup, " ", "\20") & ")"
        searchResult = search.FindOne()
        If Not (searchResult Is Nothing) Then
            ent = searchResult.GetDirectoryEntry()
            If ent.SchemaClassName = "group" Then
                Try
                    sName = ent.Properties.Item("description")(0).ToString
                Catch ex As Exception
                    sName = sGroup
                End Try
                For Each sDN As String In ent.Properties.Item("member")
                    Dim member As New System.DirectoryServices.DirectoryEntry( _
                                              "LDAP://" & sDN)
                    Try
                        s = CType(member.Properties.Item("displayName").Value, String)
                    Catch ex As Exception
                        s = sDN
                    End Try
                    If s <> "" Then
                        Dim dr As DataRow = dt.NewRow()
                        dr("MemberName") = s
                        dt.Rows.Add(dr)
                    End If
                Next
            ElseIf ent.SchemaClassName = "user" Then
                Dim dr As DataRow = dt.NewRow()
                s = CType(ent.Properties.Item("name").Value, String)
                dr("MemberName") = s
                dt.Rows.Add(dr)
            End If
        End If
    End Sub

    Private Sub InitGroupNT()
        Dim oEntry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        GroupNT(oEntry)
    End Sub

    Private Sub GroupNT(ByVal entry As System.DirectoryServices.DirectoryEntry)
        Dim MembersCollection As ActiveDs.IADsMembers

        If entry.Name = sGroup Then
            MembersCollection = CType(entry.Invoke("members"), ActiveDs.IADsMembers)
            Dim filter As System.Object() = {"user"}
            MembersCollection.Filter = filter
            Dim user As ActiveDs.IADsUser
            For Each user In MembersCollection
                Dim dr As DataRow = dt.NewRow()
                dr("MemberName") = user.Name
                dt.Rows.Add(dr)
            Next
        Else
            For Each e As System.DirectoryServices.DirectoryEntry In entry.Children
                GroupNT(e)
            Next
        End If
    End Sub
End Class
