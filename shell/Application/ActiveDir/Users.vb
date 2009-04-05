Option Explicit On 
Option Strict On

Public Class Users
    Private dt As DataTable

    Public ReadOnly Property data() As DataTable
        Get
            data = dt
        End Get
    End Property

    Public Sub New()
        GetData()
    End Sub

    Private Sub GetData()
        dt = New DataTable("Data")
        dt.Columns.Add("UserName", Type.GetType("System.String"))
        dt.Columns.Add("DisplayName", Type.GetType("System.String"))

        If Mid(Publics.Path, 1, 6) = "WinNT:" Then
            GetDataNT()
        Else
            GetDataAD()
        End If
    End Sub

    Private Sub GetDataAD()
        Dim entry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        Dim search As New System.DirectoryServices.DirectorySearcher(entry)
        Dim searchResult As System.DirectoryServices.SearchResultCollection
        Dim dr As DataRow

        search.Filter = "(&(objectCategory=person)(objectClass=user))"

        search.PropertiesToLoad.Add("sAMAccountName")
        search.PropertiesToLoad.Add("displayName")
        searchResult = search.FindAll()
        If Not (searchResult Is Nothing) Then
            For Each sr As System.DirectoryServices.SearchResult In searchResult
                If Not sr.Properties("sAMAccountName") Is Nothing Then
                    dr = dt.NewRow()
                    dr("UserName") = CType(sr.Properties("sAMAccountName")(0), String)
                    dr("DisplayName") = CType(sr.Properties("displayName")(0), String)
                    dt.Rows.Add(dr)
                End If
            Next
        End If
        dt.DefaultView.Sort = "DisplayName"
    End Sub

    Private Sub GetDataNT()
        Dim oEntry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        UserNT(oEntry)
    End Sub

    Private Sub UserNT(ByVal entry As System.DirectoryServices.DirectoryEntry)
        Dim dr As DataRow

        If LCase(entry.SchemaClassName) = "user" Then
            dr = dt.NewRow()
            dr("UserName") = entry.Name
            dr("DisplayName") = entry.Name
            dt.Rows.Add(dr)
        Else
            For Each e As System.DirectoryServices.DirectoryEntry In entry.Children
                UserNT(e)
            Next
        End If
    End Sub
End Class
