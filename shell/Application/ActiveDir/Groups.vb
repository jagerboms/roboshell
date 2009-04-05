Option Explicit On 
Option Strict On

Public Class Groups
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
        dt.Columns.Add("GroupID", Type.GetType("System.String"))
        dt.Columns.Add("GroupName", Type.GetType("System.String"))

        If Mid(Publics.Path, 1, 6) = "WinNT:" Then
            GetDataNT()
        Else
            GetDataAD()
        End If
    End Sub

    Private Sub GetDataAD()
        Dim s As String
        Dim entry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        Dim search As New System.DirectoryServices.DirectorySearcher(entry)
        Dim searchResult As System.DirectoryServices.SearchResultCollection
        Dim dr As DataRow

        search.Filter = "(objectCategory=group)"
        search.PropertiesToLoad.Add("sAMAccountName")
        search.PropertiesToLoad.Add("description")
        searchResult = search.FindAll()
        If Not (searchResult Is Nothing) Then
            For Each sr As System.DirectoryServices.SearchResult In searchResult
                If Not sr.Properties("sAMAccountName") Is Nothing Then
                    dr = dt.NewRow()
                    s = CType(sr.Properties("sAMAccountName")(0), String)
                    dr("GroupID") = s
                    If sr.Properties("description") Is Nothing Then
                        dr("GroupName") = s
                    Else
                        dr("GroupName") = CType(sr.Properties("description")(0), String)
                    End If
                    dt.Rows.Add(dr)
                End If
            Next
        End If
    End Sub

    Private Sub GetDataNT()
        Dim oEntry As New System.DirectoryServices.DirectoryEntry(Publics.Path)
        GroupNT(oEntry)
    End Sub

    Private Sub GroupNT(ByVal entry As System.DirectoryServices.DirectoryEntry)
        Dim dr As DataRow

        If LCase(entry.SchemaClassName) = "group" Then
            dr = dt.NewRow()
            dr("GroupID") = entry.Name
            dr("GroupName") = entry.Name
            dt.Rows.Add(dr)
        Else
            For Each e As System.DirectoryServices.DirectoryEntry In entry.Children
                GroupNT(e)
            Next
        End If
    End Sub
End Class
