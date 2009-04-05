Option Explicit On 
Option Strict On

Public Class [Publics]
    Private Shared sPath As String = ""

    Public Shared ReadOnly Property Path() As String
        Get
            If sPath = "" Then InitPath()
            Path = sPath
        End Get
    End Property

    Private Shared Sub InitPath()
        Dim s As String

        Try
            Dim d As System.DirectoryServices.ActiveDirectory.Domain
            d = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain
            s = d.Name
            Dim a() As String = Split(s, ".")
            Dim b As Boolean = False
            sPath = "LDAP://"
            For Each s In a
                If b Then
                    sPath &= ","
                End If
                sPath &= "dc=" & s
                b = True
            Next
        Catch
            sPath = "WinNT://" & Environment.MachineName & ",computer"
        End Try
    End Sub
End Class
