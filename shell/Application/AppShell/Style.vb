Option Explicit On
Option Strict On

Public Class shellStyle
    Private sID As String
    Private oRF As System.Drawing.Color
    Private oRB As System.Drawing.Color
    Private oSF As System.Drawing.Color
    Private oSB As System.Drawing.Color

    Public Sub New(ByVal ID As String, ByVal sRowForeColour As String, ByVal sRowBackColour As String, _
                   ByVal sSelForeColour As String, ByVal sSelBackColour As String)
        sID = ID
        oRF = NameToColour(sRowForeColour)
        oRB = NameToColour(sRowBackColour)
        oSF = NameToColour(sSelForeColour)
        oSB = NameToColour(sSelBackColour)
    End Sub

    Public ReadOnly Property ID() As String
        Get
            ID = sID
        End Get
    End Property

    Public ReadOnly Property RowForeColour() As System.Drawing.Color
        Get
            If oRF.IsKnownColor Then
                RowForeColour = oRF
            Else
                RowForeColour = DialogStyle.NameToColour(DialogStyle.ForeColour)
            End If
        End Get
    End Property

    Public ReadOnly Property RowBackColour() As System.Drawing.Color
        Get
            If oRB.IsKnownColor Then
                RowBackColour = oRB
            Else
                RowBackColour = DialogStyle.NameToColour(DialogStyle.BackNormal)
            End If
        End Get
    End Property

    Public ReadOnly Property SelForeColour() As System.Drawing.Color
        Get
            If oSF.IsKnownColor Then
                SelForeColour = oSF
            Else
                SelForeColour = DialogStyle.NameToColour(DialogStyle.ForeColour)
            End If
        End Get
    End Property

    Public ReadOnly Property SelBackColour() As System.Drawing.Color
        Get
            If oSB.IsKnownColor Then
                SelBackColour = oSB
            Else
                SelBackColour = Color.LightGray
            End If
        End Get
    End Property

    Private Function NameToColour(ByVal Name As String) As Color
        Dim c As Color

        If Mid(Name, 1, 3) = "rgb" Then
            Dim s As String = Mid(Name, 4)
            Dim a() As String = Split(s, ",")
            c = Color.FromArgb(CInt(a(0)), CInt(a(1)), CInt(a(2)))
        Else
            c = Color.FromName(Name)
        End If
        Return c
    End Function
End Class

Public Class shellStyles
    Private Values As New Hashtable
    Private Keys() As String

    Public Function Add(ByVal ID As String, _
                    ByVal RowForeColour As String, _
                    ByVal RowBackColour As String, _
                    ByVal SelForeColour As String, _
                    ByVal SelBackColour As String) As shellStyle
        Dim parm As New shellStyle(ID, RowForeColour, RowBackColour, _
                        SelForeColour, SelBackColour)
        Dim i As Integer

        i = Values.Count
        ReDim Preserve Keys(i)
        Values.Add(ID, parm)
        Keys(i) = ID
        Return parm
    End Function

    Public ReadOnly Property Item(ByVal index As String) As shellStyle
        Get
            Try
                Return CType(Values.Item(index), shellStyle)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class
