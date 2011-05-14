Public Class DialogStyle
    'Public Const ForeColour As String = "Navy"
    'Public Const ForeError As String = "Red"
    'Public Const BackColour As String = "LightCyan"
    'Public Const BackNormal As String = "White"
    'Public Const BackRequired As String = "Yellow"
    'Public Const BorderNormal As String = "RoyalBlue"
    'Public Const BorderError As String = "Red"

    'Public Const ForeColour As String = "Black"
    'Public Const ForeError As String = "Red"
    'Public Const BackColour As String = "Ivory"
    'Public Const BackNormal As String = "White"
    'Public Const BackRequired As String = "Gold"
    'Public Const BorderNormal As String = "Silver"
    'Public Const BorderError As String = "Salmon"

    Public Const ForeColour As String = "DarkGreen"
    Public Const ForeError As String = "Red"
    Public Const BackColour As String = "Ivory"
    Public Const BackNormal As String = "White"
    Public Const BackRequired As String = "rgb255,255,170"
    Public Const BorderNormal As String = "RoyalBlue"
    Public Const BorderError As String = "Red"
    Public Const BorderWidth As Integer = 1

    Public Const SelForeColour As String = "Black"
    Public Const SelBackColour As String = "yellowgreen" ' "greenyellow"

    Public Const ToolStart As String = "LightSteelBlue"
    Public Const ToolEnd As String = "Ivory"

    Public Shared Function NameToColour(ByVal Name As String) As Color
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
