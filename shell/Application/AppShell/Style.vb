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
        oRF = System.Drawing.Color.FromName(sRowForeColour)
        oRB = System.Drawing.Color.FromName(sRowBackColour)
        oSF = System.Drawing.Color.FromName(sSelForeColour)
        oSB = System.Drawing.Color.FromName(sSelBackColour)
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
                RowForeColour = Color.Black
            End If
        End Get
    End Property

    Public ReadOnly Property RowBackColour() As System.Drawing.Color
        Get
            If oRB.IsKnownColor Then
                RowBackColour = oRB
            Else
                RowBackColour = Color.White
            End If
        End Get
    End Property

    Public ReadOnly Property SelForeColour() As System.Drawing.Color
        Get
            If oSF.IsKnownColor Then
                SelForeColour = oSF
            Else
                SelForeColour = Color.Black
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
