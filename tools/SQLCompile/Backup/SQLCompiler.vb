Option Explicit On 
Option Strict On

Module SQLCompiler
    Dim bVerbose As Boolean = False

    Sub Main()
        Dim c As New scl
        AddHandler c.Notify, AddressOf WriteLine
        Dim FileName As String = GetCommandParameter("-f")
        Dim Database As String = GetCommandParameter("-d")
        If GetSwitch("-v") Then bVerbose = True

        FileName = FileName.Replace("""", "")
        If FileName <> "" Then
            c.Process(FileName, Database)
        Else
            Console.WriteLine("Usage: SQLCompile -fList.scl [-dDatabase] [-v]")
        End If
        Console.WriteLine("")
        Console.WriteLine("Press enter to continue...")
        Console.Read()
    End Sub

    Private Sub WriteLine(ByVal sMsg As String, ByVal sType As String)
        If sType = "E" Or bVerbose Then
            Console.WriteLine(sMsg)
        End If
    End Sub

    Private Function GetSwitch(ByRef sSwitch As String) As Boolean
        Dim sCommand As String

        sCommand = Microsoft.VisualBasic.Command()
        If InStr(1, sCommand, sSwitch, CompareMethod.Text) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetCommandParameter(ByRef sSwitch As String) As String
        Dim sCommand As String
        Dim sParameter As String
        Dim i As Integer

        sCommand = Microsoft.VisualBasic.Command()
        i = InStr(1, sCommand, sSwitch, CompareMethod.Text)
        sParameter = ""
        If i > 0 Then
            sParameter = Mid(sCommand, i + 2)
            If Mid(sParameter, 1, 1) = """" Then
                sParameter = Mid(sParameter, 2)
                i = InStr(1, sParameter, """", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            Else
                i = InStr(1, sParameter, " ", CompareMethod.Text)
                If i > 0 Then
                    sParameter = Mid(sParameter, 1, i - 1)
                End If
            End If
        End If
        GetCommandParameter = sParameter
    End Function
End Module
