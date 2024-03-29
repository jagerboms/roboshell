Option Explicit On 
Option Strict On

Public Class CallAsmDefn
    Inherits ObjectDefn

    Private sLibraryName As String
    Private sClassName As String
    Private sMethodName As String
    Private sObjectParamName As String
    Private sReturnParamName As String

    Public ReadOnly Property LibraryName() As String
        Get
            LibraryName = sLibraryName
        End Get
    End Property

    Public ReadOnly Property ClassName() As String
        Get
            ClassName = sClassName
        End Get
    End Property

    Public ReadOnly Property MethodName() As String
        Get
            MethodName = sMethodName
        End Get
    End Property

    Public ReadOnly Property ObjectParamName() As String
        Get
            ObjectParamName = sObjectParamName
        End Get
    End Property

    Public ReadOnly Property ReturnParamName() As String
        Get
            ReturnParamName = sReturnParamName
        End Get
    End Property

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New CallAsm(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "LibraryName"
                sLibraryName = GetString(Value)
            Case "ClassName"
                sClassName = GetString(Value)
            Case "MethodName"
                sMethodName = GetString(Value)
            Case "ObjectParamName"
                sObjectParamName = GetString(Value)
            Case "ReturnParamName"
                sReturnParamName = GetString(Value)
            Case Else
                MsgBox(Name & " property is not supported by Call Assembly object")
        End Select
    End Sub
End Class

Public Class CallAsm
    Inherits ShellObject

    Private sDefn As CallAsmDefn

    Public Sub New(ByVal Defn As CallAsmDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Try
            Me.Parms.MergeValues(Parms)
            Dim assem As System.Reflection.Assembly
            Dim objType As Type = Nothing
            Dim i As Integer
            Dim args() As Object = Nothing
            Dim oObject As Object = Nothing
            Dim pr As ShellProperty
            Dim propInfo As System.Reflection.PropertyInfo
            Dim methInfo As System.Reflection.MethodInfo
            Dim result As Object
            Dim b As Boolean = False
            Dim p As shellParameter

            If sDefn.ObjectParamName = "" Then
                b = True
            Else
                If Parms.Item(sDefn.ObjectParamName).Value Is Nothing Then
                    b = True
                End If
            End If
            If b Then
                assem = System.Reflection.Assembly.Load(sDefn.LibraryName)
                objType = assem.GetType(sDefn.LibraryName & "." & sDefn.ClassName)

                i = 0
                For Each p In Me.Parms
                    pr = sDefn.Properties.Item(p.Name, "cr")
                    If Not pr Is Nothing Then
                        ReDim Preserve args(i)
                        args(i) = p.Value
                        i += 1
                    End If
                Next
                oObject = Activator.CreateInstance(objType, args)
                If sDefn.ObjectParamName <> "" Then
                    Parms.Item(sDefn.ObjectParamName).Value = oObject  'place object in parameter 
                End If
            End If

            For Each pr In sDefn.Properties
                If pr.Type = "ip" Then
                    propInfo = objType.GetProperty(CType(pr.Value, String))
                    p = Me.Parms.Item(pr.Name)
                    Select Case p.ValueType
                        Case DbType.Int32
                            propInfo.SetValue(oObject, CInt(p.Value), Nothing)
                        Case DbType.Double
                            propInfo.SetValue(oObject, CDbl(p.Value), Nothing)
                        Case DbType.Date
                            propInfo.SetValue(oObject, CDate(p.Value), Nothing)
                        Case DbType.Object
                            propInfo.SetValue(oObject, p.Value, Nothing)
                        Case Else
                            propInfo.SetValue(oObject, p.Value, Nothing)
                    End Select
                End If
            Next

            If sDefn.MethodName <> "" Then
                i = 0
                args = Nothing
                For Each p In Me.Parms
                    pr = sDefn.Properties.Item(p.Name, "mh")
                    If Not pr Is Nothing Then
                        ReDim Preserve args(i)

                        Select Case p.ValueType
                            Case DbType.Int32
                                args(i) = CInt(p.Value)
                            Case DbType.Double
                                args(i) = CDbl(p.Value)
                            Case DbType.Date
                                args(i) = CDate(p.Value)
                            Case Else
                                args(i) = p.Value
                        End Select

                        i += 1
                    End If
                Next

                methInfo = objType.GetMethod(sDefn.MethodName)
                result = methInfo.Invoke(oObject, args)
                If sDefn.ReturnParamName <> "" Then
                    Parms.Item(sDefn.ReturnParamName).Value = result  'place result in output parameter 
                End If
            End If

            For Each pr In sDefn.Properties
                If pr.Type = "op" Then
                    propInfo = objType.GetProperty(CType(pr.Value, String))
                    Me.Parms.Item(pr.Name).Value = propInfo.GetValue(oObject, Nothing)
                    If Me.Parms.Item(pr.Name).ValueType = DbType.Date Then
                        If CType(Me.Parms.Item(pr.Name).Value, DateTime) _
                                                = System.DateTime.MinValue Then
                            Me.Parms.Item(pr.Name).Value = Nothing
                        End If
                    End If
                End If
            Next

            Me.OnExitOkay()
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                Me.Messages.Add("E", ex.ToString)
            Else
                Dim ex2 As Exception = ex.InnerException
                Do While Not ex2 Is Nothing
                    Me.Messages.Add("E", ex2.ToString)
                    ex2 = ex2.InnerException
                Loop
            End If
            Me.OnExitFail()
        End Try
    End Sub
End Class
