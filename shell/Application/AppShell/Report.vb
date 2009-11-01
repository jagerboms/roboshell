Option Explicit On 
Option Strict On

Imports System.Data.SqlClient
Imports System.Drawing.Printing

Public Class ReportDefn
    Inherits ObjectDefn

    Private sDataParameter As String
    Private sTitle As String
    Private bPrintPreview As Boolean
    Private bDefaultPrinter As Boolean

    Public ReadOnly Property DataParameter() As String
        Get
            DataParameter = sDataParameter
        End Get
    End Property

    Public ReadOnly Property Title() As String
        Get
            Title = sTitle
        End Get
    End Property

    Public ReadOnly Property PrintPreview() As Boolean
        Get
            PrintPreview = bPrintPreview
        End Get
    End Property

    Public ReadOnly Property DefaultPrinter() As Boolean
        Get
            DefaultPrinter = bDefaultPrinter
        End Get
    End Property

    Public Sub New(ByVal sName As String)
        Me.Name = sName
    End Sub

    Public Function Create() As ShellObject
        Return CType(New Report(Me), ShellObject)
    End Function

    Public Overrides Sub SetProperty(ByVal Name As String, ByVal Value As Object)

        Select Case Name
            Case "Title"
                sTitle = GetString(Value)
            Case "PrintPreview"
                bPrintPreview = (GetString(Value) = "Y")
            Case "DefaultPrinter"
                bDefaultPrinter = (GetString(Value) = "Y")
            Case "DataParameter"
                sDataParameter = GetString(Value)
            Case Else
                Publics.MessageOut(Name & " property is not supported by Report object")
        End Select
    End Sub
End Class

Public Class Report
    Inherits ShellObject

    'todo rha
    ' report type proportional/fixed
    ' report widow/orphan control of split row
    ' report parameter display (using fields as for dialog)
    ' column widths/overflow to new row (not proportional)
    ' column locate to allow auto new line progression
    ' column DisplayType Wrap/Truncate/Mark truncated/Parameter

    Private sDefn As ReportDefn

    '' sf.Trimming.EllipsisCharacter
    '' sf.FormatFlags = StringFormatFlags.NoWrap

    Private datat As DataTable
    Dim hf As New Font("Arial", 8, FontStyle.Bold Or FontStyle.Underline)
    Dim pf As New Font("Arial", 8, FontStyle.Bold)
    Dim df As New Font("Arial", 8, FontStyle.Regular)
    Dim wf As New Font("Times New Roman", 200, FontStyle.Bold)

    Private g As Graphics
    Private yPos As Single
    Private iPageNumber As Integer = 1  'Page that is currently printing
    Private iSection As Integer = 0     'Current report section being printed
    Private iNextRow As Integer         'Next data row to print
    Private sTitle As String
    Private WithEvents pd As New PrintDocument
    Private fWidths() As Single
    Private fFormats() As StringFormat
    Private headers() As String

    Public Sub New(ByVal Defn As ReportDefn)
        sDefn = Defn
        sDefn.Parms.Clone(MyBase.Parms)
    End Sub

    Public Overrides Sub Update(ByVal Parms As ShellParameters)
        Dim i As Integer
        Try
            Me.Parms.MergeValues(Parms)
            datat = CType(Me.Parms.Item(sDefn.DataParameter).Value, DataTable)

            pd.DocumentName = sDefn.Title
            sTitle = sDefn.Title
            iPageNumber = 1
            iSection = 0
            iNextRow = 0
            i = 0
            pd.DefaultPageSettings.Landscape = False
            For Each f As Field In sDefn.Fields
                If Not datat.Columns.Item(f.Name) Is Nothing Then
                    i += f.DisplayWidth
                    If i > 790 Then
                        pd.DefaultPageSettings.Landscape = True
                        Exit For
                    End If
                End If
            Next

            initdata()
            If sDefn.PrintPreview Then
                Dim Pre As New PrintPreviewDialog

                Pre.Document = pd
                Pre.WindowState = FormWindowState.Maximized
                Pre.ShowDialog()
            Else
                If Not sDefn.DefaultPrinter Then
                    Dim dialog As New PrintDialog
                    dialog.Document = pd
                    Dim result As DialogResult = dialog.ShowDialog()
                    ' If the result is OK then print the document.
                    If (result = DialogResult.OK) Then
                        pd.Print()
                    End If
                Else
                    pd.Print()
                End If
            End If
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

    Private Sub initdata()
        Dim cols As Integer = sDefn.Fields.count - 1
        Dim totalWidth As Single
        ReDim fWidths(cols)
        ReDim fFormats(cols)
        ReDim headers(cols)

        'Print the column headers array.
        Dim i As Integer = 0
        For Each f As Field In sDefn.Fields
            If Not datat.Columns.Item(f.Name) Is Nothing And f.DisplayType <> "P" Then
                fWidths(i) = f.DisplayWidth
                If f.DisplayType = "F" Then
                    headers(i) = datat.Columns.Item(f.Name).Caption
                Else
                    headers(i) = f.Label
                End If
                fFormats(i) = New StringFormat
                Select Case f.Justify
                    Case "C"
                        fFormats(i).Alignment = StringAlignment.Center
                    Case "R"
                        fFormats(i).Alignment = StringAlignment.Far
                    Case Else
                        If f.ValueType = DbType.Currency Or _
                           f.ValueType = DbType.Double Or _
                           f.ValueType = DbType.Int32 Or _
                           f.ValueType = DbType.Int64 Then
                            fFormats(i).Alignment = StringAlignment.Far
                        Else
                            fFormats(i).Alignment = StringAlignment.Near
                        End If
                End Select
                i += 1
            End If
        Next

        totalWidth = 0
        For i = 0 To cols
            totalWidth += fWidths(i)
        Next
        For i = 0 To cols
            fWidths(i) /= totalWidth
        Next
    End Sub

    Private Sub PrintPage(ByVal sender As Object, _
                ByVal e As PrintPageEventArgs) Handles pd.PrintPage
        'Print the current page of this info (header info and/or grid)
        g = e.Graphics
        Dim sf As New StringFormat
        Dim sh As String

        ' print watermark on non production reports.
        If Not LCase(Publics.GetVariable("Production")) = "y" Then
            Dim layout As New StringFormat
            Dim x, y As Single
            If pd.DefaultPageSettings.Landscape Then
                x = 100
                y = 200
            Else
                layout.FormatFlags = StringFormatFlags.DirectionVertical
                x = 200
                y = 100
            End If
            g.DrawString("TEST", wf, Brushes.LightGray, x, y, layout)
        End If

        'Print header information before trying to print the grid.

        yPos = g.VisibleClipBounds.Top
        Dim s As Single = g.VisibleClipBounds.Width / 4

        Dim layoutRect As New RectangleF(0, yPos, s, 0)
        sf.Alignment = StringAlignment.Near
        sh = Publics.GetVariable("SystemName") & " - " & DateTime.Now.ToString()
        g.DrawString(sh, hf, Brushes.Black, layoutRect, sf)
        layoutRect = New RectangleF(s, yPos, s * 2, 0)
        sf.Alignment = StringAlignment.Center
        g.DrawString(sTitle, hf, Brushes.Black, layoutRect, sf)
        layoutRect = New RectangleF(s * 3, yPos, s, 0)
        sf.Alignment = StringAlignment.Far
        g.DrawString("Page: " & iPageNumber.ToString, hf, _
                                            Brushes.Black, layoutRect, sf)
        yPos += 20

        If iSection = 0 Then
            If Not PrintParam() Then
                iPageNumber += 1
                e.HasMorePages = True
            Else
                iSection = 1
            End If
        End If

        If iSection = 1 Then
            'Now print this page of the grid.
            If PrintData() Then
                e.HasMorePages = False
                iPageNumber = 1
                iSection = 0
                iNextRow = 0
            Else
                iPageNumber += 1
                e.HasMorePages = True
            End If
        End If
    End Sub

    Private Function PrintParam() As Boolean
        Dim f As Field
        Dim layoutC As RectangleF
        Dim sfC As New StringFormat
        Dim szC As SizeF
        Dim layoutP As RectangleF
        Dim sfP As New StringFormat
        Dim szP As SizeF
        Dim h As Single
        Dim layout As SizeF
        Dim s As String

        'todo rha add parameter listing here...

        For Each p As shellParameter In Me.Parms
            f = sDefn.Fields.Item(p.Name)
            If Not f Is Nothing And p.Initialised Then
                If f.DisplayType = "P" Then
                    layoutC = New RectangleF(0, yPos, f.LabelWidth, 0)
                    sfC.Alignment = StringAlignment.Far
                    layoutP = New RectangleF(f.LabelWidth + 20, yPos, f.DisplayWidth, 0)
                    sfP.Alignment = StringAlignment.Near

                    If p.Value Is Nothing Then
                        s = f.NullText
                    Else
                        If f.Format <> "" Then
                            s = Format(p.Value, f.Format)
                        Else
                            If f.ValueType = DbType.Date Then
                                s = Format(p.Value, "d-MMM-yyyy")
                            Else
                                s = p.Value.ToString
                            End If
                        End If
                    End If

                    '' Determine the size required to print this string
                    layout = New SizeF(f.LabelWidth, 0)
                    szC = g.MeasureString(f.Label, pf, layout, sfC)
                    layout = New SizeF(f.DisplayWidth, 0)
                    szP = g.MeasureString(s, pf, layout, sfP)
                    If szC.Height > szP.Height Then
                        h = szC.Height      'Keep track of tallest column
                    Else
                        h = szP.Height
                    End If

                    g.DrawString(f.Label, pf, Brushes.Black, layoutC, sfC)

                    If p.Value Is Nothing Then
                        g.DrawString(f.NullText, pf, Brushes.Black, layoutP, sfP)
                    Else
                        g.DrawString(s, pf, Brushes.Black, layoutP, sfP)
                    End If

                    yPos += h
                End If
            End If
        Next
        Return True
    End Function

    Private Function PrintData() As Boolean
        'Print this page to the printer.

        'Use the grid's tablestyle to determine the relative widths of the columns.
        Dim i As Integer

        Dim data(sDefn.Fields.count - 1) As String

        If Not PrintLine(headers, hf) Then
            Return False
        End If

        Dim dv As DataView = datat.DefaultView
        Dim drv As DataRowView
        'Print all the column values for this page row by row
        Do While iNextRow < datat.Rows.Count
            drv = dv.Item(iNextRow)

            i = 0
            For Each f As Field In sDefn.Fields
                If Not datat.Columns.Item(f.Name) Is Nothing And f.DisplayType <> "P" Then
                    If IsDBNull(drv(f.Name)) Then
                        data(i) = f.NullText
                    Else
                        If f.Format <> "" Then
                            data(i) = Format(drv(f.Name), f.Format)
                        Else
                            If f.ValueType = DbType.Date Then
                                data(i) = Format(drv(f.Name), "d-MMM-yyyy")
                            Else
                                data(i) = drv(f.Name).ToString
                            End If
                        End If
                    End If
                    i += 1
                End If
            Next
            If Not PrintLine(data, df) Then
                Return False
            End If
            iNextRow += 1
        Loop
        Return True
    End Function

    Private Function PrintLine(ByVal s() As String, ByVal f As Font) As Boolean
        'Print this array of strings to the printer. Each field's width is 
        'specified by the corresponding index in the fWidths array.

        Dim h As Single = 0
        Dim sz As SizeF
        Dim layout As SizeF
        Dim w As Single
        Dim width As Single = g.VisibleClipBounds.Width

        For i As Integer = 0 To s.GetUpperBound(0)
            layout = New SizeF(fWidths(i) * width, 0)
            ''Determine the size required to print this string
            If fFormats(i) Is Nothing Then
                sz = g.MeasureString(s(i), f, layout)
            Else
                sz = g.MeasureString(s(i), f, layout, fFormats(i))
            End If

            If sz.Height > h Then
                h = sz.Height      'Keep track of tallest column
                If yPos + h > g.VisibleClipBounds.Height Then
                    Return False
                End If
            End If
        Next

        Dim x As Single = 0
        For i As Integer = 0 To s.GetUpperBound(0)
            w = fWidths(i) * width

            Dim lRect As New RectangleF(x, yPos, w, h)
            If fFormats(i) Is Nothing Then
                g.DrawString(s(i), f, Brushes.Black, lRect)
            Else
                g.DrawString(s(i), f, Brushes.Black, lRect, fFormats(i))
            End If
            x += w
        Next
        yPos += h
        Return True
    End Function
End Class
