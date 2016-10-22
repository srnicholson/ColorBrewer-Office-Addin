Imports System.Runtime.InteropServices
Imports System.Data
Imports Extensibility
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports System.Reflection
Imports System.Drawing
Imports System.IO

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the ColorBrewerAddinSetup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("017F85B5-99D0-4630-8371-80CD3F1D0324"), ProgIdAttribute("ColorBrewerAddin.Connect")> _
Public Class Addin

    Implements Extensibility.IDTExtensibility2, IRibbonExtensibility

    Private applicationObject As Object
    Private addInInstance As Object
    Private appExcel As Excel.Application
    Private appWord As Word.Application
    Private appPowerPoint As PowerPoint.Application
    Dim PalettesDataSet As New DataSet
    Dim PalettesDataTable As System.Data.DataTable
    Dim objShell As Object

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
    End Sub

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection
        applicationObject = application
        addInInstance = addInInst
        ''Load Palettes XML to datatable
        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim reader As New System.IO.StreamReader(thisAssembly.GetManifestResourceStream(thisAssembly.GetName.Name + ".Palettes.xml"))
        Try
            PalettesDataSet.ReadXml(reader)
            reader.Close()
            PalettesDataTable = PalettesDataSet.Tables(0)
        Catch e As Exception
            MsgBox("Unspecified Error.")
        End Try
    End Sub

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return ReadString("RibbonUI.xml")
    End Function

    Public Sub OnAction(ByVal control As IRibbonControl, Optional ByVal PalId As Integer = 0)
        Try
            Select Case control.Id
                Case "About"
                    MsgBox("Thanks for trying out the ColorBrewer Office Add-in!" _
                           & vbNewLine & vbNewLine & "Originally designed for use in cartography and GIS, the ColorBrewer project was developed by Cynthia Brewer, Professor of Geography at Penn State University.  R users may recognize these palettes, as they are employed by Hadley Wickham's popular ggplot2 (via the RColorBrewer package created by Erich Neuwirth)." _
                           & vbNewLine & vbNewLine & "The goal of developing this add-in was to provide an easy way of using these palettes in Office charts, thereby enabling users to quickly venture beyond the default options." _
                           & vbNewLine & vbNewLine & "For more information, please visit the GitHub repository or the ColorBrewer website via the provided links." _
                           & vbNewLine & vbNewLine & vbNewLine & "v6.0 ColorBrewer Office Add-in developed by Scott Nicholson.")
                Case "Palettes"
                    Select Case applicationObject.Name.ToString
                        Case "Microsoft Excel"
                            appExcel = applicationObject
                            Call Excel_Sub(PalId, False)
                        Case "Microsoft Word"
                            appWord = applicationObject
                            Call Word_Sub(PalId, False)
                        Case "Microsoft PowerPoint"
                            appPowerPoint = applicationObject
                            Call PowerPoint_Sub(PalId, False)
                        Case Else
                            MsgBox("Error: This Office application is not supported.")
                    End Select
                Case "Reverse_color_order"
                    Select Case applicationObject.Name.ToString
                        Case "Microsoft Excel"
                            appExcel = applicationObject
                            Call Excel_Sub(PalId, True)
                        Case "Microsoft Word"
                            appWord = applicationObject
                            Call Word_Sub(PalId, True)
                        Case "Microsoft PowerPoint"
                            appPowerPoint = applicationObject
                            Call PowerPoint_Sub(PalId, True)
                        Case Else
                            MsgBox("Error: This Office application is not supported.")
                    End Select
                Case "Help"
                    MsgBox("Thanks for trying out the ColorBrewer Office Add-in!" _
                           & vbNewLine & vbNewLine & vbNewLine & "To get started, select a chart and then choose either:" _
                           & vbNewLine & vbNewLine & "•""Choose a Palette"" to change the chart's color scheme, or" _
                           & vbNewLine & vbNewLine & "•""Reverse Color Order"" to reverse the chart's existing color scheme." _
                           & vbNewLine & vbNewLine & vbNewLine & "For more help, check out the GitHub page by clicking the ""GitHub"" button.")
                Case "GitHub"
                    objShell = CreateObject("Wscript.Shell")
                    objShell.Run("https://github.com/srnicholson/ColorBrewer-Office-Addin/")
                Case "ColorBrewerWebsite"
                    objShell = CreateObject("Wscript.Shell")
                    objShell.Run("http://colorbrewer2.org/")
                Case Else
                    MsgBox("Unkown Control Id: " + control.Id, , "ColorBrewer Office Addin")
            End Select

        Catch throwedException As Exception
            MsgBox("Error: Unexpected state in ColorBrewer OnAction" + vbNewLine + "Error details: " + throwedException.Message)
        End Try
    End Sub
#End Region

#Region "ColorBrewer Methods"

    Public Sub Excel_Sub(ByVal PalId As Integer, ByVal reverse As Boolean)
        Dim chart As Excel.Chart
        Dim color_name As String

        color_name = PaletteID2Name(PalId)

        If appExcel.ActiveChart Is Nothing Then
            MsgBox("Error: No chart selected.")
            Exit Sub
        End If

        Try
            chart = appExcel.ActiveChart
            Call ColorBrewerFill(chart, color_name, reverse)
        Catch
            MsgBox("Unspecified Error.")
        End Try

    End Sub

    Public Sub Word_Sub(ByVal PalId As Integer, ByVal reverse As Boolean)
        Dim inline As Word.InlineShape
        Dim shape As Word.Shape
        Dim chart As Word.Chart
        Dim color_name As String

        Try
            color_name = PaletteID2Name(PalId)
            With appWord.ActiveWindow.Selection
                'Determine if the selection is a regular shape or an inline shape (or not a chart)
                If .Type = 7 Then
                    inline = .InlineShapes(1)
                    shape = inline.ConvertToShape()
                ElseIf .Type = 8 Then
                    shape = .ShapeRange(1)
                Else
                    MsgBox("Error: No chart selected.")
                    Exit Sub
                End If
            End With

            chart = shape.ConvertToInlineShape().Chart

            Call ColorBrewerFill(chart, color_name, reverse)

        Catch e As Exception
            MsgBox("Unknown Error.")
        End Try
    End Sub

    Public Sub PowerPoint_Sub(ByVal PalId As Integer, ByVal reverse As Boolean)
        Dim chart As PowerPoint.Chart
        Dim color_name As String

        color_name = PaletteID2Name(PalId)
        Try
            chart = appPowerPoint.ActiveWindow.Selection.ShapeRange(1).Chart
        Catch
            MsgBox("Error: No chart selected.")
            Exit Sub
        End Try

        Try
            Call ColorBrewerFill(chart, color_name, reverse)
        Catch
            MsgBox("Unspecified Error.")
        End Try
    End Sub

    Sub ColorBrewerFill(ByRef chart As Object, ByVal pal As String, ByVal reverse As Boolean)
        Dim palette As Array
        Dim series_count As Integer
        Dim rgb_color As Long
        Dim i As Integer
        Dim old_colors As New ArrayList

        With chart
            series_count = .SeriesCollection.Count
            Select Case CType(.ChartType, XlChartType)
                'Chart types enumerated here: https://msdn.microsoft.com/en-us/library/office/ff838409.aspx
                Case XlChartType.xlXYScatter, XlChartType.xlXYScatterLines, XlChartType.xlXYScatterSmooth, XlChartType.xlLineMarkers, XlChartType.xlLineMarkersStacked, XlChartType.xlLineMarkersStacked100, XlChartType.xlRadarMarkers
                    'Points, Lines optional Case
                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
                        If BlankPalette(palette) Then Exit Sub
                    End If

                    old_colors = GetChartRGBs(chart, XlChartType.xlXYScatter)

                    For i = 1 To series_count
                        If reverse Then
                            rgb_color = CType(old_colors(series_count - i), Long)
                        Else
                            rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                        End If
                        With .SeriesCollection(i)
                            'MsgBox("Changing color: " & old_colors(i - 1) & " in series " & i & ".")
                            .MarkerForegroundColor = rgb_color
                            .MarkerBackgroundColor = rgb_color
                            If .Format.Line.Visible = True Then
                                .Format.Line.ForeColor.RGB = rgb_color
                            End If
                        End With
                    Next
                Case XlChartType.xlLine, XlChartType.xlLineStacked, XlChartType.xlRadar, XlChartType.xlXYScatterLinesNoMarkers, XlChartType.xlXYScatterSmoothNoMarkers, XlChartType.xlLineStacked100
                    'Line Only Case
                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
                        If BlankPalette(palette) Then Exit Sub
                    Else
                        old_colors = GetChartRGBs(chart, XlChartType.xlLine)
                    End If
                    For i = 1 To series_count
                        'MsgBox("Changing color: " & old_colors(i - 1) & " in series " & i & ".")
                        If reverse Then
                            rgb_color = old_colors(series_count - i)
                        Else
                            rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                        End If
                        With .SeriesCollection(i)
                            .Format.Line.Visible = False
                            .Format.Line.Visible = True
                            .Format.Line.ForeColor.RGB = rgb_color
                        End With
                    Next
                Case XlChartType.xl3DArea, XlChartType.xl3DAreaStacked, XlChartType.xl3DAreaStacked100, _
                    XlChartType.xl3DBarClustered, XlChartType.xl3DBarStacked, XlChartType.xl3DBarStacked100, _
                    XlChartType.xlBubble, XlChartType.xlBubble3DEffect, XlChartType.xl3DColumn, _
                    XlChartType.xl3DLine, XlChartType.xl3DColumnClustered, XlChartType.xl3DColumnStacked, _
                    XlChartType.xl3DColumnStacked100, XlChartType.xlArea, XlChartType.xlAreaStacked, _
                    XlChartType.xlAreaStacked100, XlChartType.xlBarClustered, XlChartType.xlBarStacked, _
                    XlChartType.xlBarStacked100, XlChartType.xlColumnClustered, XlChartType.xlColumnStacked, _
                    XlChartType.xlColumnStacked100, XlChartType.xlConeBarClustered, XlChartType.xlConeBarStacked, _
                    XlChartType.xlConeBarStacked100, XlChartType.xlConeCol, XlChartType.xlConeColClustered, _
                    XlChartType.xlConeColStacked, XlChartType.xlConeColStacked100, XlChartType.xlCylinderBarClustered, _
                    XlChartType.xlCylinderBarStacked, XlChartType.xlCylinderBarStacked100, XlChartType.xlCylinderCol, _
                    XlChartType.xlCylinderColClustered, XlChartType.xlCylinderColStacked, XlChartType.xlCylinderColStacked100, _
                    XlChartType.xlPyramidBarClustered, XlChartType.xlPyramidBarStacked, XlChartType.xlPyramidBarStacked100, _
                    XlChartType.xlPyramidCol, XlChartType.xlPyramidColClustered, XlChartType.xlPyramidColStacked, _
                    XlChartType.xlPyramidColStacked100, XlChartType.xlRadarFilled

                    'Area Case

                    Dim old_spacing As String
                    'prevent column spacing from changing during color change
                    old_spacing = .ChartGroups(1).Overlap

                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
                        If BlankPalette(palette) Then Exit Sub
                    Else
                        old_colors = GetChartRGBs(chart, XlChartType.xlColumnClustered)
                    End If

                    For i = 1 To series_count
                        'MsgBox("Changing color: " & old_colors(i - 1) & " in series " & i & ".")
                        If reverse Then
                            rgb_color = old_colors(series_count - i)
                        Else
                            rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                        End If
                        With .SeriesCollection(i)
                            .Interior.Color = rgb_color
                            '.Border.Color = rgb_color
                        End With
                    Next

                    'prevent column spacing from changing during color change
                    .ChartGroups(1).Overlap = old_spacing

                Case XlChartType.xl3DPie, XlChartType.xl3DPieExploded, XlChartType.xlDoughnut, XlChartType.xlDoughnutExploded, XlChartType.xlPie
                    'Pie Case
                    Dim j As Integer
                    Dim counter As Integer = 0
                    If reverse Then
                        old_colors = GetChartRGBs(chart, XlChartType.xlPie)
                    End If
                    For i = 1 To series_count
                        With .SeriesCollection(i)
                            If Not reverse Then
                                palette = GetPaletteData(pal, .Points.Count)
                                If BlankPalette(palette) Then Exit Sub
                            End If

                            For j = 1 To .Points.Count
                                If reverse Then
                                    rgb_color = old_colors(.Points.Count * series_count - counter - 1)
                                    counter = counter + 1
                                Else
                                    rgb_color = RGB(palette(j - 1)(2), palette(j - 1)(3), palette(j - 1)(4))
                                End If
                                With .Points(j)
                                    .Interior.Color = rgb_color
                                    '.Border.Color = rgb_color
                                End With
                            Next
                        End With
                    Next
                Case Else
                    MsgBox("Error: Graph type not supported.", vbOKOnly)
            End Select
        End With
    End Sub

    Private Function GetPaletteData(pal As String, NumColors As Integer) As Array
        Dim filter As String
        filter = "[C] = '" + pal + "' AND [N] = '" + NumColors.ToString + "'"
        Try
            Return PalettesDataTable.Select(filter)
        Catch e As Exception
            MsgBox("Error: Invalid GetPaletteData function query")
            Return Nothing
        End Try
    End Function

    Private Function BlankPalette(palette As Array) As Boolean
        If palette.Length = 0 Then
            MsgBox("Error: The chart's series count is outside the range for this palette." & vbNewLine & _
                   "Try a different palette or change the number of series in the chart." & vbNewLine & vbNewLine & _
                   "Tip: Valid ranges for each palette are listed in the 'Choose a Palette' drop-down menu.")
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetChartRGBs(ByVal chart As Object, ByVal type As XlChartType) As ArrayList
        'Returns ArrayList of RGB (BGR?) values corresponding to each series in the chart
        'Workaround for getting automatic colors for non-column charts is based on the
        'clever solution by David Zemens on Stack Overflow here: http://stackoverflow.com/a/25826428
        Dim chtType As Long
        Dim colors As New ArrayList
        Dim fill_value As Long
        Dim counter As Integer

        chtType = chart.ChartType

        'Select correct SeriesCollection fill value based on xlChartType
        Select Case type
            Case XlChartType.xlXYScatter
                fill_value = chart.SeriesCollection(1).MarkerForegroundColor
            Case XlChartType.xlColumnClustered
                fill_value = chart.SeriesCollection(1).Format.Fill.ForeColor.RGB
            Case XlChartType.xlLine
                If chart.SeriesCollection(1).Format.Line.ForeColor.RGB = 16777215 Then
                    '16777215 appears to be what automatic line color is in Office 2007
                    fill_value = -1
                Else
                    fill_value = chart.SeriesCollection(1).Format.Line.ForeColor.RGB
                End If
            Case XlChartType.xlPie
                fill_value = chart.SeriesCollection(1).Points(1).Interior.Color
            Case Else
                fill_value = 9999 'Throw error
        End Select
        'ONLY changes to column plot IF the series fill type is automatic
        'Otherwise, custom colors (such as from a previous ColorBrewer run) will be lost.
        If fill_value <= 0 Then
            'Temporarily change chart type to "column" in order to extract automatic RGB values
            chart.ChartType = 51
            For Each srs In chart.SeriesCollection
                colors.Add(srs.Format.Fill.ForeColor.RGB)
            Next
        Else
            Select Case type
                Case XlChartType.xlXYScatter
                    For Each srs In chart.SeriesCollection
                        colors.Add(srs.MarkerForegroundColor)
                    Next
                Case XlChartType.xlColumnClustered
                    For Each srs In chart.SeriesCollection
                        colors.Add(srs.Format.Fill.ForeColor.RGB)
                    Next
                Case XlChartType.xlLine
                    For Each srs In chart.SeriesCollection
                        colors.Add(srs.Format.Line.ForeColor.RGB)
                    Next
                Case XlChartType.xlPie
                    For Each srs In chart.SeriesCollection
                        For Each point In srs.Points
                            colors.Add(point.Interior.Color)
                        Next
                    Next
                Case Else
                    MsgBox("Error: unable to extract original series colors")
            End Select
        End If

        chart.ChartType = chtType

        Return colors
    End Function

#End Region

#Region "XML Methods"
    ' Modified from https://github.com/NetOfficeFw/NetOffice/blob/master/Examples/Misc/VB/COMAddin%20Examples/SuperAddin/Addin.vb#L192
    Private Shared Function ReadString(ByVal fileName As String) As String

        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim resourceStream As Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + fileName)
        If (IsNothing(resourceStream)) Then
            Throw (New IOException("Error accessing resource Stream."))
        End If

        Dim textStreamReader As StreamReader = New StreamReader(resourceStream)
        If (IsNothing(textStreamReader)) Then
            Throw (New IOException("Error accessing resource File."))
        End If

        Dim text As String = textStreamReader.ReadToEnd()
        resourceStream.Close()
        textStreamReader.Close()
        Return text

    End Function
#End Region

#Region "Gallery Callbacks"
    Private itemCount As Integer = 35 ' Used with GetItemCount.
    Private itemHeight As Integer = 22 ' Used with GetItemHeight.
    Private itemWidth As Integer = 182 ' Used with GetItemWidth.
    Public Function LoadImage(ByVal imageName As String) As Bitmap
        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim stream As Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + imageName)
        Return New Bitmap(stream)
    End Function
    Public Function GetLabel(ByVal control As IRibbonControl) As String
        Dim strText As String
        Select Case control.Id
            Case "Palettes" : strText = "Choose a Palette:"
        End Select
        Return strText
    End Function
    Public Function GetShowImage(ByVal control As IRibbonControl) As Boolean
        Return True
    End Function
    Public Function GetShowLabel(ByVal control As IRibbonControl) As Boolean
        Return True
    End Function
    Public Function GetItemImage(ByVal control As IRibbonControl, ByVal itemIndex As Integer) As Bitmap
        Dim imageName As String
        imageName = PaletteID2Name(itemIndex) & ".png"
        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim stream As Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + imageName)
        Return New Bitmap(stream)
    End Function
    Public Function GetSize(ByVal control As IRibbonControl) As RibbonControlSize
        Select Case control.Id
            Case "Palettes" : Return RibbonControlSize.RibbonControlSizeLarge
        End Select
    End Function
    Public Function GetEnabled(ByVal control As IRibbonControl) As Boolean
        Select Case control.Id
            Case "Palettes"
                Return True
            Case Else
                Return True
        End Select
    End Function
    Public Function GetItemCount(ByVal control As IRibbonControl) As Integer
        Return itemCount
    End Function
    Public Function getItemHeight(ByVal control As IRibbonControl) As Integer
        Return itemHeight
    End Function
    Public Function getItemWidth(ByVal control As IRibbonControl) As Integer
        Return itemWidth
    End Function
    Public Function getItemLabel(ByVal control As IRibbonControl, ByVal id As String) As String
        Select Case id
            Case 0 : Return "Accent (3 - 8)"
            Case 1 : Return "Blues (3 - 9)"
            Case 2 : Return "BrBG (3 - 11)"
            Case 3 : Return "BuGn (3 - 9)"
            Case 4 : Return "BuPu (3 - 9)"
            Case 5 : Return "Dark2 (3 - 8)"
            Case 6 : Return "GnBu (3 - 9)"
            Case 7 : Return "Greens (3 - 9)"
            Case 8 : Return "Greys (3 - 9)"
            Case 9 : Return "Oranges (3 - 9)"
            Case 10 : Return "OrRd (3 - 9)"
            Case 11 : Return "Paired (3 - 12)"
            Case 12 : Return "Pastel1 (3 - 9)"
            Case 13 : Return "Pastel2 (3 - 8)"
            Case 14 : Return "PiYG (3 - 11)"
            Case 15 : Return "PRGn (3 - 11)"
            Case 16 : Return "PuBu (3 - 9)"
            Case 17 : Return "PuBuGn (3 - 9)"
            Case 18 : Return "PuOr (3 - 11)"
            Case 19 : Return "PuRd (3 - 9)"
            Case 20 : Return "Purples (3 - 9)"
            Case 21 : Return "RdBu (3 - 11)"
            Case 22 : Return "RdGy (3 - 11)"
            Case 23 : Return "RdPu (3 - 9)"
            Case 24 : Return "Reds (3 - 9)"
            Case 25 : Return "RdYlBu (3 - 11)"
            Case 26 : Return "RdYlGn (3 - 11)"
            Case 27 : Return "Set1 (3 - 9)"
            Case 28 : Return "Set2 (3 - 8)"
            Case 29 : Return "Set3 (3 - 12)"
            Case 30 : Return "Spectral (3 - 11)"
            Case 31 : Return "YlGn (3 - 9)"
            Case 32 : Return "YlGnBu (3 - 9)"
            Case 33 : Return "YlOrBr (3 - 9)"
            Case 34 : Return "YlOrRd (3 - 9)"
        End Select
    End Function
    Function PaletteID2Name(index As Integer) As String
        Select Case index
            Case 0 : Return "Accent"
            Case 1 : Return "Blues"
            Case 2 : Return "BrBG"
            Case 3 : Return "BuGn"
            Case 4 : Return "BuPu"
            Case 5 : Return "Dark2"
            Case 6 : Return "GnBu"
            Case 7 : Return "Greens"
            Case 8 : Return "Greys"
            Case 9 : Return "Oranges"
            Case 10 : Return "OrRd"
            Case 11 : Return "Paired"
            Case 12 : Return "Pastel1"
            Case 13 : Return "Pastel2"
            Case 14 : Return "PiYG"
            Case 15 : Return "PRGn"
            Case 16 : Return "PuBu"
            Case 17 : Return "PuBuGn"
            Case 18 : Return "PuOr"
            Case 19 : Return "PuRd"
            Case 20 : Return "Purples"
            Case 21 : Return "RdBu"
            Case 22 : Return "RdGy"
            Case 23 : Return "RdPu"
            Case 24 : Return "Reds"
            Case 25 : Return "RdYlBu"
            Case 26 : Return "RdYlGn"
            Case 27 : Return "Set1"
            Case 28 : Return "Set2"
            Case 29 : Return "Set3"
            Case 30 : Return "Spectral"
            Case 31 : Return "YlGn"
            Case 32 : Return "YlGnBu"
            Case 33 : Return "YlOrBr"
            Case 34 : Return "YlOrRd"
        End Select
    End Function
    Public Function GetItemScreenTip(ByVal control As IRibbonControl, ByVal index As Integer) As String
    End Function
    Public Function GetItemSuperTip(ByVal control As IRibbonControl, ByVal index As Integer) As String
    End Function
    Public Function GetKeyTip(ByVal control As IRibbonControl) As String
    End Function
    Public Function GetSuperTip(ByVal control As IRibbonControl) As String
        Select Case control.Id
            Case "About" : Return "Click to learn more about the ColorBrewer Add-in."
            Case "Palettes" : Return "Click to open the palette gallery."
            Case "Reverse_color_order" : Return "Click to reverse the color order of the selected chart."
            Case "Help" : Return "Click for help using the ColorBrewer Add-in."
            Case "GitHub" : Return "Click to go to the ColorBrewer Add-in code repository on GitHub." + vbCrLf + vbCrLf + "https://github.com/srnicholson/ColorBrewer-Office-Addin/"
            Case "ColorBrewerWebsite" : Return "Click to go to the ColorBrewer website." + vbCrLf + vbCrLf + "http://colorbrewer2.org/"
        End Select
    End Function
    Public Sub galleryOnAction(ByVal control As IRibbonControl, ByVal selectedId As String, _
    ByVal selectedIndex As Integer)
        OnAction(control, selectedIndex)
    End Sub
#End Region

End Class