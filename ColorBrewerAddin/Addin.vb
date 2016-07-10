Imports System.Runtime.InteropServices
Imports System.Data
Imports Extensibility
Imports Microsoft.Office.Core
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
    Dim PalettesDataSet As New DataSet
    Dim PalettesDataTable As DataTable

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
            MsgBox(e.ToString)
        End Try
    End Sub
#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return ReadString("RibbonUI.xml")
    End Function

    Public Sub OnAction(ByVal control As IRibbonControl, Optional ByVal PalId As Integer = 0)
        applicationObject.ScreenUpdating = False 'TO DO: Make sure this doesn't cause problems on errors
        Try
            Select Case control.Id
                Case "About"
                    MsgBox("Here is some information about the ColorBrewer Add-in.")
                Case "Palettes"
                    Select Case applicationObject.Name.ToString
                        Case "Microsoft Excel"
                            Call Excel_Sub(PalId, False)
                        Case "Microsoft Word"
                            Call Word_Sub(PalId, False)
                        Case "Microsoft PowerPoint"
                            Call PowerPoint_Sub(PalId, False)
                        Case Else
                            MsgBox("Error: This Office application is not supported.")
                    End Select
                Case "Reverse_color_order"
                    Select Case applicationObject.Name.ToString
                        Case "Microsoft Excel"
                            Call Excel_Sub(PalId, True)
                        Case "Microsoft Word"
                            Call Word_Sub(PalId, True)
                        Case "Microsoft PowerPoint"
                            Call PowerPoint_Sub(PalId, True)
                        Case Else
                            MsgBox("Error: This Office application is not supported.")
                    End Select
                Case "Undo"
                    MsgBox("This is the undo button.")
                Case "Help"
                    MsgBox("This is the help button.")
                Case Else
                    MsgBox("Unkown Control Id: " + control.Id, , "ColorBrewer Office Addin")
            End Select

        Catch throwedException As Exception
            applicationObject.ScreenUpdating = True
            MsgBox("Error: Unexpected state in ColorBrewer OnAction" + vbNewLine + "Error details: " + throwedException.Message)

        End Try
        applicationObject.ScreenUpdating = True
    End Sub
#End Region

#Region "ColorBrewer Methods"

    Public Sub Excel_Sub(PalId As Integer, reverse As Boolean)
        Dim chart As Object
        Dim chart_type As String
        Dim series_count As Integer
        Dim color_name As String
        Dim pos_top, pos_left As Long

        Try
            color_name = PaletteID2SName(PalId)
            chart = applicationObject.ActiveChart

            pos_top = chart.Parent.Top
            pos_left = chart.Parent.Left

            chart.Parent.Copy()

            'Hide the old chart
            chart.Parent.Visible = False 'TO DO: Figure out where the old charts end up and how we can get them back on an "undo"

            applicationObject.ActiveWindow.ActiveSheet.Paste()


            'set "chart" to point to the new copy
            With applicationObject.ActiveSheet
                chart = .ChartObjects(.ChartObjects.Count).Chart
            End With

            'Align copy with original
            chart.Parent.Top = pos_top
            chart.Parent.Left = pos_left

            chart_type = chart.ChartType
            series_count = chart.SeriesCollection.Count

            Try
                Call ColorBrewerFill(chart, color_name, reverse)
            Catch e As Exception
                'Note: This error message may not be repsentative of all types of failures.
                MsgBox(e.ToString)
                'MsgBox("Error: Data series count is outside this palette's range. Please choose a different palette or change the number of series.")
            End Try
        Catch e As Exception
            MsgBox(e.ToString)
            'MsgBox("No Chart Selected.")
        End Try

    End Sub

    Public Sub Word_Sub(PalId As Integer, reverse As Boolean)
        Dim chart, new_shape, new_chart As Object
        Dim chart_type As String
        Dim wrap_format, series_count As Integer
        Dim color_name As String
        Dim pos_top, pos_left As Long
        Dim IsShape As Boolean = False

        Try
            color_name = PaletteID2SName(PalId)
            With applicationObject.ActiveWindow.Selection
                'Determine if the selection is a regular shape or an inline shape
                If .Type = 7 Then
                    chart = .InlineShapes(1)
                    chart = chart.ConvertToShape()
                ElseIf .Type = 8 Then
                    chart = .ShapeRange(1)
                    wrap_format = chart.WrapFormat.Type
                    IsShape = True
                Else
                    MsgBox("No Chart Selected") 'TO DO: Better way to handle this error?
                End If
            End With

            pos_top = chart.Top
            pos_left = chart.Left

            new_shape = chart.Duplicate()

            chart.Visible = False 'TO DO: Figure out where the old charts end up and how we can get them back on an "undo"

            'Align copy with original
            new_shape.Top = pos_top
            new_shape.Left = pos_left

            new_chart = new_shape.ConvertToInlineShape().Chart

            chart_type = new_chart.ChartType
            series_count = new_chart.SeriesCollection.Count

            Call ColorBrewerFill(new_chart, color_name, reverse)

            new_chart.Select() 'This needs to be there; otherwise the program coudl crash if nothing is selected
            'TO DO: Alternative is to figure out way to disable buttons unless chart (shape or inlineshape) is selected

            If IsShape Then
                'Convert back to shape again and adjust wrap formatting
                new_chart.Parent.ConvertToShape()
                new_shape.WrapFormat.Type = wrap_format
            End If

        Catch e As Exception
            MsgBox(e.ToString)
            ''TO DO: Put these old error message is in the right spots.
            'MsgBox("Error: Data series count is outside this palette's range. Please choose a different palette or change the number of series.")
            'MsgBox("No Chart Selected")
        End Try
    End Sub

    Public Sub PowerPoint_Sub(PalId As Integer, reverse As Boolean)
        Dim chart As Object
        Dim chart_type As String
        Dim series_count As Integer
        Dim color_name As String

        Try
            color_name = PaletteID2SName(PalId)
            chart = applicationObject.ActiveWindow.Selection.ShapeRange(1).Chart
            chart_type = chart.ChartType
            series_count = chart.SeriesCollection.Count
            Try
                Call ColorBrewerFill(chart, color_name, reverse)
            Catch
                'Note: This error message may not be repsentative of all types of failures.
                MsgBox("Error: Data series count is outside this palette's range. Please choose a different palette or change the number of series.")
            End Try
        Catch e As Exception
            MsgBox(e.ToString)
            'MsgBox("No Chart Selected")
        End Try
    End Sub

    Function GetPaletteData(pal As String, NumColors As Integer) As Array
        Dim filter As String
        filter = "[C] = '" + pal + "' AND [N] = '" + NumColors.ToString + "'"
        Try
            Return PalettesDataTable.Select(filter)
        Catch e As Exception
            MsgBox(e.Message)
            Return PalettesDataTable.Select 'TODO: Make this an empty array
        End Try
    End Function

    Sub ColorBrewerFill(ByVal chart As Object, ByVal pal As String, ByVal reverse As Boolean)
        Dim palette As Array
        Dim series_count As Integer
        Dim rgb_color As Long
        Dim i As Integer
        Dim pos_top, pos_left As Double
        Dim old_colors As New ArrayList

        With chart
            series_count = .SeriesCollection.Count
            Select Case .ChartType
                'Chart types enumerated here: https://msdn.microsoft.com/en-us/library/office/ff838409.aspx
                Case XlChartType.xlXYScatter, XlChartType.xlXYScatterLines, XlChartType.xlXYScatterLinesNoMarkers, XlChartType.xlXYScatterSmooth, XlChartType.xlRadarMarkers
                    'Points, Lines optional Case
                    'TO DO: For scatterplots, change fill or line color depending on type of point (line-type vs shape type)
                    'Otherwise everything changes to squares. UPDATE: may not be possible due to unhelpful "Automatic" property-- need a way to return the actual MarkerStyle
                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
                    End If

                    old_colors = GetChartRGBs(chart, XlChartType.xlXYScatter)
                    For i = 1 To series_count
                        If reverse Then
                            rgb_color = old_colors(series_count - i)
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
                Case XlChartType.xlLine, XlChartType.xlRadar
                    'Line Only Case
                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
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
                Case XlChartType.xlColumnClustered, XlChartType.xlConeCol, XlChartType.xl3DArea, XlChartType.xlAreaStacked, XlChartType.xlAreaStacked100, XlChartType.xlBubble3DEffect, XlChartType.xlPyramidBarClustered, XlChartType.xlRadarFilled
                    Dim old_spacing As String
                    'prevent column spacing from changing during color change
                    old_spacing = .ChartGroups(1).Overlap

                    'Area Case
                    If Not reverse Then
                        palette = GetPaletteData(pal, series_count)
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
                            .Border.Color = rgb_color
                        End With
                    Next

                    'prevent column spacing from changing during color change
                    .ChartGroups(1).Overlap = old_spacing

                Case XlChartType.xlDoughnut, XlChartType.xlDoughnutExploded, XlChartType.xlPie
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
                            End If

                            For j = 1 To .Points.Count
                                'MsgBox("Changing color: " & old_colors((i * j) - 1) & " for series " & i & " and point " & j & ".")
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
                Case XlChartType.xlSurface
                    'Surface Case
                    If Not reverse Then
                        palette = GetPaletteData(pal, .Legend.LegendEntries.Count)
                    Else
                        old_colors = GetChartRGBs(chart, XlChartType.xlSurface)
                    End If

                    'TO DO: This "With" statement is application specific
                    With .Legend
                        'TODO: If Legend doesn't exist, display it temporarily to change colors
                        '.HasLegend = True

                        Debug.Print("current major unit =" & chart.Axes(2).MajorUnit) 'Errors if this statement is commented out...race condition?

                        For i = 1 To .LegendEntries.Count
                            'MsgBox("Changing color: " & old_colors(i - 1) & " in legend " & i & ".")
                            If reverse Then
                                rgb_color = old_colors(.LegendEntries.Count - i)
                            Else
                                rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                            End If
                            .LegendEntries(i).LegendKey.Interior.Color = rgb_color
                        Next
                    End With
                Case XlChartType.xlSurfaceWireframe, XlChartType.xlSurfaceTopViewWireframe
                    'Surface Wireframe Case
                    With .Legend
                        palette = GetPaletteData(pal, .LegendEntries.Count)
                        For i = 1 To .LegendEntries.Count
                            rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                            .LegendEntries(i).LegendKey.Border.Color = rgb_color
                        Next
                    End With
                Case Else
                    MsgBox("Error: Graph type not supported.", vbOKOnly)
            End Select
        End With
        'TO ADD:
        '1) Reset Function-- Store previous SeriesCollection RGB values to give option to undo macro
        '2) Option to exclude border coloring in Area plots
        '3) "Inverse" to get sequence in opposite direction
        '4) Option to install all palettes as xml Themes
    End Sub

    Private Function GetChartRGBs(ByVal chart As Object, ByVal type As XlChartType) As ArrayList
        'NOT FINISHED! (SEE BELOW)
        'Returns ArrayList of RGB (BGR?) values corresponding to each series in the chart
        'Based on the brilliant solution by David Zemens on Stack Overflow here: http://stackoverflow.com/a/25826428
        '''Dim temp_chart As Object
        '''Dim chart_index As Long
        Dim chtType As Long
        Dim colors As New ArrayList
        Dim fill_value As Long
        Dim counter As Integer

        ''''OLD METHOD'''''
        ''''create temporary chart
        '''chart_index = chart.Parent.Index
        '''chart.Parent.Copy()
        ''''This may need to vary based on Application (Excel v Word v Powerpoint)
        '''applicationObject.ActiveWindow.ActiveSheet.Paste()

        '''temp_chart = applicationObject.ActiveChart

        chtType = chart.ChartType
        colors.Clear()

        'Select correct SeriesCollection fill value based on xlChartType
        Select Case type
            Case XlChartType.xlXYScatter
                fill_value = chart.SeriesCollection(1).MarkerForegroundColor
            Case XlChartType.xlColumnClustered
                fill_value = chart.SeriesCollection(1).Format.Fill.ForeColor.RGB
            Case XlChartType.xlLine
                If chart.SeriesCollection(1).Format.Line.ForeColor.RGB = 16777215 Then
                    'This appears to be what automatic line color is in Office 2007
                    fill_value = -1
                Else
                    fill_value = chart.SeriesCollection(1).Format.Line.ForeColor.RGB
                End If
            Case XlChartType.xlPie
                fill_value = chart.SeriesCollection(1).Points(1).Interior.Color
            Case XlChartType.xlSurface
                fill_value = chart.Legend.LegendEntries(1).LegendKey.Interior.ColorIndex
            Case Else
                fill_value = 9999 '???
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
                Case XlChartType.xlSurface
                    For Each srs In chart.Legend.LegendEntries
                        colors.Add(srs.LegendKey.Interior.Color)
                    Next
                Case Else
                    MsgBox("Error: unable to extract old series colors")
            End Select
        End If

        chart.ChartType = chtType

        ''''delete the temporary chart
        '''temp_chart.Parent.Delete()
        ''''This may also need to vary based on Application (Excel v Word v Powerpoint)
        '''applicationObject.ActiveWindow.ActiveSheet.ChartObjects(chart_index).Activate()

        Return colors
    End Function

    'Function GetPosition(ByVal chart As Object) As ArrayList
    '    Dim coords As ArrayList
    '    MsgBox("Here!")
    '    Try
    '        With chart
    '            coords.Add(.Parent.Top)
    '            coords.Add(.Parent.Left)
    '        End With
    '    Catch ex As Exception
    '        MsgBox("Error in GetPosition function")
    '    End Try

    '    Return coords
    'End Function

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
        imageName = PaletteID2SName(itemIndex) & ".png"
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
            Case "Undo"
                Return False
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
        Return PaletteID2SName(id)
    End Function
    Function PaletteID2SName(index As Integer) As String
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
        Dim tipText As String = "This is a screentip for the item."
        Return tipText
    End Function
    Public Function GetItemSuperTip(ByVal control As IRibbonControl, ByVal index As Integer) As String
        Dim tipText As String = "This is a supertip for the item."
        Return tipText
    End Function
    Public Function GetKeyTip(ByVal control As IRibbonControl) As String
        Select Case control.Id
            Case "Palettes" : Return "GL"
        End Select
    End Function
    Public Function GetScreenTip(ByVal control As IRibbonControl) As String
        Select Case control.Id
            Case "Palettes" : Return "Click to open the palette gallery."
        End Select
    End Function

    Public Sub galleryOnAction(ByVal control As IRibbonControl, ByVal selectedId As String, _
    ByVal selectedIndex As Integer)
        OnAction(control, selectedIndex)
    End Sub
#End Region

End Class
