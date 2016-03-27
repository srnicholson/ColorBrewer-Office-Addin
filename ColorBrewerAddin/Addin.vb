Imports System.Runtime.InteropServices
Imports System.IO
Imports Extensibility
Imports Microsoft.Office.Core
Imports System.Reflection


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
    End Sub
#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return ReadString("RibbonUI.xml")
    End Function

    Public Sub OnAction(ByVal control As IRibbonControl)

        Try

            Select Case control.Id
                Case "customButton1"
                    Select Case applicationObject.Name.ToString
                        Case "Microsoft Excel"
                            Call Excel_Sub()
                        Case "Microsoft Word"
                            Call Word_Sub()
                        Case "Microsoft PowerPoint"
                            Call PowerPoint_Sub()
                        Case Else
                            MsgBox("Error: This Office application is not supported.")
                    End Select

                Case "customButton2"
                    MsgBox("This is the second sample button.")
                Case Else
                    MsgBox("Unkown Control Id: " + control.Id, , "ColorBrewer Office Addin")
            End Select

        Catch throwedException As Exception

            MsgBox("Error: Unexpected state in ColorBrewer OnAction" + vbNewLine + "Error details: " + throwedException.Message)

        End Try

    End Sub

#End Region

#Region "ColorBrewer Methods"

    Public Sub Excel_Sub()
        Dim chart As Object
        Dim chart_type As String
        Dim series_count As Integer
        Dim ColorName As String
        Try
            ColorName = "Accent"
            chart = applicationObject.ActiveChart
            chart_type = chart.ChartType
            series_count = chart.SeriesCollection.Count
            Call ColorBrewerFill(chart, ColorName)
        Catch
            MsgBox("No Chart Selected")
        End Try

        'Utils.Dialog.ShowMessageBox("Chart type is: " + chart_type, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
        'Utils.Dialog.ShowMessageBox("Series count is: " + series_count.ToString, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
    End Sub

    Public Sub Word_Sub()
        Dim chart As Object
        Dim chart_type As String
        Dim series_count As Integer
        Try
            chart = applicationObject.ActiveWindow.Selection.InlineShapes(1).Chart
            chart_type = chart.ChartType
            series_count = chart.SeriesCollection.Count
        Catch
            MsgBox("No Chart Selected")
        End Try

        'Utils.Dialog.ShowMessageBox("Chart type is: " + chart_type, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
        'Utils.Dialog.ShowMessageBox("Series count is: " + series_count.ToString, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
    End Sub
    Public Sub PowerPoint_Sub()
        Dim chart As Object
        Dim chart_type As String
        Dim series_count As Integer
        Try
            chart = applicationObject.ActiveWindow.Selection.ShapeRange(1).Chart
            chart_type = chart.ChartType
            series_count = chart.SeriesCollection.Count
        Catch ex As Exception
            MsgBox("No Chart Selected")
        End Try

        'Utils.Dialog.ShowMessageBox("Chart type is: " + chart_type, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
        'Utils.Dialog.ShowMessageBox("Series count is: " + series_count.ToString, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
    End Sub

    Public Function GetPalette(ByVal pal As String, ByVal NumColors As Integer) As Object
        'Dim PaletteArray(7)() As Double
        'PaletteArray(0)(0) = {{127, 201, 127}, {190, 174, 212}, {253, 192, 134}}
        'PaletteArray(0)(1) = {{127, 201, 127}, {190, 174, 212}, {253, 192, 134}, {255, 255, 153}}
        'GetPalette = PaletteArray(0)(NumColors - 3)
        Dim PaletteArray As Object = {{127, 201, 127}, {190, 174, 212}, {253, 192, 134}}
        GetPalette = PaletteArray
    End Function

    'Function GetTable() As DataTable
    '    ' Create new DataTable instance.
    '    Dim table As New DataTable

    '    ' Create DataTable
    '    table.Columns.Add("ColorName", GetType(String))
    '    table.Columns.Add("NumOfColors", GetType(Integer))
    '    'table.Columns.Add("ColorNum", GetType(Integer))
    '    table.Columns.Add("R", GetType(Integer))
    '    table.Columns.Add("G", GetType(Integer))
    '    table.Columns.Add("B", GetType(Integer))

    '    ' Add five rows with those columns filled in the DataTable.
    '    table.Rows.Add(25, "Indocin", "David", DateTime.Now)
    '    table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now)
    '    table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now)
    '    table.Rows.Add(21, "Combivent", "Janet", DateTime.Now)
    '    table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now)
    '    Return table
    'End Function

    Sub ColorBrewerFill(ByVal chart As Object, ByVal pal As String)
        Dim palette(4, 2) As Integer 'Max dimensions should be specified here something like (9, 2) or (2, 9)
        Dim series_count As Integer
        Dim rgb_color As Integer
        Dim i As Integer
        With chart
            series_count = .SeriesCollection.Count
            Select Case .ChartType
                'Chart types enumerated here: https://msdn.microsoft.com/en-us/library/office/ff838409.aspx
                Case XlChartType.xlXYScatter, XlChartType.xlXYScatterLines, XlChartType.xlXYScatterLinesNoMarkers, XlChartType.xlXYScatterSmooth, XlChartType.xlRadarMarkers
                    'Points, Lines optional Case
                    palette = GetPalette(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1, 0), palette(i - 1, 1), palette(i - 1, 2))
                        With .SeriesCollection(i)
                            .MarkerForegroundColor = rgb_color
                            .MarkerBackgroundColor = rgb_color
                            If .Format.Line.Visible = True Then
                                .Format.Line.ForeColor.RGB = rgb_color
                            End If
                        End With
                    Next
                Case XlChartType.xlLine, XlChartType.xlRadar
                    'Line Only Case
                    palette = GetPalette(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1, 0), palette(i - 1, 1), palette(i - 1, 2))
                        With .SeriesCollection(i)
                            .Format.Line.ForeColor.RGB = rgb_color
                        End With
                    Next
                Case XlChartType.xlColumnClustered, XlChartType.xlConeCol, XlChartType.xl3DArea, XlChartType.xlAreaStacked, XlChartType.xlAreaStacked100, XlChartType.xlBubble3DEffect, XlChartType.xlPyramidBarClustered, XlChartType.xlRadarFilled
                    'Area Case
                    palette = GetPalette(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1, 0), palette(i - 1, 1), palette(i - 1, 2))
                        With .SeriesCollection(i)
                            .Interior.Color = rgb_color
                            .Border.Color = rgb_color
                        End With
                    Next
                Case XlChartType.xlDoughnut, XlChartType.xlDoughnutExploded, XlChartType.xlPie
                    Dim j As Integer
                    'Pie Case
                    For i = 1 To series_count
                        With .SeriesCollection(i)
                            palette = GetPalette(pal, .Points.Count)
                            For j = 1 To .Points.Count
                                rgb_color = RGB(palette(j - 1, 0), palette(j - 1, 1), palette(j - 1, 2))
                                With .Points(j)
                                    .Interior.Color = rgb_color
                                    .Border.Color = rgb_color
                                End With
                            Next
                        End With
                    Next
                Case XlChartType.xlSurface
                    'Surface Case
                    With .Legend
                        'TODO: If Legend doesn't exist, display it temporarily to change colors
                        palette = GetPalette(pal, .LegendEntries.Count)
                        For i = 1 To .LegendEntries.Count
                            rgb_color = RGB(palette(i - 1, 0), palette(i - 1, 1), palette(i - 1, 2))
                            .LegendEntries(i).LegendKey.Interior.Color = rgb_color
                        Next
                    End With
                Case XlChartType.xlSurfaceWireframe, XlChartType.xlSurfaceTopViewWireframe
                    'Surface Wireframe Case
                    With .Legend
                        palette = GetPalette(pal, .LegendEntries.Count)
                        For i = 1 To .LegendEntries.Count
                            rgb_color = RGB(palette(i - 1, 0), palette(i - 1, 1), palette(i - 1, 2))
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
#End Region

#Region "XML Methods"
    ' Modified from https://github.com/NetOfficeFw/NetOffice/blob/master/Examples/Misc/VB/COMAddin%20Examples/SuperAddin/Addin.vb#L192
    Private Shared Function ReadString(ByVal fileName As String) As String

        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim resourceStream As System.IO.Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + fileName)
        If (IsNothing(resourceStream)) Then
            Throw (New System.IO.IOException("Error accessing resource Stream."))
        End If

        Dim textStreamReader As System.IO.StreamReader = New System.IO.StreamReader(resourceStream)
        If (IsNothing(textStreamReader)) Then
            Throw (New System.IO.IOException("Error accessing resource File."))
        End If

        Dim text As String = textStreamReader.ReadToEnd()
        resourceStream.Close()
        textStreamReader.Close()
        Return text

    End Function

#End Region

End Class
