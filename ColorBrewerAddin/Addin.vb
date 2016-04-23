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
            MsgBox(e.Message)
        End Try
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

    End Sub

    Function GetPaletteData(pal As String, NumColors As Integer) As Array
        Dim filter As String
        filter = "[C] = '" + pal + "' AND [N] = " + NumColors.ToString
        Try
            Return PalettesDataTable.Select(filter)
        Catch e As Exception
            MsgBox(e.Message)
            Return {} 'TODO: Make this an empty array
        End Try
    End Function

    Sub ColorBrewerFill(ByVal chart As Object, ByVal pal As String)
        Dim palette As Array
        Dim series_count As Integer
        Dim rgb_color As Integer
        Dim i As Integer
        With chart
            series_count = .SeriesCollection.Count
            Select Case .ChartType
                'Chart types enumerated here: https://msdn.microsoft.com/en-us/library/office/ff838409.aspx
                Case XlChartType.xlXYScatter, XlChartType.xlXYScatterLines, XlChartType.xlXYScatterLinesNoMarkers, XlChartType.xlXYScatterSmooth, XlChartType.xlRadarMarkers
                    'Points, Lines optional Case
                    palette = GetPaletteData(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
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
                    palette = GetPaletteData(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
                        With .SeriesCollection(i)
                            .Format.Line.ForeColor.RGB = rgb_color
                        End With
                    Next
                Case XlChartType.xlColumnClustered, XlChartType.xlConeCol, XlChartType.xl3DArea, XlChartType.xlAreaStacked, XlChartType.xlAreaStacked100, XlChartType.xlBubble3DEffect, XlChartType.xlPyramidBarClustered, XlChartType.xlRadarFilled
                    'Area Case
                    palette = GetPaletteData(pal, series_count)
                    For i = 1 To series_count
                        rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
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
                            palette = GetPaletteData(pal, .Points.Count)
                            For j = 1 To .Points.Count
                                rgb_color = RGB(palette(j - 1)(2), palette(j - 1)(3), palette(j - 1)(4))
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
                        palette = GetPaletteData(pal, .LegendEntries.Count)
                        For i = 1 To .LegendEntries.Count
                            rgb_color = RGB(palette(i - 1)(2), palette(i - 1)(3), palette(i - 1)(4))
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
    Public Function LoadImage(ByVal imageName As String) As Bitmap
        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim stream As Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + imageName)
        Return New Bitmap(stream)
    End Function
    Public Function GetLabel(ByVal control As IRibbonControl) As String
        Dim strText As String
        Select Case control.Id
            Case "gallery1" : strText = "Select a Device:"
            Case "button1" : strText = "Button in Gallery"
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
        Select Case (itemIndex)
            Case 0 : imageName = "camera.bmp"
            Case 1 : imageName = "video.bmp"
            Case 2 : imageName = "mp3device.bmp"
        End Select

        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim stream As Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + imageName)
        Return New Bitmap(stream)
    End Function
    Public Function GetSize(ByVal control As IRibbonControl) As RibbonControlSize
        Select Case control.Id
            Case "gallery1" : Return RibbonControlSize.RibbonControlSizeLarge
            Case "button1" : Return RibbonControlSize.RibbonControlSizeRegular
        End Select
    End Function
    Private itemCount As Integer = 4 ' Used with GetItemCount.
    Private itemHeight As Integer = 35 ' Used with GetItemHeight.
    Private itemWidth As Integer = 35 ' Used with GetItemWidth.
    Public Function GetEnabled(ByVal control As IRibbonControl) As Boolean
        Return True
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
    Public Function getItemLabel(ByVal control As IRibbonControl, ByVal index As Integer) As String
        Select Case index
            Case 0 : Return "Camera"
            Case 1 : Return "Video Player"
            Case 2 : Return "MP3 Player"
            Case 3 : Return "Cell Phone"
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
            Case "gallery1" : Return "GL"
            Case "button1" : Return "A1"
        End Select
    End Function
    Public Function GetScreenTip(ByVal control As IRibbonControl) As String
        Select Case control.Id
            Case "gallery1" : Return "Click to open a gallery of choices."
            Case "button1" : Return "This is a screentip for the button."
        End Select
    End Function

    Public Sub galleryOnAction(ByVal control As IRibbonControl, ByVal selectedId As String, _
    ByVal selectedIndex As Integer)
        Select Case selectedIndex
            Case 0
                applicationObject.Range("A1").Value = "You clicked a camera."
            Case 1
                applicationObject.Range("A1").Value = "You clicked a video player."
            Case 2
                applicationObject.Range("A1").Value = "You clicked an mp3 device."
            Case 3
                applicationObject.Range("A1").Value = "You clicked a cell phone."
        End Select
    End Sub
    Public Sub buttonOnAction(ByVal control As IRibbonControl)
        MsgBox("Hello world.")
    End Sub
#End Region

End Class
