Imports System.Math
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Resources

Public Class CChartControl

    'Private m_intYScaleNum As Integer = 10            'Y Line scale 
    Private m_dblMajorUnit As Double = 1
    Private Const m_cstDotRadias As Single = 4     'Size for dot
    Private m_intLimitXData As Integer            'X data point limit
    Private m_intHistrogramStep As Integer = 100

    Public Enum enumGraphType
        eLineSeriesGraph = 0
        eColumnGraph
        eXYChart
        eHistrogram
        eBoxPlot
        ePieChart
    End Enum

    Public Enum enumLineLinkType
        eLineHorizontal = 0
        eLineVertical
        eNoLine
    End Enum

    Private Enum enumColumnType
        eLabel = 0
        eRealValue
        eSection
    End Enum

    Private Structure SXDetail
        Dim strXString As String
        Dim pntStart As Point
        Dim pntEnd As Point
    End Structure

    Private Structure SValueAndPoint
        Dim strColumnName As String      'Keep column name
        Dim dtbXYValue As DataTable     'Keep Raw data in chart
        Dim pntPlot() As Point                     'Keep the point
        Dim penColor As Pen                      'Keep pen: Color, Draw properties.
        Dim bWriteDot As Boolean             'Write dot or not, True=Write
        Dim rectLabel As Rectangle           'Rect for Label
        Dim eMarkType As enumMarkerType       'Marker type
    End Structure

    Public Structure SMinMax
        Dim dblMin As Double
        Dim dblMax As Double
    End Structure

    Public Structure SFontChart
        Dim hFontScaleX As Font
        Dim hFontScaleY As Font

        Dim hFontXLabel As Font
        Dim hFontYLabel As Font
        Dim hFontChartLabel As Font

        Dim hFontLinearEquation As Font
        Dim hFontChartDetail As Font
    End Structure

    Public m_frmChart As Form
    Public m_sFontChart As SFontChart
    Private WithEvents m_picChartAxis As PictureBox
    Private m_hscScrollGraph As HScrollBar
    Private WithEvents m_optPen As CheckBox
    Private m_dtbDataShow As DataTable
    Private m_dtbDataChart As DataTable

    Private m_eGraphType As enumGraphType
    Private m_eLineLinkType As enumLineLinkType

    Private m_SDataPoint() As SValueAndPoint
    Private m_dblYMin As Double
    Private m_dblYMax As Double
    Private m_dblXMin As Double
    Private m_dblXMax As Double
    Private m_rectYMin As Rectangle
    Private m_rectYMax As Rectangle
    Private m_rectYMajorUnit As Rectangle
    Private m_rectXScale As Rectangle
    Private m_rectChartArea As Rectangle
    Private m_sLinearParam As CDataAnalyzer.SLinearParameter
    Private m_bDrawXScale As Boolean
    Private m_bDrawYScale As Boolean
    Private m_intXScaleTextHeight As Integer
    Private m_intXScaleDegree As Integer
    Private m_intFormWidth As Integer
    Private m_intFormHeight As Integer
    Private m_intPicWidth As Integer
    Private m_intPicHeight As Integer
    Private m_imgChart As Image
    Private m_pntLastMouseMove As Point

    Private Enum enumMarkerType
        eNoMarker = -1
        eCircle = 0
        eTriangleDown
        eTriangleUp
        eStar
        eRectangle
        eRhombus
        eCrossX
        eCrossPlus
        eLinearHorizontal
    End Enum

    Private Sub DrawMarker(ByVal gr As Graphics, ByVal penMark As Pen, ByVal pntMark As Point, ByVal sngMarkWidth As Single, ByVal eMarkType As enumMarkerType)

        Dim rectfDraw As RectangleF
        rectfDraw.X = pntMark.X - (sngMarkWidth / 2)
        rectfDraw.Y = pntMark.Y - (sngMarkWidth / 2)
        rectfDraw.Width = sngMarkWidth
        rectfDraw.Height = sngMarkWidth
        Select Case eMarkType
            Case enumMarkerType.eCircle
                gr.FillEllipse(penMark.Brush, rectfDraw)
            Case enumMarkerType.eStar
                gr.DrawLine(penMark, pntMark.X, pntMark.Y - sngMarkWidth, pntMark.X, pntMark.Y + sngMarkWidth)
                gr.DrawLine(penMark, pntMark.X - sngMarkWidth, pntMark.Y, pntMark.X + sngMarkWidth, pntMark.Y)
                gr.DrawLine(penMark, pntMark.X - (sngMarkWidth / 2), pntMark.Y - (sngMarkWidth / 2), pntMark.X + (sngMarkWidth / 2), pntMark.Y + (sngMarkWidth / 2))
                gr.DrawLine(penMark, pntMark.X - (sngMarkWidth / 2), pntMark.Y + (sngMarkWidth / 2), pntMark.X + (sngMarkWidth / 2), pntMark.Y - (sngMarkWidth / 2))
            Case enumMarkerType.eTriangleUp
                Dim pntfTri(2) As PointF
                pntfTri(0).X = pntMark.X
                pntfTri(0).Y = pntMark.Y - sngMarkWidth
                pntfTri(1).X = pntMark.X - sngMarkWidth
                pntfTri(1).Y = pntMark.Y + sngMarkWidth
                pntfTri(2).X = pntMark.X + sngMarkWidth
                pntfTri(2).Y = pntMark.Y + sngMarkWidth
                gr.FillPolygon(penMark.Brush, pntfTri)
            Case enumMarkerType.eRectangle
                gr.FillRectangle(penMark.Brush, rectfDraw)
            Case enumMarkerType.eRhombus
                Dim pntfRhombus(3) As PointF
                pntfRhombus(0).X = pntMark.X - sngMarkWidth
                pntfRhombus(0).Y = pntMark.Y
                pntfRhombus(1).X = pntMark.X
                pntfRhombus(1).Y = pntMark.Y - sngMarkWidth
                pntfRhombus(2).X = pntMark.X + sngMarkWidth
                pntfRhombus(2).Y = pntMark.Y
                pntfRhombus(3).X = pntMark.X
                pntfRhombus(3).Y = pntMark.Y + sngMarkWidth
                gr.FillPolygon(penMark.Brush, pntfRhombus)
            Case enumMarkerType.eCrossX
                gr.DrawLine(penMark, pntMark.X - (sngMarkWidth / 2), pntMark.Y - (sngMarkWidth / 2), pntMark.X + (sngMarkWidth / 2), pntMark.Y + (sngMarkWidth / 2))
                gr.DrawLine(penMark, pntMark.X - (sngMarkWidth / 2), pntMark.Y + (sngMarkWidth / 2), pntMark.X + (sngMarkWidth / 2), pntMark.Y - (sngMarkWidth / 2))
            Case enumMarkerType.eCrossPlus
                gr.DrawLine(penMark, pntMark.X, pntMark.Y - sngMarkWidth, pntMark.X, pntMark.Y + sngMarkWidth)
                gr.DrawLine(penMark, pntMark.X - sngMarkWidth, pntMark.Y, pntMark.X + sngMarkWidth, pntMark.Y)
            Case enumMarkerType.eTriangleDown
                Dim pntfTri(2) As PointF
                pntfTri(0).X = pntMark.X - sngMarkWidth
                pntfTri(0).Y = pntMark.Y - sngMarkWidth
                pntfTri(1).X = pntMark.X + sngMarkWidth
                pntfTri(1).Y = pntMark.Y - sngMarkWidth
                pntfTri(2).X = pntMark.X
                pntfTri(2).Y = pntMark.Y + sngMarkWidth
                gr.FillPolygon(penMark.Brush, pntfTri)
            Case enumMarkerType.eLinearHorizontal
                gr.DrawLine(penMark, pntMark.X - 1, pntMark.Y, pntMark.X + 1, pntMark.Y)
            Case Else
                gr.FillEllipse(penMark.Brush, pntMark.X, pntMark.Y, sngMarkWidth, sngMarkWidth)
        End Select
    End Sub

    Private Function InitPen(ByVal nPen As Integer) As Pen()
        Dim penArray(nPen) As Pen
        For nArray As Integer = 0 To penArray.Length - 1
            If nArray = 0 Then penArray(0) = New Pen(Color.Blue)
            If nArray = 1 Then penArray(1) = New Pen(Color.Magenta)
            If nArray = 2 Then penArray(2) = New Pen(Color.Green)
            If nArray = 3 Then penArray(3) = New Pen(Color.Red)
            If nArray = 4 Then penArray(4) = New Pen(Color.Violet)
            If nArray = 5 Then penArray(5) = New Pen(Color.Aquamarine)
            If nArray = 6 Then penArray(6) = New Pen(Color.LightSeaGreen)
            If nArray = 7 Then penArray(7) = New Pen(Color.Orange)
            If nArray = 8 Then penArray(8) = New Pen(Color.Aqua)
            If nArray = 9 Then penArray(9) = New Pen(Color.Lime)
            If nArray = 10 Then penArray(10) = New Pen(Color.LightCoral)
            If nArray = 11 Then penArray(11) = New Pen(Color.Khaki)
            If nArray = 12 Then penArray(12) = New Pen(Color.Honeydew)
            If nArray = 13 Then penArray(13) = New Pen(Color.DarkMagenta)
            If nArray = 14 Then penArray(14) = New Pen(Color.PapayaWhip)
            If nArray = 15 Then penArray(15) = New Pen(Color.SteelBlue)
            If nArray = 16 Then penArray(16) = New Pen(Color.Thistle)
            If nArray = 17 Then penArray(17) = New Pen(Color.Honeydew)
            If nArray = 18 Then penArray(18) = New Pen(Color.Olive)
            If nArray = 19 Then penArray(19) = New Pen(Color.Peru)
            If nArray = 20 Then penArray(20) = New Pen(Color.PowderBlue)
            If nArray > 20 Then penArray(nArray) = New Pen(Color.DarkBlue)
        Next
        InitPen = penArray
    End Function

    Private Sub InitPenAndBrush(ByVal dtbDataChart As DataTable, ByVal bDrawMarker As Boolean)
        Dim penArray() As Pen = InitPen(dtbDataChart.Columns.Count)
        Dim nY As Integer = 0
        For nCol As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
            Dim strColName As String = dtbDataChart.Columns(nCol).ColumnName
            Dim bWriteDot As Boolean = True

            If InStr(strColName, "1Sigma") Then
                penArray(nY).Color = Color.Green
                penArray(nY).DashStyle = Drawing2D.DashStyle.DashDotDot
                penArray(nY).Width = 2
                bWriteDot = False
            ElseIf InStr(strColName, "2Sigma") Then
                penArray(nY).Color = Color.YellowGreen
                penArray(nY).DashStyle = Drawing2D.DashStyle.DashDotDot
                penArray(nY).Width = 2
                bWriteDot = False
            ElseIf InStr(strColName, "UCL") Or InStr(strColName, "LCL") Then
                penArray(nY).Color = Color.Red
                penArray(nY).DashStyle = Drawing2D.DashStyle.DashDotDot
                penArray(nY).Width = 2
                bWriteDot = False
            ElseIf InStr(strColName, "Mean") Or InStr(strColName, "Reference") Then
                penArray(nY).Color = Color.Gray
                penArray(nY).DashStyle = Drawing2D.DashStyle.DashDotDot
                penArray(nY).Width = 2
                bWriteDot = False
            ElseIf InStr(strColName, ">") Or InStr(strColName, "<") Or InStr(strColName, "=") Then
                penArray(nY).DashStyle = Drawing2D.DashStyle.DashDotDot
                penArray(nY).Width = 2
                bWriteDot = False
            End If
            m_SDataPoint(nY).strColumnName = strColName
            m_SDataPoint(nY).penColor = penArray(nY)
            m_SDataPoint(nY).bWriteDot = bWriteDot
            If bDrawMarker Then
                If bWriteDot = False Then
                    m_SDataPoint(nY).eMarkType = enumMarkerType.eLinearHorizontal
                Else
                    m_SDataPoint(nY).eMarkType = nY
                End If

            Else
                m_SDataPoint(nY).eMarkType = enumMarkerType.eNoMarker
            End If
            nY = nY + 1
        Next nCol

    End Sub

    Private Sub InitChartFont()
        m_sFontChart.hFontChartLabel = New Font("MS Sans Serif", 12, FontStyle.Bold)
        m_sFontChart.hFontScaleX = New Font("MS Sans Serif", 7.25, FontStyle.Regular)
        m_sFontChart.hFontScaleY = New Font("MS Sans Serif", 7.25, FontStyle.Italic)
        m_sFontChart.hFontXLabel = New Font("MS Sans Serif", 10, FontStyle.Bold)
        m_sFontChart.hFontYLabel = New Font("MS Sans Serif", 8, FontStyle.Bold)
        m_sFontChart.hFontLinearEquation = New Font("MS Sans Serif", 8, FontStyle.Italic)
        m_sFontChart.hFontChartDetail = New Font("MS Sans Serif", 7, FontStyle.Italic)
    End Sub

    Public Sub New(ByVal intWidth As Integer, ByVal intHeight As Integer, ByVal dtbDataChart As DataTable, ByVal eGraphType As enumGraphType, Optional ByVal bDrawXScale As Boolean = False, Optional ByVal bDrawYScale As Boolean = True, Optional ByVal eLineLink As enumLineLinkType = enumLineLinkType.eLineHorizontal, Optional ByVal bDrawMarker As Boolean = True)
        'Need to convert format 2 | X | RealNumber | Additional Info | Y1|Y2|...Yn|

        m_bDrawXScale = bDrawXScale
        m_bDrawYScale = bDrawYScale
        m_intXScaleDegree = -60

        m_intFormWidth = intWidth + 40
        m_intFormHeight = intHeight + 90
        m_intPicWidth = intWidth
        m_intPicHeight = intHeight

        m_eGraphType = eGraphType
        InitChartFont()

        m_dtbDataChart = dtbDataChart.Copy
        m_dtbDataShow = m_dtbDataChart.Copy
        m_dtbDataShow = PrepareDataByChartType(m_dtbDataShow, eGraphType)

        dtbDataChart.TableName = dtbDataChart.TableName
        m_intLimitXData = m_dtbDataShow.Rows.Count

        'CreateForm(strChartText, intWidth, intHeight)
        m_eLineLinkType = eLineLink

        ReDim m_SDataPoint(m_dtbDataChart.Columns.Count - FindFirstYColumn(m_dtbDataChart) - 1)

        InitPenAndBrush(dtbDataChart, bDrawMarker)

        'Find Y min,max
        Dim sYParam As SMinMax = FindYMinMax(m_dtbDataShow, eGraphType)
        m_dblYMin = sYParam.dblMin
        m_dblYMax = sYParam.dblMax
        m_dblMajorUnit = (m_dblYMax - m_dblYMin) / 10

        m_intXScaleTextHeight = 0
    End Sub

    Private Sub OptPen_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles m_optPen.MouseDown, m_optPen.Click
        If e.Button = MouseButtons.Left Then
            If m_optPen.Checked = True Then
                Dim ms As New System.IO.MemoryStream(My.Resources.LibResource.PenCur)
                m_picChartAxis.Cursor = New Cursor(ms)
            Else
                m_picChartAxis.Cursor = Cursors.Default
            End If
        Else
            Dim ColorDlg As New ColorDialog
            ColorDlg.Color = m_optPen.BackColor
            ColorDlg.FullOpen = True
            If ColorDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                m_optPen.BackColor = ColorDlg.Color
                m_optPen.Refresh()
            End If
        End If
    End Sub

    Private Function PrepareDataByChartType(ByVal dtbDataChart As DataTable, ByVal eChartType As enumGraphType) As DataTable
        Select Case eChartType
            Case enumGraphType.ePieChart
                Dim strColCount As String = dtbDataChart.Columns(dtbDataChart.Columns.Count - 1).ColumnName
                Dim dtbDetail As DataTable = dtbDataChart.DefaultView.ToTable(True, strColCount)
                Dim dtbCount As New DataTable(dtbDataChart.TableName)
                dtbCount.Columns.Add(strColCount)
                dtbCount.Columns.Add("CountItem", GetType(Int32))
                For nRow As Integer = 0 To dtbDetail.Rows.Count - 1
                    Dim strDetail As String = dtbDetail.Rows(nRow).Item(strColCount)
                    Dim drSelect() As DataRow = dtbDataChart.Select("[" & strColCount & "]='" & strDetail & "'")
                    dtbCount.Rows.Add(strDetail, drSelect.Length)
                Next nRow
                dtbCount.DefaultView.Sort = "CountItem"
                dtbCount = dtbCount.DefaultView.ToTable
                Return dtbCount
            Case enumGraphType.eXYChart
                If dtbDataChart.Columns(0).DataType IsNot GetType(DateTime) Then
                    Dim clsAnalyze As New CDataAnalyzer
                    m_sLinearParam = clsAnalyze.CalculateLinearRegression(dtbDataChart.Columns(0).ColumnName, dtbDataChart.Columns(FindFirstYColumn(dtbDataChart)).ColumnName, dtbDataChart)
                    dtbDataChart.Columns.Add("RealXValue", GetType(Double))
                    For nData As Integer = 0 To dtbDataChart.Rows.Count - 1
                        dtbDataChart.Rows(nData).Item("RealXValue") = dtbDataChart.Rows(nData).Item(0)
                    Next nData
                Else
                    dtbDataChart.Columns.Add("RealXValue", GetType(Double))
                    For nData As Integer = 0 To dtbDataChart.Rows.Count - 1
                        dtbDataChart.Rows(nData).Item("RealXValue") = CDate(dtbDataChart.Rows(nData).Item(0)).ToOADate
                    Next nData
                End If

                dtbDataChart.Columns("RealXValue").SetOrdinal(enumColumnType.eRealValue)
                If dtbDataChart.Columns(enumColumnType.eSection).DataType IsNot GetType(String) Then
                    dtbDataChart.Columns.Add("Section")
                    dtbDataChart.Columns("Section").SetOrdinal(enumColumnType.eSection)
                End If
                PrepareDataByChartType = dtbDataChart.Copy
            Case enumGraphType.eBoxPlot
                dtbDataChart = AddValue2Table(dtbDataChart)
                PrepareDataByChartType = dtbDataChart.Copy
            Case enumGraphType.eHistrogram
                Dim dtbHistrogram As DataTable = AnalyzeHistrogram(dtbDataChart, m_intHistrogramStep)
                dtbHistrogram.Columns.Add("RealXValue", GetType(Double))
                dtbHistrogram.Columns("RealXValue").Expression = "[" & dtbHistrogram.Columns(0).ColumnName & "]"
                dtbHistrogram.Columns("RealXValue").SetOrdinal(enumColumnType.eRealValue)
                If dtbHistrogram.Columns(enumColumnType.eSection).DataType IsNot GetType(String) Then
                    dtbHistrogram.Columns.Add("Section")
                    dtbHistrogram.Columns("Section").SetOrdinal(enumColumnType.eSection)
                End If
                PrepareDataByChartType = dtbHistrogram.Copy
            Case Else
                PrepareDataByChartType = AddOrder2Table(dtbDataChart)
        End Select
        Dim sXParam As SMinMax = FindXMinMax(PrepareDataByChartType.Columns(enumColumnType.eRealValue).ColumnName, PrepareDataByChartType)
        m_dblXMin = sXParam.dblMin
        m_dblXMax = sXParam.dblMax
    End Function

    Private Function AddOrder2Table(ByVal dtbDataChart As DataTable) As DataTable
        Dim strColName As String = dtbDataChart.Columns(0).ColumnName

        Dim dcAdd As New DataColumn("RealXValue", GetType(Int64))
        dtbDataChart.Columns.Add(dcAdd)
        'Dim dcAdd As DataColumn = dtbDataChart.Columns.Add("RealXValue", GetType(Long))
        For nData As Integer = 0 To dtbDataChart.Rows.Count - 1
            dtbDataChart.Rows(nData).Item("RealXValue") = nData + 1
        Next

        dtbDataChart.Columns("RealXValue").SetOrdinal(1)
        AddOrder2Table = dtbDataChart
    End Function

    Private Function AddValue2Table(ByVal dtbDataChart As DataTable)
        Dim strColName As String = dtbDataChart.Columns(0).ColumnName
        If dtbDataChart.Columns(0).DataType Is GetType(String) Then
            dtbDataChart.DefaultView.Sort = strColName
            Dim dtbDistrinct As DataTable = dtbDataChart.DefaultView.ToTable(True, strColName)
            dtbDataChart.Columns.Add("RealXValue", GetType(Long))
            'Dim drAdd As DataRow = dtbDataChart.Rows.Add()
            'drAdd.Item("Label") = 0
            For nDis As Integer = 0 To dtbDistrinct.Rows.Count - 1
                Dim strLabel As String = dtbDistrinct.Rows(nDis).Item(strColName).ToString
                Dim drLabel() As DataRow = dtbDataChart.Select("[" & strColName & "]='" & strLabel & "'")
                For nLabel As Integer = 0 To drLabel.Length - 1
                    drLabel(nLabel).Item("RealXValue") = nDis + 1
                Next nLabel
            Next nDis
        ElseIf dtbDataChart.Columns(0).DataType Is GetType(DateTime) Then
            dtbDataChart.Columns.Add("RealXValue", GetType(Long))
            'Dim drAdd As DataRow = dtbDataChart.Rows.Add()
            'drAdd.Item("Label") = 0
            For nData As Integer = 0 To dtbDataChart.Rows.Count - 1
                If Not dtbDataChart.Rows(nData).Item(strColName) Is DBNull.Value Then
                    dtbDataChart.Rows(nData).Item("RealXValue") = CDate(dtbDataChart.Rows(nData).Item(strColName)).ToFileTime
                End If
            Next nData
        End If

        dtbDataChart.Columns("RealXValue").SetOrdinal(1)
        AddValue2Table = dtbDataChart
    End Function

    'Private Function AnalyzeHistrogram(ByVal dtbDataChart As DataTable, ByVal nHistrogramStep As Integer) As DataTable
    '    Dim dtbHistrogram As New DataTable(dtbDataChart.TableName)

    '    dtbHistrogram.Columns.Add(dtbDataChart.Columns(0).ColumnName, Type.GetType("System.Double")) 'X
    '    Dim sYMinMax As SMinMax = FindYMinMax(dtbDataChart, m_eGraphType)
    '    Dim dblMax As Double = sYMinMax.dblMax
    '    Dim dblMin As Double = sYMinMax.dblMin

    '    For nYData As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
    '        Dim strCol As String = dtbDataChart.Columns(nYData).ColumnName
    '        dtbHistrogram.Columns.Add(strCol, Type.GetType("System.Int32"))  'Y
    '    Next nYData

    '    If dblMin = dblMax Then
    '        dblMin = dblMin - 1
    '        dblMax = dblMax + 1
    '    End If

    '    Dim dblDiscreteStep As Double = (dblMax - dblMin) / nHistrogramStep

    '    Dim dblStep As Double = dblMin
    '    While dblStep <= dblMax
    '        Dim dblStepMax As Double = dblStep + dblDiscreteStep
    '        dtbHistrogram.Rows.Add()
    '        dtbHistrogram.Rows(dtbHistrogram.Rows.Count - 1).Item(0) = Format(dblStep, "0.0000")
    '        For nYData As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
    '            Dim strCol As String = dtbDataChart.Columns(nYData).ColumnName
    '            Dim strFilter As String = "[" & strCol & "]>=" & dblStep & " AND [" & strCol & "]<" & dblStepMax
    '            Dim dtrFilter() As DataRow = dtbDataChart.Select(strFilter)
    '            dtbHistrogram.Rows(dtbHistrogram.Rows.Count - 1).Item(strCol) = dtrFilter.Length
    '        Next nYData
    '        dblStep = dblStepMax
    '    End While
    '    AnalyzeHistrogram = dtbHistrogram
    'End Function

    Private Function AnalyzeHistrogram(ByVal dtbDataChart As DataTable, ByVal nHistrogramStep As Integer) As DataTable
        Dim dtbHistrogram As New DataTable(dtbDataChart.TableName)
        dtbHistrogram.Columns.Add(dtbDataChart.Columns(0).ColumnName, Type.GetType("System.Double")) 'X
        Dim sYMinMax As SMinMax = FindYMinMax(dtbDataChart, m_eGraphType)
        Dim dblMax As Double = sYMinMax.dblMax
        Dim dblMin As Double = sYMinMax.dblMin

        For nYData As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
            Dim strCol As String = dtbDataChart.Columns(nYData).ColumnName
            dtbHistrogram.Columns.Add(strCol, Type.GetType("System.Int32"))  'Y
        Next nYData

        If dblMin = dblMax Then
            dblMin = dblMin - 1
            dblMax = dblMax + 1
        End If
        Dim dblDiscreteStep As Double = (dblMax - dblMin) / nHistrogramStep
        Dim dblStep As Double = dblMin
        While dblStep <= dblMax         'Create Data Row
            Dim dblStepMax As Double = dblStep + dblDiscreteStep
            Dim drRowAdd As DataRow = dtbHistrogram.Rows.Add(Format(dblStep, "0.0000"))
            dblStep = dblStepMax
        End While

        For nCol As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
            Dim strColName As String = dtbDataChart.Columns(nCol).ColumnName
            dtbDataChart.DefaultView.Sort = strColName & " ASC"
            Dim dtbSortData As DataTable = dtbDataChart.DefaultView.ToTable(False, strColName)
            dtbDataChart.DefaultView.Sort = ""
            dblStep = dblMin
            Dim nStart As Integer = 0
            For nHistogram As Integer = 0 To dtbHistrogram.Rows.Count - 1
                Dim drHistogramRow As DataRow = dtbHistrogram.Rows(nHistogram)
                Dim dblStepMax As Double = dblStep + dblDiscreteStep
                Dim nRowCount As Integer = 0
                For nData As Integer = nStart To dtbSortData.Rows.Count - 1
                    nStart = nData
                    Dim objValue As Object = dtbSortData.Rows(nData).Item(strColName)
                    If objValue IsNot DBNull.Value Then
                        If objValue >= dblStep And objValue < dblStepMax Then
                            nRowCount = nRowCount + 1
                        Else
                            Exit For
                        End If
                    End If
                Next nData
                drHistogramRow.Item(strColName) = nRowCount
                dblStep = dblStepMax
            Next nHistogram
        Next nCol
        AnalyzeHistrogram = dtbHistrogram
    End Function

    Private Sub CreateForm(ByVal strChartText As String, ByVal intWidth As Integer, ByVal intHeight As Integer)
        m_frmChart = New Form
        m_frmChart.Width = intWidth
        m_frmChart.Height = intHeight
        m_frmChart.StartPosition = FormStartPosition.CenterParent

        Dim mnuStrip As New StatusStrip
        mnuStrip.Items.Add("Export Data")
        Dim tsmMenu As New ToolStripDropDownButton
        tsmMenu.DropDownItems.Add("Export picture")
        tsmMenu.DropDownItems.Add("Export Data")
        mnuStrip.Items.Add(tsmMenu)
        m_frmChart.Controls.Add(mnuStrip)

        m_picChartAxis = New PictureBox
        m_picChartAxis.Left = 10
        m_frmChart.Top = 20
        m_frmChart.Controls.Add(m_picChartAxis)
        m_frmChart.Text = strChartText

        m_hscScrollGraph = New HScrollBar
        m_hscScrollGraph.Visible = True
        m_hscScrollGraph.Minimum = 0
        Dim intMax As Integer = m_dtbDataShow.Rows.Count - m_intLimitXData
        If intMax < 0 Then intMax = 0
        m_hscScrollGraph.Maximum = intMax
        m_hscScrollGraph.SmallChange = 1
        m_frmChart.Controls.Add(m_hscScrollGraph)

        m_optPen = New CheckBox
        m_optPen.Location = m_hscScrollGraph.Location
        m_optPen.Left = 10
        m_optPen.Name = "cmdPen"
        m_optPen.Size = New System.Drawing.Size(25, 25)
        'cmdPen.TabIndex = 3
        m_optPen.Image = My.Resources.LibResource.Pen.ToBitmap
        m_optPen.BackColor = Color.Red
        'm_optPen.UseVisualStyleBackColor = True
        m_optPen.Appearance = Appearance.Button
        m_frmChart.Controls.Add(m_optPen)

        SetControlSize()

        AddHandler m_frmChart.Resize, AddressOf Form_Resize
        AddHandler tsmMenu.DropDownItemClicked, AddressOf tsmMenu_Click

        'AddHandler m_frmChart.ResizeEnd, AddressOf Form_ResizeEnd
        'AddHandler m_frmChart.Validated, AddressOf Form_ResizeEnd
        If m_eGraphType <> enumGraphType.ePieChart Then
            'AddHandler m_picChartAxis.MouseDown, AddressOf picChartAxis_MouseDown
            'AddHandler m_picChartAxis.DoubleClick, AddressOf picChartAxis_DoubleClick
            'AddHandler m_picChartAxis.MouseMove, AddressOf picChartAxis_MouseMove
            AddHandler m_hscScrollGraph.ValueChanged, AddressOf hscScrollGraph_ValueChanged
        End If
        m_picChartAxis.Capture = True
    End Sub

    Private Sub tsmMenu_Click(ByVal sender As System.Object, ByVal e As ToolStripItemClickedEventArgs)
        Dim tsmnuItem As ToolStripMenuItem = e.ClickedItem
        Dim tsDropdown As ToolStripDropDownButton = sender
        tsDropdown.HideDropDown()
        Dim strSelectMenu As String = tsmnuItem.Text
        Dim dlgExport As New SaveFileDialog
        Select Case strSelectMenu.ToUpper
            Case "EXPORT PICTURE"
                dlgExport.FileName = m_dtbDataChart.TableName.Replace("'", "").Replace("*", "") & ".png"
                dlgExport.Filter = "png Files (*.png)|*.png"
                If dlgExport.ShowDialog = Windows.Forms.DialogResult.OK Then 'Equal Save (1)
                    Dim bmChart As Bitmap = GetBitmap()
                    bmChart.Save(dlgExport.FileName, Imaging.ImageFormat.Png)
                    m_picChartAxis.Image = bmChart
                End If
            Case "EXPORT DATA"
                dlgExport.FileName = m_dtbDataChart.TableName & ".csv"
                dlgExport.Filter = "csv Files (*.csv)|*.csv"
                If dlgExport.ShowDialog = Windows.Forms.DialogResult.OK Then 'Equal Save (1)
                    Dim clsDataExport As New CDataExport
                    clsDataExport.ExportDatatableToCSV(dlgExport.FileName, m_dtbDataChart, False)
                End If
        End Select
    End Sub

    Private Sub hscScrollGraph_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim intScrollValue As Integer = m_hscScrollGraph.Value
        'If intScrollValue > 0 Then
        If m_intLimitXData > m_dtbDataChart.Rows.Count Then m_intLimitXData = m_dtbDataChart.Rows.Count
        m_dtbDataShow.Columns.Clear()
        m_dtbDataShow = m_dtbDataChart.Clone
        For nData As Integer = 0 To m_intLimitXData - 1
            If nData + intScrollValue > m_dtbDataChart.Rows.Count - 1 Then Exit For
            m_dtbDataShow.Rows.Add(m_dtbDataChart.Rows(nData + intScrollValue).ItemArray)
        Next nData
        If m_eGraphType = enumGraphType.eXYChart Then
            Dim sXPara As SMinMax = FindXMinMax(m_dtbDataShow.Columns(0).ColumnName, m_dtbDataShow)
            m_dblXMax = sXPara.dblMax
            m_dblXMin = sXPara.dblMin
        End If
        m_dtbDataShow = PrepareDataByChartType(m_dtbDataShow, m_eGraphType)
        m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
        'End If
    End Sub

    Private Function FindXMinMax(ByVal strXCol As String, ByVal dtbDataChart As DataTable) As SMinMax
        ' If eGraphType = enumGraphType.eXYChart Then    'Find X min,max
        Dim dtbDataSampling As DataTable
        Dim nStep As Integer = dtbDataChart.Rows.Count \ 1000
        If dtbDataChart.Rows.Count > 1000 Then   'use sampling to increase speed to find min,max 
            dtbDataSampling = dtbDataChart.Clone
            For nSampling As Integer = 0 To dtbDataChart.Rows.Count - 1 Step nStep
                dtbDataSampling.Rows.Add(dtbDataChart.Rows(nSampling).ItemArray)
            Next nSampling
        Else
            dtbDataSampling = dtbDataChart.Copy
        End If

        Dim objMax As Object = dtbDataSampling.Compute("MAX([" & strXCol & "])", "")
        Dim objMin As Object = dtbDataSampling.Compute("MIN([" & strXCol & "])", "")

        If Not objMin Is DBNull.Value Then FindXMinMax.dblMin = objMin
        If Not objMax Is DBNull.Value Then FindXMinMax.dblMax = objMax

        If FindXMinMax.dblMin = FindXMinMax.dblMax Then
            FindXMinMax.dblMin = FindXMinMax.dblMin - 1
            FindXMinMax.dblMax = FindXMinMax.dblMax + 1
        End If
    End Function

    Private Function FindFirstYColumn(ByVal dtbData As DataTable) As Integer
        Dim nFirstCol As Integer = 0
        For nCol As Integer = 1 To dtbData.Columns.Count - 1
            Dim colDataType As Type = dtbData.Columns(nCol).DataType
            If Not colDataType Is Type.GetType("System.String") And Not colDataType Is Type.GetType("System.DateTime") _
            And dtbData.Columns(nCol).ColumnName <> "RealXValue" Then
                nFirstCol = nCol
                Exit For
            End If
        Next nCol
        If nFirstCol = 0 Then nFirstCol = 1
        FindFirstYColumn = nFirstCol
    End Function

    Private Function FindYMinMax(ByVal dtbDataChart As DataTable, ByVal eGradeType As enumGraphType) As SMinMax
        Dim dtbDataSampling As DataTable = dtbDataChart
        'Dim nStep As Integer = dtbDataChart.Rows.Count \ 1000
        'If dtbDataChart.Rows.Count > 5000 Then   'For increase speed to find min,max
        '    dtbDataSampling = dtbDataChart.Clone
        '    For nSampling As Integer = 0 To dtbDataChart.Rows.Count - 1 Step nStep
        '        dtbDataSampling.Rows.Add(dtbDataChart.Rows(nSampling).ItemArray)
        '    Next nSampling
        'Else
        '    dtbDataSampling = dtbDataChart.Copy
        'End If
        Dim dblYMin As Double = 0
        Dim dblYMax As Double = 0
        Dim nFirstYCol As Integer = FindFirstYColumn(dtbDataChart)
        For nCol As Integer = nFirstYCol To dtbDataSampling.Columns.Count - 1       'Find Y min,max
            Dim strCol As String = dtbDataSampling.Columns(nCol).ColumnName
            Dim objMin As Object = Nothing
            Dim objMax As Object = Nothing

            If eGradeType = enumGraphType.eBoxPlot Then
                objMin = dtbDataSampling.Compute("MIN([" & strCol & "])-3*STDEV([" & strCol & "])", "")
                objMax = dtbDataSampling.Compute("MAX([" & strCol & "])+3*STDEV([" & strCol & "])", "")

                'Dim objSigma As Object = dtbDataSampling.Compute("STDEV([" & strCol & "])", "")
                'Dim objMean As Object = dtbDataSampling.Compute("AVG([" & strCol & "])", "")
                'If objSigma IsNot DBNull.Value Then
                '    objMin = objMin - 3 * objSigma
                '    objMax = objMax + 3 * objSigma
                'End If
            Else
                objMin = dtbDataSampling.Compute("MIN([" & strCol & "])", "")
                objMax = dtbDataSampling.Compute("MAX([" & strCol & "])", "")
            End If

            If objMin IsNot DBNull.Value Then
                If objMin < dblYMin Then dblYMin = objMin
            End If
            If objMax IsNot DBNull.Value Then
                If objMax > dblYMax Then dblYMax = objMax
            End If
            If objMin IsNot DBNull.Value And dblYMin = 0 Then
                dblYMin = objMin
            End If
            If objMax IsNot DBNull.Value And dblYMax = 0 Then
                dblYMax = objMax
            End If
        Next nCol
        If dblYMax = dblYMin Then
            dblYMax = dblYMax + 1
            dblYMin = dblYMin - 1
        End If
        FindYMinMax.dblMax = dblYMax
        FindYMinMax.dblMin = dblYMin
    End Function

    Private Sub picChartAxis_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles m_picChartAxis.DoubleClick
        If m_eGraphType <> enumGraphType.ePieChart Then
            Dim MouseEvent As MouseEventArgs = e
            Try
                If IsMouseInArea(MouseEvent, m_rectYMax) Then
                    Dim strMax As String = InputBox("Input max scale Y", "Y", m_dblYMax)
                    m_dblYMax = CDbl(strMax)
                    m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                ElseIf IsMouseInArea(MouseEvent, m_rectYMin) Then
                    Dim strMin As String = InputBox("Input min scale Y", "Y", m_dblYMin)
                    m_dblYMin = CDbl(strMin)
                    m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                ElseIf IsMouseInArea(MouseEvent, m_rectYMajorUnit) Then
                    Dim strNewUnit As String = InputBox("Input major unit scale Y", "Y", m_dblMajorUnit)
                    m_dblMajorUnit = CDbl(strNewUnit)
                    m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                ElseIf IsMouseInArea(MouseEvent, m_rectXScale) Then
                    Dim strMin As String = ""
                    If m_eGraphType = enumGraphType.eHistrogram Then
                        strMin = InputBox("Input Binning of histrogram", "X", m_intHistrogramStep)
                        m_intHistrogramStep = CDbl(strMin)
                        m_dtbDataShow = PrepareDataByChartType(m_dtbDataChart, m_eGraphType)
                        m_dblYMin = FindYMinMax(m_dtbDataShow, m_eGraphType).dblMin
                        m_dblYMax = FindYMinMax(m_dtbDataShow, m_eGraphType).dblMax
                    Else
                        strMin = InputBox("Input display point", "X", m_intLimitXData)
                        m_intLimitXData = CInt(strMin)
                        ' m_hscScrollGraph.Value = 0
                        If m_intLimitXData > m_dtbDataChart.Rows.Count Then m_intLimitXData = m_dtbDataChart.Rows.Count
                        hscScrollGraph_ValueChanged(sender, e)
                        Dim intMax As Integer = m_dtbDataChart.Rows.Count - m_intLimitXData
                        If intMax < 0 Then intMax = 0
                        m_hscScrollGraph.Maximum = intMax + 9
                    End If
                    m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                Else
                    For nLine As Integer = 0 To m_SDataPoint.Length - 1
                        If IsMouseInArea(MouseEvent, m_SDataPoint(nLine).rectLabel) Then
                            Dim dlgColor As New ColorDialog
                            dlgColor.Color = m_SDataPoint(nLine).penColor.Color
                            dlgColor.FullOpen = True
                            If dlgColor.ShowDialog() = Windows.Forms.DialogResult.OK Then
                                m_SDataPoint(nLine).penColor.Color = dlgColor.Color
                                m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                            End If
                            Exit For
                        End If
                    Next nLine
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Function IsMouseInArea(ByVal MouseEvent As MouseEventArgs, ByVal rectArea As Rectangle) As Boolean
        IsMouseInArea = False
        If MouseEvent.X >= rectArea.Left And MouseEvent.Y >= rectArea.Top And MouseEvent.X <= rectArea.Right And MouseEvent.Y <= rectArea.Bottom Then
            IsMouseInArea = True
        End If
    End Function

    Private Function IsMouseInArea(ByVal MouseEvent As MouseEventArgs, ByVal pntArea As Point, ByVal szSize As Size) As Boolean
        IsMouseInArea = False
        If MouseEvent.X >= pntArea.X And MouseEvent.Y >= pntArea.Y And MouseEvent.X <= pntArea.X + szSize.Width And MouseEvent.Y <= pntArea.Y + szSize.Height Then
            IsMouseInArea = True
        End If
    End Function

    Private Sub picChartAxis_MouseMove(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles m_picChartAxis.MouseMove
        If m_eGraphType <> enumGraphType.ePieChart And e.Button = MouseButtons.Left And IsMouseInArea(e, m_rectChartArea) And m_optPen.Checked = False Then
            Dim imgChart As New Bitmap(m_imgChart)        'Draw cross mark
            Dim grGraph As Graphics = Graphics.FromImage(imgChart)
            Dim penMark As New Pen(Color.Black)
            penMark.DashStyle = Drawing2D.DashStyle.Dot
            grGraph.DrawLine(penMark, e.Location.X, m_rectChartArea.Top, e.Location.X, m_rectChartArea.Bottom)   'Draw vertical
            grGraph.DrawLine(penMark, m_rectChartArea.Left, e.Location.Y, m_rectChartArea.Right, e.Location.Y)     'Draw horizontal
            'Draw Label
            Dim nSize As Integer = 8
            Dim pntArrow(2) As Point
            pntArrow(0).X = e.X
            pntArrow(0).Y = e.Y
            pntArrow(1).X = e.X + (nSize)
            pntArrow(1).Y = e.Y - (nSize)
            pntArrow(2).X = e.X + (nSize * 2)
            pntArrow(2).Y = e.Y - (nSize)
            grGraph.FillPolygon(Brushes.MediumVioletRed, pntArrow)
            Dim sXY As SMinMax = ConvertPointToData(e.X, e.Y, m_rectChartArea, m_dblXMin, m_dblXMax, m_dblYMin, m_dblYMax, m_eGraphType)
            Dim strText As String = "X = " & Format(sXY.dblMin, "0.000") & "   Y = " & Format(sXY.dblMax, "0.000") & Environment.NewLine
            Dim nYUpper As Integer = 0
            Dim nYLower As Integer = 0
            For nYData As Integer = FindFirstYColumn(m_dtbDataShow) To m_dtbDataShow.Columns.Count - 1
                Dim strColName As String = m_dtbDataShow.Columns(nYData).ColumnName
                Dim drUpper() As DataRow = m_dtbDataShow.Select("[" & strColName & "]>" & sXY.dblMax)
                Dim drLower() As DataRow = m_dtbDataShow.Select("[" & strColName & "]<" & sXY.dblMax)
                nYUpper = nYUpper + drUpper.Length
                nYLower = nYLower + drLower.Length
            Next nYData
            strText = strText & "Upper = " & nYUpper & "   Lower = " & nYLower

            Dim szText As Size = grGraph.MeasureString(strText, m_sFontChart.hFontScaleY).ToSize
            Dim rectText As Rectangle
            rectText.Location = pntArrow(1)
            rectText.Y = rectText.Y - szText.Height
            rectText.Height = szText.Height
            rectText.Width = szText.Width + 5
            grGraph.FillRectangle(Brushes.Aquamarine, rectText)
            grGraph.DrawRectangle(Pens.Black, rectText)
            grGraph.DrawString(strText, m_sFontChart.hFontScaleY, Brushes.Black, rectText)
            m_picChartAxis.Image = imgChart
            m_picChartAxis.Refresh()
        ElseIf m_optPen.Checked = True And IsMouseInArea(e, m_picChartAxis.Location, m_picChartAxis.Size) And e.Button = MouseButtons.Left Then
            Dim imgChart As Bitmap = m_picChartAxis.Image 'Draw pen line
            Dim grGraph As Graphics = Graphics.FromImage(imgChart)
            Dim penMark As New Pen(m_optPen.BackColor)
            penMark.Width = 5
            If m_pntLastMouseMove.IsEmpty = False Then
                grGraph.DrawLine(penMark, m_pntLastMouseMove.X, m_pntLastMouseMove.Y, e.X, e.Y)
            Else
                grGraph.FillEllipse(penMark.Brush, e.X, e.Y, 5, 5)
            End If
            m_picChartAxis.Image = imgChart
            m_picChartAxis.Refresh()
            m_pntLastMouseMove.X = e.X
            m_pntLastMouseMove.Y = e.Y
        End If
    End Sub

    Private Sub picChartAxis_MouseDown(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles m_picChartAxis.MouseDown
        Dim MouseEvent As MouseEventArgs = e
        If m_optPen.Checked = True And IsMouseInArea(e, m_picChartAxis.Location, m_picChartAxis.Size) And e.Button = MouseButtons.Left Then
            Dim imgChart As New Bitmap(m_picChartAxis.Image)      'Draw pen point
            Dim grGraph As Graphics = Graphics.FromImage(imgChart)
            Dim penMark As New Pen(m_optPen.BackColor)
            penMark.DashStyle = Drawing2D.DashStyle.Dot
            grGraph.FillEllipse(penMark.Brush, e.X, e.Y, 5, 5)
            m_picChartAxis.Image = imgChart
            m_picChartAxis.Refresh()
        ElseIf e.Button = MouseButtons.Left And m_eGraphType <> enumGraphType.ePieChart Then
            For nLineNum As Integer = 0 To m_SDataPoint.Length - 1  'Draw cross mark
                Dim dtbDataPoint As DataTable = m_SDataPoint(nLineNum).dtbXYValue
                Dim strSelect As String = "((XPoint-" & MouseEvent.X & ")*(XPoint-" & MouseEvent.X & "))+((YPoint-" & MouseEvent.Y & ")*(YPoint-" & MouseEvent.Y & "))<=" & Math.Pow(m_cstDotRadias, 2)
                Dim dtrSelect() As DataRow = dtbDataPoint.Select(strSelect)
                If dtrSelect.Length > 0 Then
                    'm_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
                    Dim imgPic As New Bitmap(m_imgChart)
                    m_picChartAxis.Image = imgPic
                    Dim strGraphName As String = dtbDataPoint.TableName
                    Dim strX As String = dtrSelect(0).Item("XValue").ToString
                    Dim strY As String = Format(dtrSelect(0).Item("YValue"), "0.0000")

                    Dim strPointText As String = strGraphName & ":" & strX & ", " & strY
                    Dim pntClick As Point
                    pntClick.X = dtrSelect(0).Item("XPoint")
                    pntClick.Y = dtrSelect(0).Item("YPoint")
                    Dim grDraw As Graphics = Graphics.FromImage(m_picChartAxis.Image)
                    DrawMarker(grDraw, Pens.Aquamarine, pntClick, m_cstDotRadias, m_SDataPoint(nLineNum).eMarkType)
                    'grDraw.FillEllipse(Brushes.Aquamarine, pntClick.X - m_cstDotRadias, pntClick.Y - m_cstDotRadias, m_cstDotRadias * 2, m_cstDotRadias * 2)

                    Dim rectText As Rectangle
                    rectText.Size = grDraw.MeasureString(strPointText, m_sFontChart.hFontChartDetail).ToSize
                    rectText.X = pntClick.X + m_cstDotRadias - (rectText.Width / 2)
                    rectText.Y = pntClick.Y + m_cstDotRadias

                    Dim rectBackground As Rectangle = rectText
                    rectBackground.Height = rectBackground.Height + 3
                    grDraw.FillRectangle(Brushes.AntiqueWhite, rectBackground)
                    grDraw.DrawRectangle(Pens.DarkGoldenrod, rectBackground)
                    grDraw.DrawString(strPointText, m_sFontChart.hFontChartDetail, Brushes.DarkGoldenrod, rectText)
                    Exit Sub
                End If
            Next nLineNum
        ElseIf e.Button = MouseButtons.Right Then
            Dim imgChart As New Bitmap(m_imgChart)
            m_picChartAxis.Image = imgChart
            m_picChartAxis.Refresh()
        End If
        'DrawGraph(m_dtbDataShow)
    End Sub

    Private Sub DrawYScale(ByVal grDraw As Graphics, ByVal rectGraphArea As Rectangle, ByVal eGraphType As enumGraphType)
        'Draw Y Axis
        Dim strMajorUint As String = "Unit"
        m_rectYMajorUnit.X = rectGraphArea.X - 5
        m_rectYMajorUnit.Y = rectGraphArea.Y / 2 - 10
        m_rectYMajorUnit.Size = grDraw.MeasureString(strMajorUint, m_sFontChart.hFontYLabel).ToSize
        grDraw.DrawRectangle(Pens.CadetBlue, m_rectYMajorUnit)
        DrawRotateText(grDraw, m_sFontChart.hFontYLabel, Brushes.Black, strMajorUint, m_rectYMajorUnit.X, m_rectYMajorUnit.Y, 0, 1)


        Dim drawBrush As New SolidBrush(Color.Black)

        Dim nYScaleNum As Integer = (m_dblYMax - m_dblYMin) / m_dblMajorUnit
        If nYScaleNum > 30 Then
            nYScaleNum = 30
        End If

        For nScaleY As Integer = 0 To nYScaleNum
            Dim pntScale(1) As Point
            pntScale(0).X = rectGraphArea.Left - 10
            pntScale(0).Y = rectGraphArea.Top + (rectGraphArea.Height / nYScaleNum * nScaleY)
            If m_bDrawYScale = True And m_eLineLinkType = enumLineLinkType.eLineHorizontal Then
                pntScale(1).X = rectGraphArea.Right
                pntScale(1).Y = rectGraphArea.Top + (rectGraphArea.Height / nYScaleNum * nScaleY)
            Else
                pntScale(1).X = rectGraphArea.Left + 10
                pntScale(1).Y = rectGraphArea.Top + (rectGraphArea.Height / nYScaleNum * nScaleY)
            End If
            grDraw.DrawLines(Pens.LightGray, pntScale)
            Dim strYText As String = ""
            If eGraphType = enumGraphType.eHistrogram Then
                strYText = Format(m_dblYMax - (nScaleY * (m_dblYMax - m_dblYMin) / nYScaleNum), "0")
            Else
                strYText = Format(m_dblYMax - (nScaleY * (m_dblYMax - m_dblYMin) / nYScaleNum), "0.00")
            End If
            grDraw.DrawString(strYText, m_sFontChart.hFontScaleY, Brushes.Brown, pntScale(0).X - 15, pntScale(0).Y)

            If nScaleY = 0 Then
                m_rectYMax.X = pntScale(0).X - 15
                m_rectYMax.Y = pntScale(0).Y + 2
                m_rectYMax.Size = grDraw.MeasureString(strYText, m_sFontChart.hFontScaleX).ToSize
                grDraw.DrawRectangle(Pens.CadetBlue, m_rectYMax)
            ElseIf nScaleY = nYScaleNum Then
                m_rectYMin.X = pntScale(0).X - 15
                m_rectYMin.Y = pntScale(0).Y + 2
                m_rectYMin.Size = grDraw.MeasureString(strYText, m_sFontChart.hFontScaleX).ToSize
                grDraw.DrawRectangle(Pens.CadetBlue, m_rectYMin)
            End If
        Next nScaleY
        If m_dblYMin < 0 And m_dblYMax > 0 Then    'Draw Y Zero reference
            Dim pntZero(1) As Point
            Dim nY As Integer = (rectGraphArea.Height * m_dblYMax / (m_dblYMax - m_dblYMin)) + rectGraphArea.Y
            pntZero(0).X = rectGraphArea.Left
            pntZero(0).Y = nY
            pntZero(1).X = rectGraphArea.Right
            pntZero(1).Y = nY
            Dim penZero As New Pen(Color.LightGray)
            penZero.DashStyle = Drawing2D.DashStyle.DashDot
            penZero.Width = 2
            grDraw.DrawLines(penZero, pntZero)
            'grDraw.DrawString("0.00", m_sFontChart.hFontScaleY, Brushes.Brown, pntZero(0).X - 15, pntZero(0).Y)
        End If
    End Sub

    Private Sub DrawXScale(ByVal grDraw As Graphics, ByVal dblXMin As Double, ByVal dblXMax As Double, ByVal rectDrawArea As Rectangle, ByVal dtbDataChart As DataTable, ByVal eGraphType As enumGraphType)

        Dim dtcXAxis As DataColumn = dtbDataChart.Columns(0)
        Dim drawBrush As New SolidBrush(Color.Black)

        grDraw.DrawString(dtcXAxis.ColumnName, m_sFontChart.hFontXLabel, drawBrush, rectDrawArea.Right + 20, rectDrawArea.Bottom)
        m_rectXScale.X = rectDrawArea.Right + 20
        m_rectXScale.Y = rectDrawArea.Bottom
        m_rectXScale.Size = grDraw.MeasureString(dtcXAxis.ColumnName, m_sFontChart.hFontXLabel).ToSize
        grDraw.DrawRectangle(Pens.CadetBlue, m_rectXScale)
        Dim intLimitXScale As Integer = m_intLimitXData
        If intLimitXScale > 70 Then intLimitXScale = 70

        If dtcXAxis.Table.Rows.Count < 100 Then intLimitXScale = dtcXAxis.Table.Rows.Count
        Dim nMod As Integer = dtcXAxis.Table.Rows.Count \ intLimitXScale
        Select Case eGraphType
            Case enumGraphType.eXYChart
                Dim dblXRange As Double = dblXMax - dblXMin
                For nLine As Integer = 0 To intLimitXScale - 1
                    Dim intXScale As Integer = rectDrawArea.Left + (rectDrawArea.Width / intLimitXScale * nLine)
                    If m_bDrawXScale Then
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Top - 10, intXScale, rectDrawArea.Bottom)
                    Else
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Bottom - 5, intXScale, rectDrawArea.Bottom + 5)
                    End If
                    'grDraw.DrawString(Format(dblXMin + (nLine * dblXRange / intLimitXScale), "0.00"), m_sFontChart.hFontScaleX, Brushes.Brown, intXScale - 5, rectDrawArea.Bottom + 5, stringVertical)
                    Dim dblRedian As Double = Abs(m_intXScaleDegree) * PI / 180
                    Dim nX As Integer = 0
                    Dim nY As Integer = 0
                    If m_intXScaleDegree < 0 Then
                        nX = intXScale - (m_intXScaleTextHeight * Cos(dblRedian))
                        nY = rectDrawArea.Bottom + (m_intXScaleTextHeight * Sin(dblRedian)) + 10
                    Else
                        nX = intXScale
                        nY = rectDrawArea.Bottom + 5
                    End If
                    If dtcXAxis.DataType Is GetType(DateTime) Then
                        DrawRotateText(grDraw, m_sFontChart.hFontScaleX, Brushes.Brown, Format(DateTime.FromOADate(dblXMin + (nLine * dblXRange / intLimitXScale)), "d-MMM-yy HH:mm:ss"), nX, nY, m_intXScaleDegree, 1)
                    Else
                        DrawRotateText(grDraw, m_sFontChart.hFontScaleX, Brushes.Brown, Format(dblXMin + (nLine * dblXRange / intLimitXScale), "0.0000"), nX, nY, m_intXScaleDegree, 1)
                    End If
                Next nLine
            Case enumGraphType.eBoxPlot
                Dim dtbDistinctLabel As DataTable = m_SDataPoint(0).dtbXYValue.DefaultView.ToTable(True, "XValue", "XPoint")
                nMod = dtbDistinctLabel.Rows.Count \ intLimitXScale
                If nMod = 0 Then nMod = 1
                For nDist As Integer = 0 To dtbDistinctLabel.Rows.Count - 1 Step nMod
                    Dim intXScale As Integer = dtbDistinctLabel.Rows(nDist).Item("XPoint")
                    If m_bDrawXScale Then
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Top - 10, intXScale, rectDrawArea.Bottom)
                    Else
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Bottom - 5, intXScale, rectDrawArea.Bottom + 5)
                    End If
                    'grDraw.DrawString(Format(dblXMin + (nLine * dblXRange / intLimitXScale), "0.00"), m_sFontChart.hFontScaleX, Brushes.Brown, intXScale - 5, rectDrawArea.Bottom + 5, stringVertical)
                    Dim dblRedian As Double = Abs(m_intXScaleDegree) * PI / 180
                    Dim nX As Integer = 0
                    Dim nY As Integer = 0
                    If m_intXScaleDegree < 0 Then
                        nX = intXScale - (m_intXScaleTextHeight * Cos(dblRedian))
                        nY = rectDrawArea.Bottom + (m_intXScaleTextHeight * Sin(dblRedian)) + 10
                    Else
                        nX = intXScale
                        nY = rectDrawArea.Bottom + 5
                    End If
                    Dim strXScale As String = ""
                    If dtcXAxis.DataType Is GetType(DateTime) Then
                        If Not dtbDistinctLabel.Rows(nDist).Item("XValue") Is DBNull.Value Then strXScale = Format(dtbDistinctLabel.Rows(nDist).Item("XValue"), "dd/MMM/yy HH:mm:ss")
                    Else
                        If Not dtbDistinctLabel.Rows(nDist).Item("XValue") Is DBNull.Value Then strXScale = dtbDistinctLabel.Rows(nDist).Item("XValue")
                    End If
                    DrawRotateText(grDraw, m_sFontChart.hFontScaleX, Brushes.Brown, strXScale, nX, nY, m_intXScaleDegree, 1)
                Next nDist
            Case Else
                If nMod = 0 Then nMod = 1
                For nData As Integer = 0 To dtbDataChart.Rows.Count - 1 Step nMod
                    Dim nCount As Integer = dtbDataChart.Rows.Count - 1
                    If nCount = 0 Then nCount = 1
                    Dim intXScale As Integer = (rectDrawArea.Width / (nCount)) * nData + rectDrawArea.X
                    If m_bDrawXScale Then
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Bottom + 5, intXScale, rectDrawArea.Top)
                    Else
                        grDraw.DrawLine(Pens.Gray, intXScale, rectDrawArea.Bottom - 5, intXScale, rectDrawArea.Bottom + 5)
                    End If
                    'grDraw.DrawString(dtbDataChart.Rows(nData).Item(0).ToString, m_sFontChart.hFontScaleX, Brushes.Brown, intXScale - 5, rectDrawArea.Bottom + 5, stringVertical)
                    Dim dblRedian As Double = Abs(m_intXScaleDegree) * PI / 180
                    Dim nX As Integer = 0
                    Dim nY As Integer = 0
                    If m_intXScaleDegree < 0 Then
                        nX = intXScale - (m_intXScaleTextHeight * Cos(dblRedian))
                        nY = rectDrawArea.Bottom + (m_intXScaleTextHeight * Sin(dblRedian)) + 10
                    Else
                        nX = intXScale
                        nY = rectDrawArea.Bottom + 5
                    End If
                    Dim strX As String = dtbDataChart.Rows(nData).Item(0).ToString
                    If dtcXAxis.DataType.ToString = "System.DateTime" Then
                        strX = Format(dtbDataChart.Rows(nData).Item(0), "dd/MMM/yy HH:mm:ss")
                    End If
                    DrawRotateText(grDraw, m_sFontChart.hFontScaleX, Brushes.Brown, strX, nX, nY, m_intXScaleDegree, 1)
                Next nData
        End Select
    End Sub

    Private Sub DrawXDetail(ByVal grDraw As Graphics, ByVal dblXMin As Double, ByVal dblXMax As Double, ByVal rectDrawArea As Rectangle, ByVal dtbDataChart As DataTable, ByVal eGraphType As enumGraphType)

        If dtbDataChart.Columns(enumColumnType.eSection).DataType Is Type.GetType("System.String") Or dtbDataChart.Columns(enumColumnType.eSection).DataType Is Type.GetType("System.DateTime") Then
            Dim dtcXAxis As DataColumn = dtbDataChart.Columns(enumColumnType.eSection)
            Dim drawBrush As New SolidBrush(Color.Black)

            Dim intLimitXScale As Integer = m_intLimitXData
            If intLimitXScale > 70 Then intLimitXScale = 70

            If dtcXAxis.Table.Rows.Count < 100 Then intLimitXScale = dtcXAxis.Table.Rows.Count
            Dim nMod As Integer = dtcXAxis.Table.Rows.Count \ intLimitXScale
            If eGraphType = enumGraphType.eHistrogram Then

            Else
                If nMod = 0 Then nMod = 1
                Dim sXDetail(0) As SXDetail
                sXDetail(0).strXString = dtbDataChart.Rows(0).Item(enumColumnType.eSection).ToString
                sXDetail(0).pntStart.X = rectDrawArea.X
                sXDetail(0).pntStart.Y = rectDrawArea.Top
                sXDetail(0).pntEnd.X = rectDrawArea.Right
                sXDetail(0).pntEnd.Y = rectDrawArea.Top
                Dim strStringLast As String = dtbDataChart.Rows(0).Item(enumColumnType.eSection).ToString
                For nData As Integer = 0 To dtbDataChart.Rows.Count - 1 Step nMod
                    Dim nCount As Integer = dtbDataChart.Rows.Count - 1
                    If nCount = 0 Then nCount = 1
                    Dim intXScale As Integer = (rectDrawArea.Width / (nCount)) * nData + rectDrawArea.X
                    Dim nX As Integer = 0
                    Dim nY As Integer = 0
                    nX = intXScale
                    nY = rectDrawArea.Top
                    Dim strStringNow As String = dtbDataChart.Rows(nData).Item(enumColumnType.eSection).ToString
                    If strStringLast.ToUpper <> strStringNow.ToUpper Then
                        ReDim Preserve sXDetail(sXDetail.Length)
                        sXDetail(sXDetail.Length - 1).strXString = strStringNow
                        sXDetail(sXDetail.Length - 1).pntStart.X = nX
                        sXDetail(sXDetail.Length - 1).pntStart.Y = nY
                        sXDetail(sXDetail.Length - 1).pntEnd.X = rectDrawArea.Right
                        sXDetail(sXDetail.Length - 1).pntEnd.Y = rectDrawArea.Top

                        sXDetail(sXDetail.Length - 2).pntEnd.X = nX
                        sXDetail(sXDetail.Length - 2).pntEnd.Y = nY
                        strStringLast = strStringNow
                    End If
                Next nData
                Dim PenDetail As New Pen(Color.Black)
                PenDetail.DashStyle = Drawing2D.DashStyle.DashDot
                For nDetail As Integer = 0 To sXDetail.Length - 1
                    Dim strDraw As String = sXDetail(nDetail).strXString
                    Dim SizeString As SizeF = grDraw.MeasureString(strDraw, m_sFontChart.hFontScaleX)
                    Dim nStringLen As Integer = SizeString.Width
                    Dim pntStart As Point = sXDetail(nDetail).pntStart
                    Dim pntEnd As Point = sXDetail(nDetail).pntEnd
                    Dim nXCenter As Integer = ((pntEnd.X + pntStart.X) / 2) - (nStringLen / 2)

                    'grDraw.DrawLine(PenDetail, pntStart.X, pntStart.Y - 5, pntStart.X, pntStart.Y + rectDrawArea.Height)
                    grDraw.DrawLine(PenDetail, pntEnd.X, pntEnd.Y - 5, pntEnd.X, pntEnd.Y + rectDrawArea.Height)
                    'grDraw.DrawLine(Pens.Black, pntStart.X, pntStart.Y, pntEnd.X, pntEnd.Y)
                    grDraw.DrawString(strDraw, m_sFontChart.hFontScaleX, Brushes.Brown, nXCenter, pntStart.Y - SizeString.Height - 1)
                Next nDetail
            End If
        End If
    End Sub

    Private Function ConvertPointToData(ByVal intX As Integer, ByVal intY As Integer, ByVal rectDrawArea As Rectangle, ByVal dblXMin As Double, ByVal dblXMax As Double, ByVal dblYMin As Double, ByVal dblYMax As Double, ByVal eGraphType As enumGraphType) As SMinMax
        If eGraphType = enumGraphType.eXYChart Then
            ConvertPointToData.dblMin = (intX - rectDrawArea.Left) / rectDrawArea.Width * (dblXMax - dblXMin) + dblXMin
        Else
            ConvertPointToData.dblMin = (intX - rectDrawArea.Left) / rectDrawArea.Width * (m_dtbDataShow.Rows.Count)
        End If
        ConvertPointToData.dblMax = (rectDrawArea.Height - intY + rectDrawArea.Top) / rectDrawArea.Height * (dblYMax - dblYMin) + dblYMin
    End Function

    Private Function ConvertDataToPoint(ByVal sValuePoint As SValueAndPoint, ByVal rectDrawArea As Rectangle, ByVal dtbDataChart As DataTable, ByVal strColumnName As String, ByVal dblXMin As Double, ByVal dblXMax As Double, ByVal dblYMin As Double, ByVal dblYMax As Double, ByVal eGraphType As enumGraphType) As SValueAndPoint
        Dim dtcXAxis As DataColumn
        'If eGraphType = enumGraphType.eBoxPlot Then
        dtcXAxis = dtbDataChart.Columns(enumColumnType.eRealValue)
        'Else
        'dtcXAxis = dtbDataChart.Columns(0)
        'End If
        Dim dtcYAxis As DataColumn = dtbDataChart.Columns(strColumnName)
        Dim dblYRange As Double = dblYMax - dblYMin
        If dblYRange < 0.0001 Then dblYRange = 0.0001
        Dim dblXRange As Double = dblXMax - dblXMin
        Dim nDataNum As Integer = dtcYAxis.Table.Rows.Count
        Dim pntData(dtcYAxis.Table.Rows.Count - 1) As Point
        Dim dtbXYValue As New DataTable(dtcYAxis.ColumnName)
        dtbXYValue.Columns.Add("XValue", dtbDataChart.Columns(0).DataType)
        dtbXYValue.Columns.Add("YValue", Type.GetType("System.Double"))
        dtbXYValue.Columns.Add("XPoint", Type.GetType("System.Int64"))
        dtbXYValue.Columns.Add("YPoint", Type.GetType("System.Int64"))

        For nData As Integer = 0 To nDataNum - 1
            If dtcYAxis.Table.Rows(nData).Item(dtcYAxis.ColumnName) Is DBNull.Value Then
                pntData(nData).X = 0
                pntData(nData).Y = 0
            Else
                'If eGraphType = enumGraphType.eXYChart Or eGraphType = enumGraphType.eBoxPlot Then
                Dim objXValue As Object = dtcXAxis.Table.Rows(nData).Item(dtcXAxis.ColumnName)
                'If dtcXAxis.DataType Is GetType(DateTime) Then
                'dblXValue = CDate(dtcXAxis.Table.Rows(nData).Item(dtcXAxis.ColumnName)).ToOADate
                'Else
                If objXValue Is DBNull.Value Then
                    pntData(nData).X = 0
                    pntData(nData).Y = 0
                Else
                    Dim dblXValue As Double = dtcXAxis.Table.Rows(nData).Item(dtcXAxis.ColumnName)
                    Dim dblYValue As Double = dtcYAxis.Table.Rows(nData).Item(dtcYAxis.ColumnName)
                    'End If

                    pntData(nData) = Convert2ScreenPoint(rectDrawArea, dblXMin, dblXMax, dblYMin, dblYMax, dblXValue, dblYValue)

                    'pntData(nData).X = (rectDrawArea.Width * (dblXValue - dblXMin) / dblXRange) + rectDrawArea.X
                    'pntData(nData).Y = (rectDrawArea.Height * dblYValue / dblYRange) + rectDrawArea.Y
                End If
            End If

            dtbXYValue.Rows.Add(dtcXAxis.Table.Rows(nData).Item(0), dtcYAxis.Table.Rows(nData).Item(dtcYAxis.ColumnName), pntData(nData).X, pntData(nData).Y)
        Next nData
        ConvertDataToPoint.pntPlot = pntData
        ConvertDataToPoint.dtbXYValue = dtbXYValue
        ConvertDataToPoint.bWriteDot = sValuePoint.bWriteDot
        ConvertDataToPoint.penColor = sValuePoint.penColor
        ConvertDataToPoint.strColumnName = sValuePoint.strColumnName
        ConvertDataToPoint.eMarkType = sValuePoint.eMarkType
    End Function

    Private Function Convert2ScreenPoint(ByVal rectDrawArea As Rectangle, ByVal dblXMin As Double, ByVal dblXMax As Double, ByVal dblYMin As Double, ByVal dblYMax As Double, ByVal dblXData As Double, ByVal dblYData As Double) As Point
        Dim pntScreen As Point
        Dim dblYRange As Double = dblYMax - dblYMin
        Dim dblXRange As Double = dblXMax - dblXMin

        pntScreen.X = (rectDrawArea.Width * (dblXData - dblXMin) / dblXRange) + rectDrawArea.X
        pntScreen.Y = (rectDrawArea.Height * (dblYMax - dblYData) / dblYRange) + rectDrawArea.Y
        Convert2ScreenPoint = pntScreen
    End Function

    Private Sub SetControlSize()
        m_intFormWidth = m_frmChart.Width
        m_intFormHeight = m_frmChart.Height

        m_picChartAxis.Width = m_frmChart.Width - 40
        m_picChartAxis.Height = m_frmChart.Height - 90
        m_picChartAxis.BorderStyle = BorderStyle.FixedSingle

        m_intPicWidth = m_picChartAxis.Width
        m_intPicHeight = m_picChartAxis.Height

        m_hscScrollGraph.Width = m_picChartAxis.Width / 3
        m_hscScrollGraph.Height = 20
        m_hscScrollGraph.Left = m_picChartAxis.Width - m_hscScrollGraph.Width
        m_hscScrollGraph.Top = m_picChartAxis.Bottom + 5

        m_optPen.Location = m_hscScrollGraph.Location
        m_optPen.Left = 10
    End Sub

    Private Function DrawChartBitmap(ByVal dtbDataChart As DataTable) As Bitmap
        Dim strChartText As String = dtbDataChart.TableName

        If Not m_picChartAxis Is Nothing Then
            If Not m_picChartAxis.Image Is Nothing Then m_picChartAxis.Image.Dispose()
            m_picChartAxis.Image = Nothing
        End If
        If Not m_imgChart Is Nothing Then
            m_imgChart.Dispose()
            m_imgChart = Nothing
        End If
        'SetControlSize()

        If m_intPicWidth <= 0 Or m_intPicHeight <= 0 Then Return Nothing

        Dim imgChart As New Bitmap(m_intPicWidth, m_intPicHeight)
        Dim rectImage As Rectangle
        rectImage.Width = imgChart.Width
        rectImage.Height = imgChart.Height
        rectImage.X = 0
        rectImage.Y = 0

        Dim intOfsetX As Integer = 40
        Dim intOfsetY As Integer = 50
        'm_intXScaleTextHeight = 0
        Dim grGraph As Graphics = Graphics.FromImage(imgChart)

        If dtbDataChart.Rows.Count > 0 And m_intXScaleTextHeight = 0 Then     'Find max X string
            Dim dtbTemp As DataTable = dtbDataChart.Clone
            Dim nStep As Integer = dtbDataChart.Rows.Count \ 1000
            If nStep = 0 Then nStep = 1
            For nSampling As Integer = 0 To dtbDataChart.Rows.Count - 1 Step nStep
                dtbTemp.Rows.Add(dtbDataChart.Rows(nSampling).ItemArray)
            Next nSampling
            Dim strX As String = dtbTemp.Columns(0).ColumnName
            dtbTemp.Columns.Add("LenOfX", Type.GetType("System.Int32"), "LEN(CONVERT([" & strX & "],System.String))")

            Dim dtrMaxLen() As DataRow = dtbTemp.Select("LenOfX=MAX(LenOfX)")
            'Dim ojbMaxLen As Object = dtbTemp.Compute("MAX(LEN(CONVERT([" & strX & "],System.String)))", "")
            If dtrMaxLen.Length > 0 Then
                Dim strMaxLen As String = dtrMaxLen(0).Item(strX)
                If dtbTemp.Columns(0).DataType Is GetType(String) Then
                ElseIf dtbTemp.Columns(0).DataType Is GetType(DateTime) Then
                    strMaxLen = Format(dtrMaxLen(0).Item(strX), "dd/MMM/yy HH:mm:ss")
                Else
                    strMaxLen = Format(CDbl(strMaxLen), "0.0000")
                End If
                Dim dblRedian As Double = m_intXScaleDegree * PI / 180
                m_intXScaleTextHeight = grGraph.MeasureString(strMaxLen, m_sFontChart.hFontScaleX).ToSize.Width * Abs(Sin(dblRedian))
            Else
                m_intXScaleTextHeight = intOfsetY
            End If
        End If

        Dim rectGraphArea As Rectangle
        rectGraphArea.Width = imgChart.Width - intOfsetX * 4.5
        rectGraphArea.Height = imgChart.Height - m_intXScaleTextHeight - intOfsetY - 15
        rectGraphArea.X = 60
        rectGraphArea.Y = intOfsetY
        grGraph.DrawRectangle(Pens.Black, rectGraphArea)
        m_rectChartArea = rectGraphArea
        grGraph.FillRectangle(Brushes.PaleGoldenrod, rectImage)
        grGraph.FillRectangle(Brushes.WhiteSmoke, rectGraphArea)
        grGraph.DrawRectangle(Pens.Black, rectGraphArea)

        grGraph.DrawString(strChartText, m_sFontChart.hFontChartLabel, Brushes.Black, (rectImage.Width / 2) - (grGraph.MeasureString(strChartText, m_sFontChart.hFontChartLabel).Width / 2), 2)

        Dim dtcXAxis As DataColumn = dtbDataChart.Columns(0)

        Dim AxisPen As New Pen(Brushes.Gray, 4)
        AxisPen.EndCap = Drawing2D.LineCap.Triangle
        grGraph.DrawLine(AxisPen, rectGraphArea.Left, rectGraphArea.Bottom + 2, rectGraphArea.Right + 5, rectGraphArea.Bottom + 2) 'Draw X Axis

        grGraph.DrawLine(AxisPen, rectGraphArea.Left, rectGraphArea.Bottom + 4, rectGraphArea.Left, rectGraphArea.Top - 15)     'Draw Y Axis

        DrawYScale(grGraph, rectGraphArea, m_eGraphType)       'Draw Y Scale

        If m_dblYMin <> m_dblYMax Then
            For nCol As Integer = 0 To m_SDataPoint.Length - 1        'Draw Detail
                Dim strColName As String = m_SDataPoint(nCol).strColumnName
                m_SDataPoint(nCol) = ConvertDataToPoint(m_SDataPoint(nCol), rectGraphArea, dtbDataChart, strColName, m_dblXMin, m_dblXMax, m_dblYMin, m_dblYMax, m_eGraphType)
                m_SDataPoint(nCol).rectLabel.X = rectGraphArea.Right + 5
                m_SDataPoint(nCol).rectLabel.Y = rectGraphArea.Top + (m_sFontChart.hFontYLabel.SizeInPoints * nCol) + (10 * nCol)
                Dim sizeString As SizeF = grGraph.MeasureString(strColName, m_sFontChart.hFontYLabel)
                m_SDataPoint(nCol).rectLabel.Width = m_intPicWidth - m_SDataPoint(nCol).rectLabel.X - 5
                m_SDataPoint(nCol).rectLabel.Height = sizeString.Height
                grGraph.DrawString(strColName, m_sFontChart.hFontYLabel, m_SDataPoint(nCol).penColor.Brush, m_SDataPoint(nCol).rectLabel.X + (m_cstDotRadias * 4), m_SDataPoint(nCol).rectLabel.Y)
                grGraph.DrawRectangle(m_SDataPoint(nCol).penColor, m_SDataPoint(nCol).rectLabel)
                Dim pntMark As Point
                pntMark.X = m_SDataPoint(nCol).rectLabel.X + (sizeString.Height / 2)
                pntMark.Y = m_SDataPoint(nCol).rectLabel.Y + 5
                DrawMarker(grGraph, m_SDataPoint(nCol).penColor, pntMark, m_cstDotRadias, m_SDataPoint(nCol).eMarkType)
                grGraph.DrawLine(m_SDataPoint(nCol).penColor, pntMark.X - m_cstDotRadias, pntMark.Y, pntMark.X + m_cstDotRadias, pntMark.Y)
            Next nCol
            PlotValue(grGraph, rectGraphArea, m_SDataPoint)
        End If

        DrawXScale(grGraph, m_dblXMin, m_dblXMax, rectGraphArea, dtbDataChart, m_eGraphType)
        DrawXDetail(grGraph, m_dblXMin, m_dblXMax, rectGraphArea, dtbDataChart, m_eGraphType)

        'grGraph.DrawImage(m_picChartAxis.Image, 0, 0)         
        'm_picChartAxis.Image = imgChart
        DrawChartBitmap = imgChart
        m_imgChart = New Bitmap(imgChart)
        grGraph.Dispose()
        grGraph = Nothing
    End Function

    Private Sub PlotValue(ByVal grGraph As Graphics, ByVal rectGraphArea As Rectangle, ByVal sArValuePointData() As SValueAndPoint)

        Dim penConnect As New Pen(Color.Sienna, 2)

        For nLine As Integer = 0 To sArValuePointData.Length - 1
            Dim sValuePointData As SValueAndPoint = sArValuePointData(nLine)
            Dim penChart As Pen = sValuePointData.penColor
            Dim bIsDotPlot As Boolean = sValuePointData.bWriteDot
            Dim strColName As String = sValuePointData.strColumnName
            Dim eMarkType As enumMarkerType = sValuePointData.eMarkType
            Select Case m_eGraphType
                Case enumGraphType.eColumnGraph
                    Dim pntPlot() As Point = sValuePointData.pntPlot
                    Dim nBarWidth As Integer = (rectGraphArea.Width / (pntPlot.Length - 1)) / (m_SDataPoint.Length)
                    If nBarWidth = 0 Then nBarWidth = 1

                    For nDot As Integer = 0 To pntPlot.Length - 1
                        If pntPlot(nDot).IsEmpty = False Then
                            Dim rectBar As Rectangle
                            rectBar.X = pntPlot(nDot).X + (nBarWidth * (nLine - m_SDataPoint.Length + 1))
                            rectBar.Y = pntPlot(nDot).Y
                            rectBar.Width = nBarWidth
                            rectBar.Height = rectGraphArea.Bottom - rectBar.Y
                            grGraph.FillRectangle(penChart.Brush, rectBar)
                        End If
                    Next nDot
                    'grGraph.DrawCurve(penChart, pntPlot )
                Case CChartControl.enumGraphType.eLineSeriesGraph
                    Dim pntPlot() As Point = sValuePointData.pntPlot
                    For nDot As Integer = 0 To pntPlot.Length - 1
                        If pntPlot(nDot).IsEmpty = False Then
                            If nDot > 0 And m_eLineLinkType = enumLineLinkType.eLineHorizontal Then
                                For nBackData As Integer = nDot - 1 To 0 Step -1
                                    Dim pntBackData As Point = pntPlot(nBackData)
                                    If pntBackData.IsEmpty = False Then
                                        grGraph.DrawLine(penChart, pntBackData.X, pntBackData.Y, pntPlot(nDot).X, pntPlot(nDot).Y)
                                        Exit For
                                    End If
                                Next nBackData
                                'If pntPlot(nDot - 1).IsEmpty = False Then
                                '    grGraph.DrawLine(penChart, pntPlot(nDot - 1).X, pntPlot(nDot - 1).Y, pntPlot(nDot).X, pntPlot(nDot).Y)
                                'End If
                            ElseIf m_eLineLinkType = enumLineLinkType.eLineVertical Then
                                For nData As Integer = 0 To m_SDataPoint.Length - 2
                                    Dim pnt0 As Point = m_SDataPoint(nData).pntPlot(nDot)
                                    Dim pnt1 As Point = m_SDataPoint(nData + 1).pntPlot(nDot)
                                    If pnt0.IsEmpty = False And pnt1.IsEmpty = False Then
                                        grGraph.DrawLine(penConnect, pnt0, pnt1)
                                    End If
                                Next nData
                            End If
                            'If sValuePointData.bWriteDot = True Then
                            DrawMarker(grGraph, penChart, pntPlot(nDot), m_cstDotRadias, eMarkType)
                            'End If
                        End If
                    Next nDot

                Case enumGraphType.eXYChart
                    Dim pntPlot() As Point = sValuePointData.pntPlot
                    For nDot As Integer = 0 To pntPlot.Length - 1
                        'grGraph.FillEllipse(penChart.Brush, sPlotData.pntPlot(nDot).X - m_cstDotRadias, sPlotData.pntPlot(nDot).Y - m_cstDotRadias, m_cstDotRadias * 2, m_cstDotRadias * 2)
                        DrawMarker(grGraph, penChart, pntPlot(nDot), m_cstDotRadias, eMarkType)
                    Next nDot
                    If m_sLinearParam.dblIntercept <> 0 And m_sLinearParam.dblRSqr <> 0 And m_sLinearParam.dblSlope <> 0 Then
                        Dim dtbLinear As New DataTable
                        dtbLinear.Columns.Add("X", System.Type.GetType("System.Double"))
                        dtbLinear.Columns.Add("Y", System.Type.GetType("System.Double"))
                        dtbLinear.Rows.Add(m_dblXMin, (m_sLinearParam.dblSlope * m_dblXMin) + m_sLinearParam.dblIntercept)
                        dtbLinear.Rows.Add(m_dblXMax, (m_sLinearParam.dblSlope * m_dblXMax) + m_sLinearParam.dblIntercept)
                        Dim pntLinear(1) As Point
                        pntLinear(0) = Convert2ScreenPoint(rectGraphArea, m_dblXMin, m_dblXMax, m_dblYMin, m_dblYMax, m_dblXMin, (m_sLinearParam.dblSlope * m_dblXMin) + m_sLinearParam.dblIntercept)
                        pntLinear(1) = Convert2ScreenPoint(rectGraphArea, m_dblXMin, m_dblXMax, m_dblYMin, m_dblYMax, m_dblXMax, (m_sLinearParam.dblSlope * m_dblXMax) + m_sLinearParam.dblIntercept)
                        If m_eLineLinkType = enumLineLinkType.eLineHorizontal Then
                            grGraph.DrawLines(Pens.DarkViolet, pntLinear)
                        End If
                        Dim strIntercept As String = ""
                        If m_sLinearParam.dblIntercept >= 0 Then
                            strIntercept = "+" & Format(m_sLinearParam.dblIntercept, "0.0000")
                        Else
                            strIntercept = Format(m_sLinearParam.dblIntercept, "0.0000")
                        End If
                        Dim strLinear As String = "y=" & Format(m_sLinearParam.dblSlope, "0.0000") & "x" & strIntercept & ",R-Square=" & Format(m_sLinearParam.dblRSqr, "0.0000")
                        grGraph.DrawString(strLinear, m_sFontChart.hFontLinearEquation, Brushes.BlueViolet, rectGraphArea.Left, rectGraphArea.Top - m_sFontChart.hFontLinearEquation.Height)
                    End If
                Case enumGraphType.eHistrogram
                    Dim pntPlot() As Point = sValuePointData.pntPlot
                    Dim nBarWidth As Integer = (rectGraphArea.Width / (pntPlot.Length - 1)) / (m_SDataPoint.Length)
                    If nBarWidth = 0 Then nBarWidth = 1
                    For nDot As Integer = 0 To pntPlot.Length - 1
                        Dim rectBar As Rectangle
                        rectBar.X = pntPlot(nDot).X + (nBarWidth * (nLine - m_SDataPoint.Length + 1))
                        rectBar.Y = pntPlot(nDot).Y
                        rectBar.Width = nBarWidth
                        rectBar.Height = rectGraphArea.Bottom - rectBar.Y
                        grGraph.FillRectangle(penChart.Brush, rectBar)
                    Next nDot
                Case enumGraphType.eBoxPlot
                    Dim sBoxplotParam() As SBoxPlotParameter = GetBoxPlotParameter(sValuePointData.dtbXYValue)
                    DrawBoxPlotInfo(grGraph, sBoxplotParam, penChart, eMarkType)
            End Select

        Next nLine
    End Sub

    Private Sub DrawBoxPlotInfo(ByVal grGraph As Graphics, ByVal sBoxPlotParam() As SBoxPlotParameter, ByVal penChart As Pen, ByVal eMarkType As enumMarkerType)
        Dim nBoxWidth As Integer = 20
        Dim penBox As New Pen(Color.Black)
        penBox.Width = 2
        For nInfo As Integer = 0 To sBoxPlotParam.Length - 1
            Dim dtbOutlier As DataTable = sBoxPlotParam(nInfo).dtbOutlier
            For nDot As Integer = 0 To dtbOutlier.Rows.Count - 1
                Dim pntOutlier As New Point(dtbOutlier.Rows(nDot).Item("XPoint"), dtbOutlier.Rows(nDot).Item("YPoint"))
                DrawMarker(grGraph, penChart, pntOutlier, m_cstDotRadias, enumMarkerType.eStar)
            Next nDot

            Dim intXCenter As Integer = sBoxPlotParam(nInfo).intX
            Dim intX1 As Integer = intXCenter - nBoxWidth / 2
            Dim intX2 As Integer = intXCenter + nBoxWidth / 2
            grGraph.DrawLine(penBox, intX1, sBoxPlotParam(nInfo).intMax, intX2, sBoxPlotParam(nInfo).intMax)    'Draw max cap
            grGraph.DrawLine(penBox, intX1, sBoxPlotParam(nInfo).intMin, intX2, sBoxPlotParam(nInfo).intMin)      'Draw min cap

            grGraph.DrawLine(penBox, intXCenter, sBoxPlotParam(nInfo).intMin, intXCenter, sBoxPlotParam(nInfo).intMax)      'Draw min to max 

            Dim rectBox As Rectangle
            rectBox.X = intX1
            rectBox.Y = sBoxPlotParam(nInfo).int75thPercent
            rectBox.Width = intX2 - intX1
            rectBox.Height = sBoxPlotParam(nInfo).int25thPercent - sBoxPlotParam(nInfo).int75thPercent
            grGraph.FillRectangle(Brushes.LightGray, rectBox)
            grGraph.DrawLine(penBox, intX1, sBoxPlotParam(nInfo).intMedian, intX2, sBoxPlotParam(nInfo).intMedian)      'Draw Median
            grGraph.DrawRectangle(penBox, rectBox)     'Draw Box

        Next nInfo
    End Sub

    Private Function GetPieChart(ByVal dtbData As DataTable, ByVal nWidth As Integer, ByVal nHeight As Integer) As Bitmap

        Dim hLabelFont As Font = m_sFontChart.hFontChartLabel
        Dim hDetailFont As Font = m_sFontChart.hFontChartDetail

        Dim bmChart As New Bitmap(nWidth, nHeight)
        Dim grChart As Graphics = Graphics.FromImage(bmChart)
        grChart.FillRectangle(Brushes.Beige, 0, 0, nWidth, nHeight)

        Dim nWidthDetail As Integer = 300
        nWidth = nWidth - nWidthDetail
        If nWidth > nHeight Then nWidth = nHeight
        If nHeight > nWidth Then nHeight = nWidth
        'Dim nDataCount As Integer = dtbMEWDelta.Rows.Count

        Dim rectProduct As Rectangle
        rectProduct.X = 0
        rectProduct.Y = 0
        rectProduct.Width = nWidth
        rectProduct.Height = nWidth

        Dim penCircle As New Pen(Color.WhiteSmoke, 1)
        grChart.DrawPie(penCircle, rectProduct, -90, 360)
        grChart.FillPie(Brushes.White, rectProduct, -90, 360)

        Dim rectDetail As Rectangle
        rectDetail.Width = 12
        rectDetail.Height = 12
        rectDetail.X = rectProduct.Right + 1
        rectDetail.Y = rectProduct.Y

        Dim nStartAngle As Single = -90
        Dim nEndAngle As Single = 0
        Dim nTotalData As Integer = dtbData.Compute("SUM(CountItem)", "")
        Dim PenPie() As Pen = InitPen(dtbData.Rows.Count)
        For nItem As Integer = 0 To dtbData.Rows.Count - 1
            rectDetail.Y = rectDetail.Top + hDetailFont.Height + 4
            nStartAngle = nStartAngle + nEndAngle
            Dim strDetail As String = dtbData.Rows(nItem).Item(0)
            Dim nCount As Integer = dtbData.Rows(nItem).Item(1)
            nEndAngle = nCount * 360 / nTotalData
            'If strDetail.Contains("PASS") Then
            '    PenPie(nItem).Color = Color.Green
            'ElseIf strDetail.Contains("REJECT") Then
            '    PenPie(nItem).Color = Color.Red
            'ElseIf strDetail.Contains("FAIL") Then
            '    PenPie(nItem).Color = Color.Brown
            'End If
            grChart.FillPie(PenPie(nItem).Brush, rectProduct, nStartAngle, nEndAngle)
            If dtbData.Rows.Count <> 1 Then grChart.DrawPie(penCircle, rectProduct, nStartAngle, nEndAngle)
            grChart.FillRectangle(PenPie(nItem).Brush, rectDetail)
            Dim strFirstText As String = strDetail & "=" & nCount & "(" & Format(nCount / nTotalData * 100, "0.0") & "%)"
            grChart.DrawString(strFirstText, hDetailFont, Brushes.Black, rectDetail.Right, rectDetail.Top)
        Next nItem
        GetPieChart = bmChart

    End Function

    Private Structure SBoxPlotParameter      'Represent in point
        Dim intX As Integer
        Dim intMax As Integer
        Dim intMin As Integer
        Dim intMedian As Integer
        Dim intMean As Integer
        Dim int25thPercent As Integer
        Dim int75thPercent As Integer
        Dim dtbOutlier As DataTable
    End Structure

    Private Function GetBoxPlotParameter(ByVal dtbData As DataTable) As SBoxPlotParameter()
        dtbData.DefaultView.RowFilter = ""
        Dim dtbDistinctX As DataTable = dtbData.DefaultView.ToTable(True, "XPoint")
        Dim sBoxPlot(dtbDistinctX.Rows.Count - 1) As SBoxPlotParameter
        For nX As Integer = 0 To dtbDistinctX.Rows.Count - 1
            Dim nXPoint As Integer = dtbDistinctX.Rows(nX).Item("XPoint")
            dtbData.DefaultView.RowFilter = "XPoint=" & nXPoint
            dtbData.DefaultView.Sort = "YValue ASC"
            Dim dtbGroup As DataTable = dtbData.DefaultView.ToTable

            Dim objSigma As Object = dtbGroup.Compute("STDEV(YPoint)", "")
            Dim dblSigma As Double = 0
            If Not objSigma Is DBNull.Value Then dblSigma = objSigma
            Dim dblMean As Double = dtbGroup.Compute("AVG(YPoint)", "")
            sBoxPlot(nX).intMean = dblMean
            sBoxPlot(nX).intX = nXPoint
            sBoxPlot(nX).intMin = dblMean + 3 * dblSigma
            sBoxPlot(nX).intMax = dblMean - 3 * dblSigma

            dtbData.DefaultView.RowFilter = "XPoint=" & nXPoint & " AND YPoint>=" & sBoxPlot(nX).intMax & " AND YPoint<=" & sBoxPlot(nX).intMin
            dtbData.DefaultView.Sort = "YValue ASC"
            Dim dtbValue As DataTable = dtbData.DefaultView.ToTable

            Dim nDiv As Integer = dtbValue.Rows.Count \ 2
            Dim nMod As Integer = dtbValue.Rows.Count Mod 2
            If nMod = 0 Then      'Even
                sBoxPlot(nX).intMedian = (dtbValue.Rows(nDiv - 1).Item("YPoint") + dtbValue.Rows(nDiv).Item("YPoint")) / 2
            Else 'Odd
                sBoxPlot(nX).intMedian = dtbValue.Rows(nDiv).Item("YPoint")
            End If

            Dim n25Row As Integer = dtbValue.Rows.Count / 4
            Dim n75Row As Integer = dtbValue.Rows.Count * 3 / 4

            sBoxPlot(nX).int25thPercent = dtbValue.Rows(n25Row).Item("YPoint")
            sBoxPlot(nX).int75thPercent = dtbValue.Rows(n75Row - 1).Item("YPoint")

            dtbGroup.DefaultView.RowFilter = "YPoint>" & sBoxPlot(nX).intMin & " OR YPoint<" & sBoxPlot(nX).intMax
            sBoxPlot(nX).dtbOutlier = dtbGroup.DefaultView.ToTable
        Next nX
        dtbData.DefaultView.RowFilter = ""
        GetBoxPlotParameter = sBoxPlot
    End Function

    Public Function GetBitmap() As Bitmap
        If m_imgChart Is Nothing Then
            If m_eGraphType = enumGraphType.ePieChart Then
                m_imgChart = GetPieChart(m_dtbDataShow, m_picChartAxis.Width, m_picChartAxis.Height)
            Else
                m_imgChart = DrawChartBitmap(m_dtbDataShow)
            End If
        End If
        Return m_imgChart
    End Function

    Public Sub ShowChart()
        CreateForm(m_dtbDataChart.TableName, m_intFormWidth, m_intFormHeight)

        m_picChartAxis.Image = GetBitmap()
    
        m_frmChart.Show()
        m_frmChart.Cursor = Cursors.Arrow
    End Sub

    Public Property XScaleMinMax() As SMinMax
        Get
            XScaleMinMax.dblMax = m_dblXMax
            XScaleMinMax.dblMin = m_dblXMin
        End Get
        Set(ByVal value As SMinMax)
            m_dblXMax = value.dblMax
            m_dblXMin = value.dblMin
        End Set
    End Property

    Public Property YScaleMinMax() As SMinMax
        Get
            YScaleMinMax.dblMax = m_dblYMax
            YScaleMinMax.dblMin = m_dblYMin
        End Get
        Set(ByVal value As SMinMax)
            m_dblYMax = value.dblMax
            m_dblYMin = value.dblMin
        End Set
    End Property

    Public ReadOnly Property GetDataChart() As DataTable
        Get
            Return m_dtbDataChart
        End Get
    End Property

    Private Sub Form_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        SetControlSize()
        If m_eGraphType = enumGraphType.ePieChart Then
            m_picChartAxis.Image = GetPieChart(m_dtbDataShow, m_picChartAxis.Width, m_picChartAxis.Height)
        Else
            m_picChartAxis.Image = DrawChartBitmap(m_dtbDataShow)
        End If
    End Sub

    'Private Sub Form_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs)
    '    DrawChart(m_dtbDataShow)
    'End Sub

    Private Sub DrawRotateText(ByVal gr As Graphics, ByVal _
    the_font As Font, ByVal the_brush As Brush, ByVal txt _
    As String, ByVal x As Integer, ByVal y As Integer, _
    ByVal angle_degrees As Single, ByVal y_scale As Single)
        ' Translate the point to the origin.
        gr.TranslateTransform(-x, -y, _
            Drawing2D.MatrixOrder.Append)

        ' Rotate through the angle.
        gr.RotateTransform(angle_degrees, _
            Drawing2D.MatrixOrder.Append)

        ' Scale vertically by a factor of Tan(angle).
        Dim angle_radians As Double = angle_degrees * PI / 180
        gr.ScaleTransform(1, y_scale, _
            Drawing2D.MatrixOrder.Append)

        ' Find the inverse angle and rotate back.
        'angle_radians = Math.Atan(y_scale * Tan(angle_radians))
        'angle_degrees = CSng(angle_radians * 180 / PI)
        'gr.RotateTransform(-angle_degrees, _
        '    Drawing2D.MatrixOrder.Append)

        ' Translate the origin back to the point.
        gr.TranslateTransform(x, y, _
            Drawing2D.MatrixOrder.Append)

        ' Draw the text.
        gr.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
        gr.DrawString(txt, the_font, the_brush, x, y)
        gr.ResetTransform()
    End Sub

    Private Sub DrawAngledText(ByVal gr As Graphics, ByVal _
  the_font As Font, ByVal the_brush As Brush, ByVal txt _
  As String, ByVal x As Integer, ByVal y As Integer, _
  ByVal angle_degrees As Single, ByVal y_scale As Single)
        ' Translate the point to the origin.
        gr.TranslateTransform(-x, -y, _
            Drawing2D.MatrixOrder.Append)

        ' Rotate through the angle.
        gr.RotateTransform(angle_degrees, _
            Drawing2D.MatrixOrder.Append)

        ' Scale vertically by a factor of Tan(angle).
        Dim angle_radians As Double = angle_degrees * PI / 180
        gr.ScaleTransform(1, y_scale, _
            Drawing2D.MatrixOrder.Append)

        ' Find the inverse angle and rotate back.
        angle_radians = Math.Atan(y_scale * Tan(angle_radians))
        angle_degrees = CSng(angle_radians * 180 / PI)
        gr.RotateTransform(-angle_degrees, _
            Drawing2D.MatrixOrder.Append)

        ' Translate the origin back to the point.
        gr.TranslateTransform(x, y, _
            Drawing2D.MatrixOrder.Append)

        ' Draw the text.
        gr.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
        gr.DrawString(txt, the_font, the_brush, x, y)
        gr.ResetTransform()
    End Sub

    Private Sub m_picChartAxis_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles m_picChartAxis.MouseUp
        m_pntLastMouseMove = Nothing
    End Sub
End Class
