
''' <summary>
'''This class use for analyze data
''' </summary>
''' 
Public Class CDataAnalyzer

    Public Structure SLinearParameter
        Dim dblSlope As Double
        Dim dblIntercept As Double
        Dim dblRSqr As Double
    End Structure

    ''' <summary>
    ''' 'Calculate datatable to  moving average table by n-Point.
    ''' </summary>
    ''' <param name="dtbData">Data to convert</param>
    ''' <param name="nPointNum">Moving AVG N-Point</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function ConvertToMovingAverage(ByVal dtbData As DataTable, ByVal nPointNum As Integer) As DataTable
        Dim dtbMovingMean As DataTable = dtbData.Clone

        For nDataIndex As Integer = nPointNum - 1 To dtbData.Rows.Count - 1
            Dim dblMean As Double = 0
            dtbMovingMean.Rows.Add()
            For nCol As Integer = 0 To dtbData.Columns.Count - 1
                dtbMovingMean.Columns(nCol).ReadOnly = False
                If Not dtbData.Columns(nCol).DataType Is GetType(String) And Not dtbData.Columns(nCol).DataType Is GetType(DateTime) Then
                    Dim strColName As String = dtbData.Columns(nCol).ColumnName
                    Dim nPoint As Integer = nPointNum - 1
                    Dim nMovingIndex As Integer = nDataIndex
                    While nPoint > -1 And nMovingIndex > -1
                        If Not dtbData.Rows(nMovingIndex).Item(strColName) Is DBNull.Value Then
                            dblMean = dblMean + dtbData.Rows(nMovingIndex).Item(strColName)
                            nPoint = nPoint - 1
                        End If
                        nMovingIndex = nMovingIndex - 1
                    End While
                    If nPoint = -1 Then
                        dblMean = dblMean / nPointNum
                        dtbMovingMean.Rows(dtbMovingMean.Rows.Count - 1).Item(strColName) = Format(dblMean, "0.0000")
                    End If
                Else
                    dtbMovingMean.Rows(dtbMovingMean.Rows.Count - 1).Item(nCol) = dtbData.Rows(nDataIndex).Item(nCol)
                End If
            Next nCol
        Next nDataIndex

        ConvertToMovingAverage = dtbMovingMean
    End Function
    ''' <summary>
    '''Calculate Linear regression then return slope,intercept and R-Square
    ''' </summary>
    Public Function CalculateLinearRegression(ByVal strX As String, ByVal strY As String, ByVal dtbDataChart As DataTable) As SLinearParameter
        Dim strOldFilter As String = dtbDataChart.DefaultView.RowFilter
        dtbDataChart.DefaultView.RowFilter = "[" & strX & "] IS NOT NULL AND [" & strY & "] IS NOT NULL"
        Dim dtbTemp As DataTable = dtbDataChart.DefaultView.ToTable
        dtbDataChart.DefaultView.RowFilter = strOldFilter

        If dtbTemp.Rows.Count > 1 Then
            dtbTemp.Columns.Add("SumX", GetType(Double), "SUM([" & strX & "])")
            dtbTemp.Columns.Add("SumY", GetType(Double), "SUM([" & strY & "])")

            dtbTemp.Columns.Add("MeanX", GetType(Double), "Avg([" & strX & "])")
            dtbTemp.Columns.Add("MeanY", GetType(Double), "Avg([" & strY & "])")

            dtbTemp.Columns.Add("XY", GetType(Double), "([" & strX & "]*[" & strY & "])")
            dtbTemp.Columns.Add("Sum_XY", GetType(Double), "SUM(XY)")


            dtbTemp.Columns.Add("XX", GetType(Double), "([" & strX & "]*[" & strX & "])")
            dtbTemp.Columns.Add("Sum_XX", GetType(Double), "SUM(XX)")

            Dim dtrRow As DataRow = dtbTemp.Rows(0)
            Dim nData As Integer = dtbTemp.Rows.Count
            Dim sLinear As SLinearParameter
            sLinear.dblSlope = (dtrRow.Item("SUM_XY") - (dtrRow.Item("SUMX") * dtrRow.Item("SumY") / nData)) / _
                                                     (dtrRow.Item("Sum_XX") - (nData * Math.Pow(dtrRow.Item("MeanX"), 2)))
            sLinear.dblIntercept = dtrRow.Item("MeanY") - (sLinear.dblSlope * dtrRow.Item("MeanX"))

            If Not Double.IsNaN(sLinear.dblSlope) And Not Double.IsNaN(sLinear.dblIntercept) And Not Double.IsInfinity(sLinear.dblSlope) And Not Double.IsInfinity(sLinear.dblIntercept) Then
                'Find R-squar
                dtbTemp.Columns.Add("YLine", GetType(Double), sLinear.dblSlope & "*[" & strX & "]+" & sLinear.dblIntercept)
                dtbTemp.Columns.Add("YDiffError", GetType(Double), "([" & strY & "]" & "-YLine)*([" & strY & "]" & "-YLine)")
                dtbTemp.Columns.Add("SSError", GetType(Double), "SUM(YDiffError)")

                dtbTemp.Columns.Add("YDiffTotal", GetType(Double), "([" & strY & "]" & "-MeanY)*([" & strY & "]" & "-MeanY)")
                dtbTemp.Columns.Add("SSTotal", GetType(Double), "SUM(YDiffTotal)")
                dtbTemp.Columns.Add("SSRegression", GetType(Double), "SSTotal-SSError")

                sLinear.dblRSqr = dtrRow.Item("SSRegression") / dtrRow.Item("SSTotal")
            End If
            CalculateLinearRegression = sLinear
        End If
    End Function


    Private Function AnalyzeHistrogram(ByVal dtbDataChart As DataTable, ByVal nHistrogramStep As Integer) As DataTable
        Dim dtbHistrogram As New DataTable(dtbDataChart.TableName)

        dtbHistrogram.Columns.Add(dtbDataChart.Columns(0).ColumnName, GetType(Double)) 'X
        Dim nYColumn As Integer = FindFirstYColumn(dtbDataChart)
        Dim strFirstYColName As String = dtbDataChart.Columns(nYColumn).ColumnName
        For nYData As Integer = nYColumn To dtbDataChart.Columns.Count - 1
            Dim strCol As String = dtbDataChart.Columns(nYData).ColumnName
            dtbHistrogram.Columns.Add(strCol, GetType(Int32))  'Y
        Next nYData

        Dim dblMax As Double = dtbDataChart.Compute("MAX([" & strFirstYColName & "])", "")
        Dim dblMin As Double = dtbDataChart.Compute("MIN([" & strFirstYColName & "])", "")

        If dblMin = dblMax Then
            dblMin = dblMin - 1
            dblMax = dblMax + 1
        End If

        Dim dblDiscreteStep As Double = (dblMax - dblMin) / nHistrogramStep

        Dim dblStep As Double = dblMin
        While dblStep <= dblMax
            Dim dblStepMax As Double = dblStep + dblDiscreteStep
            dtbHistrogram.Rows.Add()
            dtbHistrogram.Rows(dtbHistrogram.Rows.Count - 1).Item(0) = Format(dblStep, "0.0000")
            For nYData As Integer = FindFirstYColumn(dtbDataChart) To dtbDataChart.Columns.Count - 1
                Dim strCol As String = dtbDataChart.Columns(nYData).ColumnName
                Dim strFilter As String = "[" & strCol & "]>=" & dblStep & " AND [" & strCol & "]<" & dblStepMax
                Dim dtrFilter() As DataRow = dtbDataChart.Select(strFilter)
                dtbHistrogram.Rows(dtbHistrogram.Rows.Count - 1).Item(strCol) = dtrFilter.Length
            Next nYData
            dblStep = dblStepMax
        End While
        AnalyzeHistrogram = dtbHistrogram
    End Function

    Private Function FindFirstYColumn(ByVal dtbData As DataTable) As Integer
        Dim nFirstCol As Integer = 0
        For nCol As Integer = 0 To dtbData.Columns.Count - 1
            Dim colDataType As Type = dtbData.Columns(nCol).DataType
            If Not colDataType Is GetType(String) And Not colDataType Is GetType(DateTime) Then
                nFirstCol = nCol
                Exit For
            End If
        Next nCol
        If nFirstCol = 0 Then nFirstCol = 1
        FindFirstYColumn = nFirstCol
    End Function

    Public Function Convert2IntervalAverage(ByVal dtbRawData As DataTable, ByVal nAverageGroup As Integer) As DataTable

        Dim dtbMean As DataTable = dtbRawData.Clone
        Dim dtbTemp As DataTable = dtbRawData.Clone
        For nData As Integer = dtbRawData.Rows.Count - 1 To 0 Step -1
            dtbTemp.Rows.Add(dtbRawData.Rows(nData).ItemArray) 'Add data to temp
            If (nData) Mod nAverageGroup = 0 Or (nData = 0 And dtbTemp.Rows.Count >= nAverageGroup / 2) Then     '
                Dim drMean As DataRow = dtbMean.Rows.Add(dtbTemp.Rows(0).ItemArray)   'Add info to mean table
                For nCol As Integer = FindFirstYColumn(dtbRawData) To dtbRawData.Columns.Count - 1
                    Dim strColName As String = dtbRawData.Columns(nCol).ColumnName
                    drMean.Item(strColName) = dtbTemp.Compute("AVG([" & strColName & "])", "")       'Compute mean
                Next nCol
                dtbTemp.Rows.Clear()
            End If
        Next nData
        Convert2IntervalAverage = dtbMean
    End Function


End Class
