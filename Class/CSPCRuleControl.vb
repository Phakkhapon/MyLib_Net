'Analyze only first column
'Detect from last data
Public Class CSPCRuleControl
    Private Const m_cstMaxSPCPoints = 15

    Public Enum enumDataType
        eRawData = 0
        eMovingMean
        eDiscreteMean
    End Enum

    Public Structure SControlParameter
        Dim dblSigma As Double
        Dim dblMean As Double
        Dim dblUCL As Double
        Dim dblLCL As Double
    End Structure

    'Private m_dtbData As DataTable      'Original data
    Public m_dtbDataControlChart As DataTable        'Keap data for plot graph and analyze
    Public m_sXBarParam As SControlParameter

    Public Sub New(ByVal dtbDataControl As DataTable, ByVal strColName As String, ByVal eDataType As enumDataType)
        Dim dtbFilter As DataTable = FilterOutNullValue(strColName, dtbDataControl)
        Select Case eDataType
            Case enumDataType.eRawData
                'm_dtbData = dtbFilter
            Case enumDataType.eMovingMean
                dtbFilter = ConvertToMovingMean(strColName, dtbFilter, 5)
        End Select
        m_sXBarParam = GetXBarControlParameter(dtbFilter, strColName)
    End Sub

    Private Function FilterOutNullValue(ByVal strColName As String, ByVal dtbData As DataTable) As DataTable
        Dim strFilter As String = "[" & strColName & "] IS NOT NULL"
        dtbData.DefaultView.RowFilter = strFilter
        Dim dtbFilter As DataTable = dtbData.DefaultView.ToTable(False, strColName)
        dtbData.DefaultView.RowFilter = ""

        dtbFilter.Columns.Add("X", GetType(Int32))
        dtbFilter.Columns("X").SetOrdinal(0)
        For nData As Integer = 0 To dtbFilter.Rows.Count - 1
            dtbFilter.Rows(nData).Item("X") = nData
        Next nData
        FilterOutNullValue = dtbFilter
    End Function

    Private Function ConvertToMovingMean(ByVal strColName As String, ByVal dtbData As DataTable, ByVal nPointNum As Integer)
        Dim dtbMovingMean As New DataTable
        dtbMovingMean.Columns.Add("X", GetType(Int32))
        dtbMovingMean.Columns("X").AutoIncrement = True
        dtbMovingMean.Columns.Add(strColName, GetType(Double))

        Dim dblSum As Double = 0
        For nData As Integer = 0 To dtbData.Rows.Count - 1
            dblSum = dblSum + dtbData.Rows(nData).Item(strColName)
            If nData >= nPointNum - 1 Then
                Dim drRowAdd As DataRow = dtbMovingMean.Rows.Add()
                drRowAdd.Item(strColName) = dblSum / nPointNum
                dblSum = dblSum - dtbData.Rows(nData - nPointNum + 1).Item(strColName)
            End If
        Next nData
        ConvertToMovingMean = dtbMovingMean
    End Function

    Private Function GetXBarControlParameter(ByVal dtbData As DataTable, ByVal strParam As String) As SControlParameter
        If dtbData.Rows.Count > 1 Then
            Dim dtbSigma As DataTable = dtbData.Copy
            Dim dblSigma As Double = dtbSigma.Compute("StDev([" & strParam & "])", "")
            Dim dblMean As Double = dtbSigma.Compute("Avg([" & strParam & "])", "")

            dtbSigma.Columns.Add("Sigma", GetType(Double), dblSigma)
            dtbSigma.Columns.Add("Mean", GetType(Double), dblMean)
            dtbSigma.Columns.Add("UCL", GetType(Double), dblMean + 3 * dblSigma)
            dtbSigma.Columns.Add("LCL", GetType(Double), dblMean - 3 * dblSigma)
            dtbSigma.Columns.Add("1Sigma", GetType(Double), dblMean + dblSigma)
            dtbSigma.Columns.Add("2Sigma", GetType(Double), dblMean + 2 * dblSigma)
            dtbSigma.Columns.Add("-1Sigma", GetType(Double), dblMean - dblSigma)
            dtbSigma.Columns.Add("-2Sigma", GetType(Double), dblMean - 2 * dblSigma)

            GetXBarControlParameter.dblSigma = dblSigma
            GetXBarControlParameter.dblMean = dblMean
            GetXBarControlParameter.dblUCL = dblMean + 3 * dblSigma
            GetXBarControlParameter.dblLCL = dblMean - 3 * dblSigma

            m_dtbDataControlChart = dtbSigma
            dtbSigma.Dispose()
            dtbSigma = Nothing
        End If
    End Function

    Public Function IsDetectSPCRule(ByVal nRule As Integer) As Boolean
        Dim bRes As Boolean = False
        Select Case nRule
            Case 1
                bRes = IsDetectSPCRule1()
            Case 2
                bRes = IsDetectSPCRule2()
            Case 3
                bRes = IsDetectSPCRule3()
            Case 4
                bRes = IsDetectSPCRule4()
            Case 5
                bRes = IsDetectSPCRule5()
            Case 6
                bRes = IsDetectSPCRule6()
            Case 7
                bRes = IsDetectSPCRule7()
            Case 8
                bRes = IsDetectSPCRule8()
            Case Else
                bRes = False
        End Select
        IsDetectSPCRule = bRes
    End Function

    Public Function IsDetectSPCRule1() As Boolean
        'One of data out +-3 Sigma
        IsDetectSPCRule1 = False
        If m_dtbDataControlChart IsNot Nothing Then
            If m_dtbDataControlChart.Rows.Count > 0 Then
                Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                Dim dblValue As Double = m_dtbDataControlChart.Rows(nRowCount - 1).Item(strColName)
                If dblValue > m_sXBarParam.dblUCL Then IsDetectSPCRule1 = True
            End If
        End If
    End Function

    Public Function IsDetectSPCRule2() As Boolean
        '9 Point continue upper/lower mean
        IsDetectSPCRule2 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 9
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim dblMean As Double = m_sXBarParam.dblMean
                Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dtrRowPos() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND [" & strColName & "]>" & dblMean)
                    Dim dtrRowNeg() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND [" & strColName & "]<" & dblMean)
                    If dtrRowPos.Length = nCheckPoint Or dtrRowNeg.Length = nCheckPoint Then
                        IsDetectSPCRule2 = True
                    End If
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule3() As Boolean
        '6 Point continue to increment/decrement
        IsDetectSPCRule3 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 6
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                Dim nPointPos As Integer = 0
                Dim nPointNeg As Integer = 0
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblLastValue As Double = m_dtbDataControlChart.Rows(nRowCount - nCheckPoint).Item(strColName)
                    For nData As Integer = nRowCount - nCheckPoint To nRowCount - 1
                        Dim dblValue As Double = m_dtbDataControlChart.Rows(nData).Item(strColName)
                        If dblValue > dblLastValue Then
                            nPointPos = nPointPos + 1
                            nPointNeg = 0
                        Else
                            nPointPos = 0
                            nPointNeg = nPointNeg + 1
                        End If
                        If nPointNeg >= nCheckPoint Or nPointPos >= nCheckPoint Then
                            IsDetectSPCRule3 = True
                            Exit Function
                        End If
                        dblLastValue = dblValue
                    Next nData
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule4() As Boolean
        '14 points up->down->up->...
        IsDetectSPCRule4 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 14
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblLastValue As Double = m_dtbDataControlChart.Rows(nRowCount - nCheckPoint).Item(strColName)
                    Dim dblValue As Double = m_dtbDataControlChart.Rows(nRowCount - nCheckPoint + 1).Item(strColName)
                    Dim nCountPoint As Integer = 0
                    Dim bCompare As Boolean
                    If dblValue > dblLastValue Then    'this for initial compare
                        bCompare = True
                    Else
                        bCompare = False
                    End If
                    For nData As Integer = nRowCount - nCheckPoint To nRowCount - 1
                        dblValue = m_dtbDataControlChart.Rows(nData).Item(strColName)
                        bCompare = bCompare Xor True
                        If bCompare = True Then
                            If dblValue > dblLastValue Then
                                nCountPoint = nCountPoint + 1
                            End If
                        Else
                            If dblValue < dblLastValue Then
                                nCountPoint = nCountPoint + 1
                            End If
                        End If
                        If nCountPoint >= nCheckPoint Then
                            IsDetectSPCRule4 = True
                            Exit Function
                        End If
                        dblLastValue = dblValue
                    Next nData
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule5() As Boolean
        '2 from 3 points out 2-3 sigma
        IsDetectSPCRule5 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 3
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblSigma As Double = m_sXBarParam.dblSigma
                    Dim dblMean As Double = m_sXBarParam.dblMean
                    Dim dblUCL As Double = m_sXBarParam.dblUCL
                    Dim dblLCL As Double = m_sXBarParam.dblLCL
                    Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                    Dim dtrRowPos() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nCheckPoint & " AND [" & strColName & "]>" & dblMean + (2 * dblSigma) & " AND [" & strColName & "]<" & dblUCL)
                    Dim dtrRowNeg() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nCheckPoint & " AND [" & strColName & "]<" & dblMean - (2 * dblSigma) & " AND [" & strColName & "]<" & dblLCL)
                    If dtrRowPos.Length = nCheckPoint - 1 Or dtrRowNeg.Length = nCheckPoint - 1 Then
                        IsDetectSPCRule5 = True
                    End If
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule6() As Boolean
        '4 from 5 points     out +-1sigma
        IsDetectSPCRule6 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 5
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblSigma As Double = m_sXBarParam.dblSigma
                    Dim dblMean As Double = m_sXBarParam.dblMean
                    Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                    Dim dtrRowPos() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND " & strColName & ">" & dblMean + dblSigma)
                    Dim dtrRowNeg() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND " & strColName & "<" & dblMean - dblSigma)
                    If dtrRowPos.Length = nCheckPoint - 1 Or dtrRowNeg.Length = nCheckPoint - 1 Then
                        IsDetectSPCRule6 = True
                    End If
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule7() As Boolean
        '15 points cotinue in +-1 sigma  
        IsDetectSPCRule7 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 15
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblSigma As Double = m_sXBarParam.dblSigma
                    Dim dblMean As Double = m_sXBarParam.dblMean
                    Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                    Dim dtrRow() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND ([" & strColName & "]<" & dblMean + dblSigma & " AND [" & strColName & "]>" & dblMean - dblSigma & ")")
                    If dtrRow.Length = nCheckPoint Then
                        IsDetectSPCRule7 = True
                    End If
                End If
            End If
        End If
    End Function

    Public Function IsDetectSPCRule8() As Boolean
        '8 points out off +-1sigma  
        IsDetectSPCRule8 = False
        If m_dtbDataControlChart IsNot Nothing Then
            Dim nCheckPoint As Integer = 8
            If m_dtbDataControlChart.Rows.Count >= nCheckPoint Then
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                If nRowCount >= nCheckPoint Then
                    Dim dblSigma As Double = m_sXBarParam.dblSigma
                    Dim dblMean As Double = m_sXBarParam.dblMean
                    Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                    Dim dtrRowPos() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND " & strColName & ">" & dblMean + dblSigma)
                    Dim dtrRowNeg() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nCheckPoint & " AND [X]<" & nRowCount & " AND " & strColName & "<" & dblMean - dblSigma)
                    If dtrRowPos.Length = nCheckPoint Or dtrRowNeg.Length = nCheckPoint Then
                        IsDetectSPCRule8 = True
                    End If
                End If
            End If
        End If
    End Function

    Public Function IsPattenRun(ByVal nPatternPoint As Integer) As Integer
        IsPattenRun = 0
        'Dim nCheckPoint As Integer = nPatternPoint
        If m_dtbDataControlChart IsNot Nothing Then
            If m_dtbDataControlChart.Rows.Count >= nPatternPoint Then
                Dim nRowCount As Integer = m_dtbDataControlChart.Rows.Count
                Dim strColName As String = m_dtbDataControlChart.Columns(1).ColumnName
                Dim dtrRowPos() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nPatternPoint & " AND [X]<" & nRowCount & " AND " & strColName & ">0")
                Dim dtrRowNeg() As DataRow = m_dtbDataControlChart.Select("[X]>=" & nRowCount - nPatternPoint & " AND [X]<" & nRowCount & " AND " & strColName & "<0")
                If dtrRowPos.Length = nPatternPoint Then
                    IsPattenRun = 1
                ElseIf dtrRowNeg.Length = nPatternPoint Then
                    IsPattenRun = -1
                Else
                    IsPattenRun = 0
                End If
            End If
        End If
    End Function

    Public Sub PlotChart()
        If Not m_dtbDataControlChart Is Nothing Then
            Dim dtbTemp As DataTable = m_dtbDataControlChart.Copy
            dtbTemp.Columns.Remove("Sigma")
            m_dtbDataControlChart.TableName = "Real-Time SPC Chart"
            Dim clsChart As New CChartControl(1000, 500, dtbTemp, CChartControl.enumGraphType.eLineSeriesGraph, False, False)
            clsChart.ShowChart()
        End If
    End Sub

End Class
