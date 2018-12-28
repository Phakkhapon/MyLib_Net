Imports MySql.Data.MySqlClient

Public Class CGetYieldSDETOCR
    Private m_myRawConn As MySqlConnection

    Public Sub New(ByVal myRawConn As MySqlConnection)
        m_myRawConn = myRawConn
    End Sub

    Public Function GetYieldOCR(ByVal strProduct As String, ByVal dtbTester As DataTable, ByVal dtbHeader As DataTable, ByVal dtStart As DateTime, _
    ByVal dtEnd As DateTime) As DataTable

        Dim strSQL As String
        strSQL = "SELECT max(A.test_time) Date_Time,"
        strSQL = strSQL & "A.Tester,"
        'strSQL = strSQL & "A.Grade_Rev,"
        'strSQL = strSQL & "A.MediaSN,"
        'strSQL = strSQL & "A.OprID,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "A.Spec,"
        'strSQL = strSQL & "A.Assy,"
        strSQL = strSQL & "SUM(A.GradeName NOT LIKE '%ALIGN%' AND A.GradeName NOT LIKE '%NOHGA%' AND LENGTH(A.OCR_SN)>0) Total,"
        strSQL = strSQL & "SUM(A.Hga_SN=A.OCR_SN AND A.GradeName NOT LIKE '%ALIGN%' AND A.GradeName NOT LIKE '%NOHGA%' AND A.OCR_SN NOT LIKE '%?%' AND LENGTH(A.OCR_SN)>0) AccuracyCount,"
        strSQL = strSQL & "SUM(A.OCR_SN NOT LIKE '%?%' AND A.GradeName NOT LIKE '%ALIGN%' AND A.GradeName NOT LIKE '%NOHGA%' AND LENGTH(A.OCR_SN)>0) ReadabilityCount,"

        strSQL = strSQL & "SUM(IF(MID(Lot,5,1)='1' OR MID(Lot,5,1)='3' OR MID(Lot,5,1)='5' OR MID(Lot,5,1)='7' OR MID(Lot,5,1)='A' OR MID(Lot,5,1)='C' OR MID(Lot,5,1)='E' OR MID(Lot,5,1)='G' OR MID(Lot,5,1)='J' OR MID(Lot,5,1)='L' OR MID(Lot,5,1)='Q' OR MID(Lot,5,1)='S',"
        strSQL = strSQL & "ASCII(MID(Hga_SN,5,1))<> ASCII(MID(lot,5,1)) AND ASCII(MID(Hga_SN,5,1))<>ASCII(MID(lot,5,1))+1,IF(MID(Lot,5,1)='9',MID(Hga_SN,5,1)<>'9' AND MID(Hga_SN,5,1)<>'0' ,IF(MID(Hga_SN,5,1)='N',MID(Hga_SN,5,1)<>'N' AND MID(Hga_SN,5,1)<>'P',0))) AND Hga_SN NOT LIKE '%?%' AND LENGTH(Hga_SN)=8 AND Shoe='1') 'ErrorHgaSN5Th.Shoe1',"
        strSQL = strSQL & "SUM(IF(MID(Lot,5,1)='1' OR MID(Lot,5,1)='3' OR MID(Lot,5,1)='5' OR MID(Lot,5,1)='7' OR MID(Lot,5,1)='A' OR MID(Lot,5,1)='C' OR MID(Lot,5,1)='E' OR MID(Lot,5,1)='G' OR MID(Lot,5,1)='J' OR MID(Lot,5,1)='L' OR MID(Lot,5,1)='Q' OR MID(Lot,5,1)='S',"
        strSQL = strSQL & "ASCII(MID(Hga_SN,5,1))<> ASCII(MID(lot,5,1)) AND ASCII(MID(Hga_SN,5,1))<>ASCII(MID(lot,5,1))+1,IF(MID(Lot,5,1)='9',MID(Hga_SN,5,1)<>'9' AND MID(Hga_SN,5,1)<>'0' ,IF(MID(Hga_SN,5,1)='N',MID(Hga_SN,5,1)<>'N' AND MID(Hga_SN,5,1)<>'P',0))) AND Hga_SN NOT LIKE '%?%' AND LENGTH(Hga_SN)=8 AND Shoe='2') 'ErrorHgaSN5Th.Shoe2',"

        strSQL = strSQL & "SUM(IF(MID(Lot,5,1)='1' OR MID(Lot,5,1)='3' OR MID(Lot,5,1)='5' OR MID(Lot,5,1)='7' OR MID(Lot,5,1)='A' OR MID(Lot,5,1)='C' OR MID(Lot,5,1)='E' OR MID(Lot,5,1)='G' OR MID(Lot,5,1)='J' OR MID(Lot,5,1)='L' OR MID(Lot,5,1)='Q' OR MID(Lot,5,1)='S',"
        strSQL = strSQL & "ASCII(MID(OCR_SN,5,1))<> ASCII(MID(lot,5,1)) AND ASCII(MID(OCR_SN,5,1))<>ASCII(MID(lot,5,1))+1,IF(MID(Lot,5,1)='9',MID(OCR_SN,5,1)<>'9' AND MID(OCR_SN,5,1)<>'0' ,IF(MID(OCR_SN,5,1)='N',MID(OCR_SN,5,1)<>'N' AND MID(OCR_SN,5,1)<>'P',0))) AND OCR_SN NOT LIKE '%?%' AND LENGTH(OCR_SN)=8 AND Shoe='1') 'ErrorOCR5Th.Shoe1',"
        strSQL = strSQL & "SUM(IF(MID(Lot,5,1)='1' OR MID(Lot,5,1)='3' OR MID(Lot,5,1)='5' OR MID(Lot,5,1)='7' OR MID(Lot,5,1)='A' OR MID(Lot,5,1)='C' OR MID(Lot,5,1)='E' OR MID(Lot,5,1)='G' OR MID(Lot,5,1)='J' OR MID(Lot,5,1)='L' OR MID(Lot,5,1)='Q' OR MID(Lot,5,1)='S',"
        strSQL = strSQL & "ASCII(MID(OCR_SN,5,1))<> ASCII(MID(lot,5,1)) AND ASCII(MID(OCR_SN,5,1))<>ASCII(MID(lot,5,1))+1,IF(MID(Lot,5,1)='9',MID(OCR_SN,5,1)<>'9' AND MID(OCR_SN,5,1)<>'0' ,IF(MID(OCR_SN,5,1)='N',MID(OCR_SN,5,1)<>'N' AND MID(OCR_SN,5,1)<>'P',0))) AND OCR_SN NOT LIKE '%?%' AND LENGTH(OCR_SN)=8 AND Shoe='2') 'ErrorOCR5Th.Shoe2' "

        strSQL = strSQL & " FROM db_" & strProduct & ".tabdetail_header A "
        strSQL = strSQL & " WHERE (A.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
        strSQL = strSQL & " AND ("
        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            If nTester <> dtbTester.Rows.Count - 1 Then
                strSQL = strSQL & " a.tester='" & dtbTester.Rows(nTester).Item("tester") & "' OR "
            Else
                strSQL = strSQL & "a.tester='" & dtbTester.Rows(nTester).Item("tester") & "') "
            End If
        Next nTester
        strSQL = strSQL & " AND A.Spec NOT LIKE 'C%'"
        strSQL = strSQL & "GROUP BY a.Tester,a.Lot,a.Spec "
        strSQL = strSQL & "ORDER BY a.Tester,test_time_bigint;"
        Dim clsGetYield As New CMySQL
        Dim dtbGetOCR As DataTable = clsGetYield.CommandMySqlDataTable(strSQL, m_myRawConn)

        dtbGetOCR.Columns.Add("%Accuracy", System.Type.GetType("System.Double"))
        dtbGetOCR.Columns("%Accuracy").Expression = "IIF(ReadabilityCount<>0,AccuracyCount/ReadabilityCount*100,0)"
        dtbGetOCR.Columns.Add("%Readability", System.Type.GetType("System.Double"))
        dtbGetOCR.Columns("%Readability").Expression = "IIF(Total<>0,ReadabilityCount/Total*100,0)"
        dtbGetOCR.Columns.Add("%Capability", System.Type.GetType("System.Double"))
        dtbGetOCR.Columns("%Capability").Expression = "[%Accuracy] * [%Readability]/100"
        'For nData As Integer = 0 To dtbGetOCR.Rows.Count - 1
        '    'If dtbGetOCR.Rows(nData).Item("AccuracyCount") > dtbGetOCR.Rows(nData).Item("ReadabilityCount") Then
        '    '    dtbGetOCR.Rows(nData).Item("AccuracyCount") = dtbGetOCR.Rows(nData).Item("ReadabilityCount")
        '    'End If
        '    dtbGetOCR.Rows(nData).Item("%Accuracy") = dtbGetOCR.Rows(nData).Item("AccuracyCount") / dtbGetOCR.Rows(nData).Item("ReadabilityCount") * 100
        '    dtbGetOCR.Rows(nData).Item("%Readability") = dtbGetOCR.Rows(nData).Item("ReadabilityCount") / dtbGetOCR.Rows(nData).Item("Total") * 100
        '    dtbGetOCR.Rows(nData).Item("%Capability") = dtbGetOCR.Rows(nData).Item("%Accuracy") * dtbGetOCR.Rows(nData).Item("%Readability") / 100
        'Next nData
        GetYieldOCR = dtbGetOCR
    End Function

End Class
