
Imports MySql.Data.MySqlClient

Public Class CGetSetting

    Public Function GetAutoX2LotByProduct(ByVal strProduct As String, ByVal mySqlConn As MySqlConnection) As DataTable
        'Dim strSQL As String
        'strSQL = "SELECT A.param_rttc EDB,"
        'strSQL = strSQL & "ParamMDB,"
        'strSQL = strSQL & "B.Sigma,"
        'strSQL = strSQL & "B.AutoCalSigma,"
        'strSQL = strSQL & "AdjustX2Lot,"
        'strSQL = strSQL & "AdjustOption,"
        'strSQL = strSQL & "AdjustType,"
        'strSQL = strSQL & "RefType,"
        'strSQL = strSQL & "TargetSpec,"
        'strSQL = strSQL & "Weight3Sigma_Neg '-3Sigma',"
        ''strSQL = strSQL & "B.Sigma*(-3),"
        'strSQL = strSQL & "Weight2Sigma_Neg '-2Sigma',"
        ''strSQL = strSQL & "B.Sigma*(-2),"
        'strSQL = strSQL & "Weight1Sigma_Neg '-1Sigma',"
        ''strSQL = strSQL & "B.Sigma*(-1),"
        'strSQL = strSQL & "Weight1Sigma_Pos '+1Sigma',"
        ''strSQL = strSQL & "B.Sigma,"
        'strSQL = strSQL & "Weight2Sigma_pos '+2Sigma',"
        ''strSQL = strSQL & "B.Sigma*2,"
        'strSQL = strSQL & "Weight3Sigma_Pos '+3Sigma',"
        ''strSQL = strSQL & "B.Sigma*3,"
        'strSQL = strSQL & "MinTesterByLot 'MinTester',"
        'strSQL = strSQL & "MinHgaByTester 'MinHGA',"
        'strSQL = strSQL & "MinHgaByLotT2 'MinHGAT2',"
        'strSQL = strSQL & "MinHgaByLotT3 'MinHGAT3',"
        'strSQL = strSQL & "MinHgaByLotT4 'MinHGAT4',"
        'strSQL = strSQL & "MinHgaByLotT4AndMore 'MoreThanT4',"
        'strSQL = strSQL & "PatternRunPoint,"
        'strSQL = strSQL & "OutlierPoint,"
        'strSQL = strSQL & "LockCode,"
        'strSQL = strSQL & "OutlierOption,"
        'strSQL = strSQL & "EmailID "
        'strSQL = strSQL & "FROM db_" & strProduct & ".tabparameterbyproduct A "
        'strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabautox2lotsetting B USING(param_rttc) "
        'strSQL = strSQL & "WHERE ParamMachine <>'' "
        'strSQL = strSQL & "AND (param_add=True OR param_mul=True) "
        ''strSQL = strSQL & "AND NOT Sigma IS NULL "
        'strSQL = strSQL & "ORDER BY AdjustX2Lot DESC,A.paramID ASC;"

        Dim strSQL As String = "SELECT * FROM db_" & strProduct & ".tabautox2lotsetting A "
        strSQL = strSQL & "ORDER BY AdjustX2Lot DESC,A.param_rttc ASC;"
        Dim clsSQL As New CMySQL
        GetAutoX2LotByProduct = clsSQL.CommandMySqlDataTable(strSQL, mySqlConn)
    End Function

    Public Function GetAutoX2FasttrackSettingByProduct(ByVal strProduct As String, ByVal mySqlConn As MySqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "SELECT A.param_rttc EDB,"
        strSQL = strSQL & "A.ParamMDB,"
        strSQL = strSQL & "B.IsEnable AdjustX2FT,"
        strSQL = strSQL & "B.DoBinning,"
        strSQL = strSQL & "B.CorrelateParam,"
        strSQL = strSQL & "B.CorrelateValue,"
        strSQL = strSQL & "B.CorrelateDelta,"
        strSQL = strSQL & "B.TargetValue,"
        strSQL = strSQL & "B.TargetDelta,"
        strSQL = strSQL & "B.BinningValue,"
        strSQL = strSQL & "B.Sigma,"
        strSQL = strSQL & "B.AdjustOption,"
        strSQL = strSQL & "B.AdjustType,"
        strSQL = strSQL & "A.LCL,"
        strSQL = strSQL & "A.UCL,"
        strSQL = strSQL & "Weight3Sigma_Neg '-3Sigma',"
        strSQL = strSQL & "Weight2Sigma_Neg '-2Sigma',"
        strSQL = strSQL & "Weight1Sigma_Neg '-1Sigma',"
        strSQL = strSQL & "Weight1Sigma_Pos '+1Sigma',"
        strSQL = strSQL & "Weight2Sigma_pos '+2Sigma',"
        strSQL = strSQL & "Weight3Sigma_Pos '+3Sigma',"
        strSQL = strSQL & "MinTesterByLot 'MinTester',"
        strSQL = strSQL & "MinHgaByTester 'MinHGA',"
        strSQL = strSQL & "MinHgaByLotT2 'MinHGAT2',"
        strSQL = strSQL & "MinHgaByLotT3 'MinHGAT3',"
        strSQL = strSQL & "MinHgaByLotT4 'MinHGAT4',"
        strSQL = strSQL & "MinHgaByLotT4AndMore 'MoreThanT4',"
        strSQL = strSQL & "PatternRunPoint,"
        strSQL = strSQL & "EmailID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabparameterbyproduct A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabautox2fasttracksetting B USING(param_rttc) "
        strSQL = strSQL & "WHERE ParamMachine <>'' "
        strSQL = strSQL & "AND (param_add=True OR param_mul=True) "
        'strSQL = strSQL & "AND NOT Sigma IS NULL "
        strSQL = strSQL & "ORDER BY B.IsEnable DESC,A.paramID ASC;"
        Dim clsSQL As New CMySQL
        GetAutoX2FasttrackSettingByProduct = clsSQL.CommandMySqlDataTable(strSQL, mySqlConn)
    End Function

    Public Function GetSkipLotByProduct(ByVal strProduct As String, ByVal MySqlConn As MySqlConnection) As DataTable
        Dim strSQL As String = "SELECT LotSkip,True 'SkipType' FROM db_" & strProduct & ".tabctr_checkreadingskiplot "
        strSQL = strSQL & "UNION SELECT lotSkip,False 'SkipType' FROM db_" & strProduct & ".tabctr_failureskiplot;"
        Dim clsMySql As New CMySQL
        GetSkipLotByProduct = clsMySql.CommandMySqlDataTable(strSQL, MySqlConn)

    End Function


End Class
