Imports MySql.Data.MySqlClient

Public Class CDatabaseManage

    Private Enum enuTesterType
        eNone = 0
        eEH = 1
        eDCT = 2
        eV2002 = 3
    End Enum

    Private m_mySqlWexConn As MySqlConnection

    Public Sub New(ByVal mySqlWexConn As MySqlConnection)
        m_mySqlWexConn = mySqlWexConn
    End Sub

    Public Function GetDatabaseList() As DataTable
        Dim strSQL As String = "SHOW DATABASES;"
        Dim CMyProduct As New CMySQL
        GetDatabaseList = CMyProduct.CommandMySqlDataTable(strSQL, m_mySqlWexConn)
    End Function

    Public Function DropDatabase(ByVal strDatabaseName As String) As Boolean
        On Error Resume Next
        Dim strSQL As String
        DropDatabase = False
        Dim clsDropDatabase As New CMySQL
        strSQL = "DROP DATABASE db_" & strDatabaseName & ";"
        clsDropDatabase.CommandNonQuery(strSQL, m_mySqlWexConn, strDatabaseName, Me.ToString)

        'strSQL = "DELETE FROM rttc_parameter.tabparameter WHERE ProductName='" & strDatabaseName & "';"
        'cDropDatabas.CommandMySqlDataTable(strSQL, m_mySqlRTTCCon)

    End Function

    Public Function AddNewProduct(ByVal strProductName As String) As Boolean
        AddNewProduct = False
        Try
            CreateNewRawDatabase(strProductName)
            AddNewProduct = True
        Catch ex As Exception
            AddNewProduct = False
            If m_mySqlWexConn.State = ConnectionState.Open Then m_mySqlWexConn.Close()
            MySqlConnection.ClearPool(m_mySqlWexConn)
        End Try
    End Function

    Private Function CreateNewRawDatabase(ByVal strProduct As String) As Boolean 'Create database by copy from db_master
        CreateNewRawDatabase = False
        If strProduct = "" Then Return False
        Dim strSQL As String
        Dim clsSql As New CMySQL

        Try
            strSQL = "CREATE DATABASE db_" & strProduct
            clsSql.CommandNonQuery(strSQL, m_mySqlWexConn, strProduct, Me.ToString) 'Create new database
        Catch ex As Exception

        End Try

        'm_myConn.ChangeDatabase("db_" & strProduct)  'Point  to new database created

        Dim dtbTable As DataTable = ShowTableInDatabase("db_master")    'Show all table inside db_master
        Dim nTableNum As Integer = dtbTable.Rows.Count
        For nTable As Integer = 0 To nTableNum - 1
            strSQL = "CREATE TABLE db_" & strProduct & "." & dtbTable.Rows(nTable).Item(0) & " Like db_master." & dtbTable.Rows(nTable).Item(0) & ";"
            clsSql.CommandNonQuery(strSQL, m_mySqlWexConn, strProduct, Me.ToString)  'Show create table string
        Next nTable
        'InsertProductList(strProduct.ToUpper)
        InitParameterByProduct(strProduct.ToUpper)
        InserReject2lotTable(strProduct.ToUpper)  'Insert Reject2lot table to new product 

        CreateNewRawDatabase = True
    End Function

    'Private Sub InsertProductList(ByVal strProduct As String)
    '    Try
    '        Dim strSQL As String
    '        Dim clsSql As New CMySQL
    '        strSQL = "INSERT INTO ctr_controlsetting.tabproductdetail(productId,Productname) "
    '        strSQL = strSQL & "SELECT (SELECT IFNULL(MAX(productid)+1,1) FROM ctr_controlsetting.tabproductdetail),"
    '        strSQL = strSQL & "'" & strProduct & "';"
    '        clsSql.CommandMySqlDataTable(strSQL, m_mySqlWexConn)
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Private Sub InitParameterByProduct(ByVal strProduct As String)

        Dim nTesterType As Integer
        If InStr(1, strProduct.ToString.ToUpper, "DCT") > 0 Then
            nTesterType = enuTesterType.eDCT
        ElseIf InStr(1, strProduct.ToString.ToUpper, "EH300") > 0 Then
            nTesterType = enuTesterType.eEH
        ElseIf InStr(1, strProduct.ToString.ToUpper, "V2002") > 0 Then
            nTesterType = enuTesterType.eV2002
        Else
            nTesterType = enuTesterType.eEH
        End If
        Dim clsParameter As New CParameterRTTCMapping(m_mySqlWexConn)
        Dim dtbParam As DataTable = clsParameter.GetParameterMapping
        'dtbParam = clsParameter.GetParameterMapping
        Dim strSQL As String
        Dim clsSQL As New CMySQL
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            'Try
            Dim strParamRTTC As String = dtbParam.Rows(nParam).Item("param_rttc")
            strSQL = "INSERT INTO db_" & strProduct & ".tabparameterbyproduct "
            strSQL = strSQL & "(paramID,paraOrder,param_rttc,parammachine,paramMDB,ParamFasttrack,param_add,param_mul,Zone,Setup,IsEnable,LCL,UCL) VALUES("
            strSQL = strSQL & dtbParam.Rows(nParam).Item("paramID") & ","
            strSQL = strSQL & dtbParam.Rows(nParam).Item("paramID") & ","
            strSQL = strSQL & "'" & strParamRTTC & "',"
            Select Case nTesterType
                Case enuTesterType.eDCT
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_dct") & "',"
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_dct") & "',"
                Case enuTesterType.eEH
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_eh") & "',"
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_dct") & "',"
                Case enuTesterType.eV2002
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_v2002") & "',"
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_dct") & "',"
                Case enuTesterType.eNone
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_rttc") & "',"
                    strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_dct") & "',"
            End Select
            strSQL = strSQL & "'" & dtbParam.Rows(nParam).Item("param_rttc") & "',"
            strSQL = strSQL & dtbParam.Rows(nParam).Item("param_add") & ","
            strSQL = strSQL & dtbParam.Rows(nParam).Item("param_mul") & ","
            strSQL = strSQL & "'Zone2',"
            strSQL = strSQL & "'Setup1',"
            strSQL = strSQL & dtbParam.Rows(nParam).Item("IsEnable") & ","
            If strParamRTTC.ToUpper = "CYCLETIME" Then
                strSQL = strSQL & "0,"
            Else
                strSQL = strSQL & "-9999,"
            End If
            strSQL = strSQL & "9999)"
            strSQL = strSQL & " ON DUPLICATE KEY UPDATE paramID='" & dtbParam.Rows(nParam).Item("paramID") & "';"
            clsSQL.CommandNoQuery(strSQL, m_mySqlWexConn)
            'Catch ex As Exception
            'Skip error when existing value
            'End Try
        Next nParam
    End Sub

    Private Sub InserReject2lotTable(strProduct As String)

        Dim strSQL As String = "INSERT INTO db_"
        If strProduct.Contains("DOE") = False Then
            strSQL = strSQL & strProduct & ".tabctr_failurecounter "
            strSQL = strSQL & "SELECT * FROM db_master.tabctr_failurecounter"
            Dim clsSQL As New CMySQL
            clsSQL.CommandNoQuery(strSQL, m_mySqlWexConn)
        End If

    End Sub



    Public Function ShowTableInDatabase(ByVal strDatabaseName As String) As DataTable
        Dim strSQL As String
        Dim clsSql As New CMySQL
        strSQL = "SHOW TABLES FROM " & strDatabaseName & ";"
        ShowTableInDatabase = clsSql.CommandMySqlDataTable(strSQL, m_mySqlWexConn)  'Show create table string
    End Function

End Class
