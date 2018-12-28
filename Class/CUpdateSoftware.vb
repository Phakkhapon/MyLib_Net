
Imports MySql.Data.MySqlClient
Imports System.IO
Public Class CUpdateSoftware

    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

    Public Function GetFileVersion(ByVal strFullFileName As String) As String
        If File.Exists(strFullFileName) Then
            Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(strFullFileName)
            GetFileVersion = myFileVersionInfo.FileVersion
        Else
            GetFileVersion = ""
        End If
    End Function

    Public Function GetRTTCFileVersionFromServer(ByVal strFileName As String) As String
        Dim strSQL As String = "SELECT ProgramVersion FROM rttc_program.rttc_program_control "
        strSQL = strSQL & "WHERE ProgramName='" & strFileName & "';"
        Dim clsMySql As New CMySQL
        Dim dtbFileVersion As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        If dtbFileVersion.Rows.Count > 0 Then
            GetRTTCFileVersionFromServer = dtbFileVersion.Rows(0).Item("ProgramVersion")
        Else
            GetRTTCFileVersionFromServer = ""
        End If
    End Function

    Public Sub UploadRTTCFileToServer(ByVal sFileInfo As SRTTCSoftware)
        Dim byteFile() As Byte = File.ReadAllBytes(sFileInfo.strFullFileName)

        Dim strParameterValue As String = "?SoftwareData"

        Dim MyParam As New MySqlParameter(strParameterValue, MySqlDbType.Blob, byteFile.Length, _
                        ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, byteFile)
        Dim obj As System.Array = MyParam.Value


        Dim strSQL As String = "INSERT INTO rttc_program.rttc_program_control(ProgramName,ProgramVersion,SoftwareData ) VALUES ('" & _
                                  sFileInfo.strFileName & "', '" & sFileInfo.strVersion & "'," & strParameterValue & " ) "
        strSQL = strSQL & "ON DUPLICATE KEY UPDATE "
        strSQL = strSQL & "ProgramVersion='" & sFileInfo.strVersion & "',"
        strSQL = strSQL & "SoftwareData=" & strParameterValue & ";"

        Dim myComm As New MySqlCommand(strSQL, m_mySqlConn)
        myComm.Parameters.Add(MyParam)
        m_mySqlConn.Open()
        myComm.ExecuteNonQuery()
        m_mySqlConn.Close()
        MySqlConnection.ClearPool(m_mySqlConn)
    End Sub

    Public Function DownloadRTTCFileFromServer(ByVal strFileName As String) As Byte()
        Dim strSQL As String = "SELECT SoftwareData FROM rttc_program.rttc_program_control WHERE ProgramName='" & strFileName & "'"
        Dim myComm As New MySqlCommand(strSQL, m_mySqlConn)
        'Data Adapter to fill the Dataset
        Dim myData As New MySqlDataAdapter(myComm)
        Dim myDS As New DataTable
        'Populate dataset
        myData.Fill(myDS)

        DownloadRTTCFileFromServer = myDS.Rows(0).Item(0)
        m_mySqlConn.Close()
        MySqlConnection.ClearPool(m_mySqlConn)
    End Function

    Public Function GetCFFileVersionFromServer(ByVal strFileName As String) As String
        Dim strSQL As String = "SELECT ProgramVersion FROM rttc_program.cf_program_control "
        strSQL = strSQL & "WHERE ProgramName='" & strFileName & "';"
        Dim clsMySql As New CMySQL
        Dim dtbFileVersion As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        If dtbFileVersion.Rows.Count > 0 Then
            GetCFFileVersionFromServer = dtbFileVersion.Rows(0).Item("ProgramVersion")
        Else
            GetCFFileVersionFromServer = ""
        End If
    End Function

    Public Sub UploadCFFileToServer(ByVal sFileInfo As SRTTCSoftware)
        Dim byteFile() As Byte = File.ReadAllBytes(sFileInfo.strFullFileName)

        Dim strParameterValue As String = "?SoftwareData"

        Dim MyParam As New MySqlParameter(strParameterValue, MySqlDbType.Blob, byteFile.Length, _
                        ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, byteFile)
        Dim obj As System.Array = MyParam.Value

        Dim strSQL As String = "INSERT INTO rttc_program.cf_program_control(ProgramName,ProgramVersion,SoftwareData ) VALUES ('" & _
                                  sFileInfo.strFileName & "', '" & sFileInfo.strVersion & "'," & strParameterValue & " ) "
        strSQL = strSQL & "ON DUPLICATE KEY UPDATE "
        strSQL = strSQL & "ProgramVersion='" & sFileInfo.strVersion & "',"
        strSQL = strSQL & "SoftwareData=" & strParameterValue & ";"

        Dim myComm As New MySqlCommand(strSQL, m_mySqlConn)
        myComm.Parameters.Add(MyParam)
        m_mySqlConn.Open()
        myComm.ExecuteNonQuery()
        m_mySqlConn.Close()
        MySqlConnection.ClearPool(m_mySqlConn)
    End Sub

    Public Function DownloadCFFileFromServer(ByVal strFileName As String) As Byte()
        Dim strSQL As String = "SELECT FileData FROM rttc_program.cf_program_control WHERE ProgramName='" & strFileName & "'"
        Dim myComm As New MySqlCommand(strSQL, m_mySqlConn)
        'Data Adapter to fill the Dataset
        Dim myData As New MySqlDataAdapter(myComm)
        Dim myDS As New DataTable
        'Populate dataset
        myData.Fill(myDS)

        DownloadCFFileFromServer = myDS.Rows(0).Item(0)
        m_mySqlConn.Close()
        MySqlConnection.ClearPool(m_mySqlConn)
    End Function

End Class
