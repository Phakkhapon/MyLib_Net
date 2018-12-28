
Imports System.IO
Imports MySql.Data.MySqlClient

Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Drawing

Public Class CSendEMail
    Private m_strMailDetail() As String = Nothing

    Public Sub New()

    End Sub

    Public Sub ComposeEMail(ByVal strToAddress As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strContent As String)

        Dim strCmm As String = """"
        Dim strNewLine As String = Environment.NewLine

        ReDim m_strMailDetail(10)
        m_strMailDetail(0) = "<!--To{" & strToAddress & "}-->"
        m_strMailDetail(1) = "<!--From{" & strFrom & "}-->"
        m_strMailDetail(2) = "<!--Subject{" & strSubject & "}-->"
        m_strMailDetail(3) = "<!--Body-->"
        m_strMailDetail(4) = "<html>"
        m_strMailDetail(5) = "<head>"
        m_strMailDetail(6) = "<title>automatically</title>"
        m_strMailDetail(7) = "<meta http-equiv=" & strCmm & " Content-Type" & strCmm & " content=" & strCmm & " text/html; charset=windows-874" & strCmm & " >"
        m_strMailDetail(8) = "</head>"
        m_strMailDetail(9) = "<body>"
        m_strMailDetail(10) = "<FONT SIZE=" & strCmm & "2px" & strCmm & " Color=" & strCmm & "#000066" & strCmm & ">"
        Dim strSplitContent() As String = Split(strContent, ";")
        ReDim Preserve m_strMailDetail(m_strMailDetail.Length + strSplitContent.Length + 9)
        For nContent As Integer = 0 To strSplitContent.Length - 1
            m_strMailDetail(11 + nContent) = strSplitContent(nContent) & "<br>"
        Next nContent
        m_strMailDetail(strSplitContent.Length + 11) = "</p>"
        m_strMailDetail(strSplitContent.Length + 12) = "<Font Color=" & strCmm & "009933" & strCmm & "> Remark</Font>: <br>"
        m_strMailDetail(strSplitContent.Length + 13) = "  <br> </p>"
        m_strMailDetail(strSplitContent.Length + 14) = " Best Regards <br>"
        m_strMailDetail(strSplitContent.Length + 15) = " Test-Eng.,Ext 76050<br>"
        m_strMailDetail(strSplitContent.Length + 16) = " This is an automatically  email. <Font Color=" & strCmm & "#FF0000" & strCmm & ">Please <strong>do not </strong> reply</Font>.<br>"
        m_strMailDetail(strSplitContent.Length + 17) = " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        m_strMailDetail(strSplitContent.Length + 18) = " </p></FONT></body>"
        m_strMailDetail(strSplitContent.Length + 19) = "</html>"

    End Sub

    Public Sub SendEmailThroughSMTP(ByVal strUser As String, ByVal strPasswd As String, ByVal strToAddress As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strContent As String, Optional ByVal bmImageAttached() As Bitmap = Nothing, Optional ByVal dtsTableAttached As DataSet = Nothing)
        Try
            Dim strCmm As String = """"
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(strUser, strPasswd)
            Smtp_Server.Port = 25
            Smtp_Server.EnableSsl = False
            Smtp_Server.Host = "wdtbtsd08" ' "172.16.51.19"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(strFrom)
            strToAddress = Replace(strToAddress, ";", ",")
            e_mail.To.Add(strToAddress)
            e_mail.Subject = strSubject
            e_mail.IsBodyHtml = True
            e_mail.Priority = MailPriority.High

            Dim strBody As String = ""
            strBody = strBody & "<html>"
            strBody = strBody & "<head>"
            strBody = strBody & "<title>automatically</title>"
            strBody = strBody & "<meta http-equiv=" & strCmm & " Content-Type" & strCmm & " content=" & strCmm & " text/html; charset=windows-874" & strCmm & " >"
            strBody = strBody & "</head>"
            strBody = strBody & "<body>"
            strBody = strBody & "<FONT SIZE=" & strCmm & "2px" & strCmm & " Color=" & strCmm & "#000066" & strCmm & ">"
            Dim strSplitContent() As String = Split(strContent, ";")
            For nContent As Integer = 0 To strSplitContent.Length - 1
                strBody = strBody & strSplitContent(nContent) & "<br>"
            Next nContent
            strBody = strBody & "</p>"
            strBody = strBody & "<Font Color=" & strCmm & "009933" & strCmm & "> Remark</Font>: <br>"
            If Not bmImageAttached Is Nothing Then
                For nBitmap As Integer = 0 To bmImageAttached.Length - 1
                    strBody = strBody & "<td><img src=Chart" & nBitmap & ".Jpg></td>"
                    Dim imgStream As New MemoryStream()
                    Dim filename As String = "Chart" & nBitmap & ".Jpg"
                    bmImageAttached(nBitmap).Save(imgStream, System.Drawing.Imaging.ImageFormat.Jpeg)
                    imgStream.Position = 0
                    e_mail.Attachments.Add(New Attachment(imgStream, filename, System.Net.Mime.MediaTypeNames.Image.Jpeg))
                Next nBitmap
            End If
            strBody = strBody & "<br> </p>"
            If Not dtsTableAttached Is Nothing Then
                For nTable As Integer = 0 To dtsTableAttached.Tables.Count - 1
                    Dim dtbTableAttached As DataTable = dtsTableAttached.Tables(nTable)
                    If Not dtbTableAttached Is Nothing Then
                        strBody = strBody & dtbTableAttached.TableName & "<br>"
                        strBody = strBody & "<table width=" & strCmm & "550" & strCmm & " border=" & strCmm & "0" & strCmm & ">"
                        strBody = strBody & "<tr bgcolor=" & strCmm & "#99FFFF" & strCmm & ">"
                        For nCol As Integer = 0 To dtbTableAttached.Columns.Count - 1
                            Dim strColName As String = dtbTableAttached.Columns(nCol).ColumnName
                            strBody = strBody & "<td width=" & strCmm & "100" & strCmm & " scope=" & strCmm & "col" & strCmm & "><div align=" & strCmm & "center" & strCmm & "><Font color=" & strCmm & "#0000FF" & strCmm & " size=" & strCmm & "1.5" & strCmm & ">" & strColName & "</Font></div></td>"
                        Next nCol
                        strBody = strBody & "</tr>"
                        For nRow As Integer = 0 To dtbTableAttached.Rows.Count - 1
                            strBody = strBody & "<tr bgcolor=" & strCmm & "#FFFFF0" & strCmm & ">"
                            For nCol As Integer = 0 To dtbTableAttached.Columns.Count - 1
                                Dim strValue As String = dtbTableAttached.Rows(nRow).Item(nCol).ToString
                                If strValue <> "" And dtbTableAttached.Columns(nCol).DataType Is Type.GetType("System.Double") Then
                                    strValue = Format(dtbTableAttached.Rows(nRow).Item(nCol), "0.0000")
                                End If
                                strBody = strBody & "<td><div align=" & strCmm & "center" & strCmm & "><Font size=" & strCmm & "1" & strCmm & ">" & strValue & "</Font></div></td>"
                            Next nCol
                            strBody = strBody & "</tr>"
                        Next nRow
                        strBody = strBody & "</table>"
                    End If
                Next nTable
            End If
            strBody = strBody & "<br> </p>"
            strBody = strBody & " Best Regards <br>"
            strBody = strBody & " Test-Eng.,Ext 76050<br>"
            strBody = strBody & " This is an automatically  email. <Font Color=" & strCmm & "#FF0000" & strCmm & ">Please <strong>do not </strong> reply</Font>.<br>"
            strBody = strBody & " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
            strBody = strBody & " </p></FONT></body>"
            strBody = strBody & "</html>"
            e_mail.Body = strBody
            Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network
            Smtp_Server.Send(e_mail)
        Catch error_t As Exception
            'MsgBox(error_t.ToString)
        End Try
    End Sub

    Public Sub SendEmail(ByVal strEmailPath As String)
        If Not m_strMailDetail Is Nothing Then
            File.WriteAllLines(strEmailPath, m_strMailDetail)
        End If
    End Sub

    Public Sub Destroy()
        Finalize()
    End Sub

    Protected Overrides Sub Finalize()
        GC.Collect()
        MyBase.Finalize()
    End Sub

End Class
