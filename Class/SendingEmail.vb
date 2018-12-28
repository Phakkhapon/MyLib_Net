
Imports System.IO
Imports MySql.Data.MySqlClient
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Drawing

Public Class SendingEmail
    Private m_strMailDetail() As String = Nothing

    Private _retErr As String
    Public ReadOnly Property RetError() As String
        Get
            Return _retErr
        End Get
    End Property

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

    Public Function SendEmailThroughSMTP(ByVal strUser As String, ByVal strPasswd As String, ByVal strToAddress As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strContent As String, Optional ByVal bmImageAttached As Bitmap = Nothing, Optional ByVal dtbTableAttached As DataTable = Nothing, Optional ByVal NeedDetail As Boolean = False) As Integer

        Dim ret As Integer = 0
        Try
            Dim strCmm As String = """"
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(strUser, strPasswd)
            Smtp_Server.Port = 25
            Smtp_Server.EnableSsl = False
            Smtp_Server.Host = "mailrelay.wdc.com" ' "10.81.120.69"  ' "wdtbtsd08" ' "172.16.51.19"

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

            If NeedDetail = True Then
                Dim strSplitContent() As String = Split(strContent, ";")
                For nContent As Integer = 0 To strSplitContent.Length - 1
                    strBody = strBody & strSplitContent(nContent) & "<br>"
                Next nContent
                strBody = strBody & "</p>"
                strBody = strBody & "<Font Color=" & strCmm & "009933" & strCmm & "> Remark</Font>: <br>"
                'If Not dtbTableAttached Is Nothing Then
                '    strBody = strBody & "<table width=" & strCmm & "330" & strCmm & " border=" & strCmm & "0" & strCmm & ">"
                '    strBody = strBody & "<tr bgcolor=" & strCmm & "#99FFFF" & strCmm & ">"
                '    For nCol As Integer = 0 To dtbTableAttached.Columns.Count - 1
                '        Dim strColName As String = dtbTableAttached.Columns(nCol).ColumnName
                '        strBody = strBody & "<td width=" & strCmm & "80" & strCmm & " scope=" & strCmm & "col" & strCmm & "><div align=" & strCmm & "center" & strCmm & "><Font color=" & strCmm & "#0000FF" & strCmm & " size=" & strCmm & "2px" & strCmm & ">" & strColName & "</Font></div></td>"
                '    Next nCol
                '    strBody = strBody & "</tr>"
                '    For nRow As Integer = 0 To dtbTableAttached.Rows.Count - 1
                '        strBody = strBody & "<tr bgcolor=" & strCmm & "#FFFFF0" & strCmm & ">"
                '        For nCol As Integer = 0 To dtbTableAttached.Columns.Count - 1
                '            Dim strValue As String = dtbTableAttached.Rows(nRow).Item(nCol).ToString
                '            If strValue <> "" And dtbTableAttached.Columns(nCol).DataType Is Type.GetType("System.Double") Then
                '                strValue = Format(dtbTableAttached.Rows(nRow).Item(nCol), "0.0000")
                '            End If
                '            strBody = strBody & "<td><div align=" & strCmm & "center" & strCmm & "><Font size=" & strCmm & "2px" & strCmm & ">" & strValue & "</Font></div></td>"
                '        Next nCol
                '        strBody = strBody & "</tr>"
                '    Next nRow
                '    strBody = strBody & "</table>"
                'End If

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

                strBody = strBody & "<br> </p>"
                ' strBody = strBody & " Best Regards <br>"
                ' strBody = strBody & " Test-Eng.,Ext 76050<br>"
                strBody = strBody & " This is an automatically  email. <Font Color=" & strCmm & "#FF0000" & strCmm & ">Please <strong>do not </strong> reply</Font>.<br>"
                strBody = strBody & " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
                strBody = strBody & " </p></FONT></body>"
                strBody = strBody & "</html>"
                If Not bmImageAttached Is Nothing Then
                    strBody = strBody & "<td><img src=Chart.Jpg></td>"
                    Dim imgStream As New MemoryStream()
                    Dim filename As String = "Chart.Jpg"
                    bmImageAttached.Save(imgStream, System.Drawing.Imaging.ImageFormat.Jpeg)
                    imgStream.Position = 0
                    e_mail.Attachments.Add(New Attachment(imgStream, filename, System.Net.Mime.MediaTypeNames.Image.Jpeg))
                End If
            End If
            e_mail.Body = strBody
            Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network

            Smtp_Server.Send(e_mail)

        Catch error_t As Exception
            ret = -1
            _retErr = error_t.Message
        End Try
        Return ret

    End Function

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


    Public Function SendMailAlert(ByVal MailAddress As String, ByVal ReceiveMail As String, ByVal subjectMail As String, ByVal DetailsMail As String) As Boolean
        Dim statusMail As Boolean = False

        Dim myMail As New MailMessage()
        ' myMail.From = New MailAddress("webmaster@thaicreate.com")
        myMail.From = New MailAddress(MailAddress)
        myMail.To.Add(New MailAddress(ReceiveMail))
        myMail.Subject = subjectMail
        myMail.Body = DetailsMail
        myMail.Priority = MailPriority.High ' MailPriority.High , MailPriority.Low  MailPriority.Normal
        Dim Client As New SmtpClient()
        Client.Send(myMail)

        Return statusMail
    End Function


End Class

