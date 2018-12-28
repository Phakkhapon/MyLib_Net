Public Class CDataEncode

    Public Function base64Encode(ByVal data As String) As String
        Try
            Dim encData_byte As Byte() = New Byte(data.Length - 1) {}
            encData_byte = System.Text.Encoding.UTF8.GetBytes(data)
            Dim encodedData As String = Convert.ToBase64String(encData_byte)
            Return encodedData
        Catch ex As Exception
            Throw New Exception("Error in base64Encode" + ex.Message)
        End Try
    End Function

    Public Function base64Decode(ByVal data As String) As String
        Try
            Dim encoder As New System.Text.UTF8Encoding()
            Dim utf8Decode As System.Text.Decoder = encoder.GetDecoder()
            Dim todecode_byte As Byte() = Convert.FromBase64String(data)
            Dim charCount As Integer = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length)
            Dim decoded_char As Char() = New Char(charCount - 1) {}
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0)
            Dim result As String = New [String](decoded_char)
            Return result
        Catch ex As Exception
            Throw New Exception("Error in base64Decode" + ex.Message)
        End Try
    End Function

End Class

