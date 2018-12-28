
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Net.NetworkInformation
Imports System.Text

    ''' <summary>
''' TCP sevices. 
''' sends and receives Encoding.Binary data
    ''' the max length of one data frame is 8196 bytes
    ''' last update: 8.1.2011
    ''' 10.5.2011 - remote/local  IP Endpoint
    ''' update 24.6.2011 DNS 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CTCPDataControl
        ''' <summary>
        ''' Event data send back to calling form
        ''' </summary>
    Public Event Datareceived(ByVal txt As String)
        ''' <summary>
        ''' connection status back to form True: ok
        ''' </summary>
    Public Event Connection(ByVal cStatus As Boolean)
        ''' <summary>
        ''' data send successfull (True)
        ''' </summary>
    Public Event sendOK(ByVal sStatus As Boolean)
        ''' <summary>
        ''' data receive successfull (True)
        ''' </summary>
    Public Event recOK(ByVal sReceive As Boolean)

    ' Private serverRuns As Boolean
    Private m_ServerListener As TcpListener
    Private clientSocket As TcpClient
    Private sc As SynchronizationContext
    Private isConnected, receiveStatus, sendStatus As Boolean
    Private iRemote, pLocal As EndPoint

    ''' <summary>
    ''' reads endpoints
    ''' </summary>
    Public ReadOnly Property Remote() As EndPoint
        Get
            Return iRemote
        End Get
    End Property

    ''' <summary>
    ''' reads local point
    ''' </summary>
    Public ReadOnly Property Local() As EndPoint
        Get
            Return pLocal
        End Get
    End Property

    ''' <summary>
    ''' TCP connect with server
    ''' </summary>
    Public Sub Connect(ByVal hostAdress As String, ByVal hostPort As Integer)

        sc = SynchronizationContext.Current

        Try
            m_ServerListener = New TcpListener(IPAddress.Parse(hostAdress), hostPort)
        Catch ex As Exception
            MsgBox("server create: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try

        Try
            With m_ServerListener
                .Start()
                Dim res As IAsyncResult = .BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), m_ServerListener)
                ' res.AsyncWaitHandle.WaitOne(2000,False)
                isConnected = True
            End With
        Catch ex As Exception
            MsgBox("server listen: " & ex.Message, MsgBoxStyle.Exclamation)
            isConnected = False
        Finally
            RaiseEvent Connection(isConnected)
        End Try

    End Sub

    ''' <summary>
    ''' disConnect server
    ''' </summary>
    Public Sub Disconnect()
        Try
            isConnected = False
            If Not m_ServerListener Is Nothing Then m_ServerListener.Stop()
        Catch ex As Exception
            MsgBox("disConnect server: " & ex.Message, MsgBoxStyle.Exclamation)
            isConnected = True
        Finally
            RaiseEvent Connection(isConnected)
        End Try
    End Sub
    ''' <summary>
    ''' TCP send data
    ''' </summary>
    Public Function SendData(ByVal txt As String, ByVal remoteAddress As String, ByVal remotePort As Integer, Optional ByVal intTimeOut As Integer = 2000) As String
        SendData = ""
        Try
            Dim clientSocket = New TcpClient
            'Dim iP As IPAddress = IPAddress.Any
            'Dim isIp As Boolean = IPAddress.TryParse(remoteAddress, iP)
            sendStatus = False
            With clientSocket

                'Set connection time out
                'Dim s As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                Dim result = clientSocket.BeginConnect(IPAddress.Parse(remoteAddress), remotePort, New AsyncCallback(AddressOf DoBeginConnect), clientSocket)
                result.AsyncWaitHandle.WaitOne(intTimeOut, False)
                'If isIp Then    ' ip address
                '    .Connect(IPAddress.Parse(remoteAddress), remotePort)
                'Else            ' DNS name
                '    .Connect(remoteAddress, remotePort)
                'End If

                If clientSocket.Connected Then
                    clientSocket.SendTimeout = intTimeOut
                    clientSocket.ReceiveTimeout = intTimeOut
                    Dim data() As Byte = Encoding.Unicode.GetBytes(txt)
                    .NoDelay = True
                    .GetStream().Write(data, 0, data.Length)

                    Dim dataRespose = New [Byte](.ReceiveBufferSize) {}

                    ' String to store the response ASCII representation. 
                    Dim responseData As [String] = [String].Empty
                    ' Read the first batch of the TcpServer response bytes. 
                    Dim bytes As Int32 = .GetStream().Read(dataRespose, 0, dataRespose.Length)
                    responseData = System.Text.Encoding.Unicode.GetString(dataRespose, 0, bytes)
                    .GetStream().Close()
                    .Close()
                    sendStatus = True
                    SendData = responseData
                Else
                    .Close()
                End If

            End With
        Catch ex As Exception
            'Throw New System.Exception(ex.Message)
            'MsgBox("sendData: " & ex.Message, MsgBoxStyle.Exclamation)
            SendData = ""
            sendStatus = False
        Finally
            RaiseEvent sendOK(sendStatus)
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' last update 10.5.2011
    ''' </summary>
    Private Sub DoBeginConnect()

    End Sub
    ''' <summary>
    ''' TCP asynchronous receive on secondary thread
    ''' last update 10.5.2011
    ''' </summary>
    Private Sub DoAccept(ByVal ar As IAsyncResult)

        Dim sb As New StringBuilder
        Dim buf() As Byte
        Dim datalen As Integer
        Dim listener As TcpListener
        Dim clientSocket As TcpClient
        If Not isConnected Then Exit Sub
        Try
            listener = CType(ar.AsyncState, TcpListener)
            clientSocket = listener.EndAcceptTcpClient(ar)
            clientSocket.ReceiveTimeout = 5
            'update 10.5.2011
            iRemote = clientSocket.Client.RemoteEndPoint
            pLocal = clientSocket.Client.LocalEndPoint
        Catch ex As ObjectDisposedException
            MsgBox("DoAccept ObjectDisposedException " & ex.Message, MsgBoxStyle.Exclamation)
            ' after server.stop() AsyncCallback is also active, but the object server is disposed
            Exit Sub
        End Try

        Try
            With clientSocket
                datalen = 0
                ' somtimes it occurs that .available returns the value 0 also data in buffer exists
                While datalen = 0
                    ' data in read Buffer
                    datalen = .Available
                End While

                buf = New Byte(datalen - 1) {}
                'get entire bytes at once
                .GetStream().Read(buf, 0, buf.Length)
                .GetStream().Write(buf, 0, buf.Length)         'Response data back
                sb.Append(Encoding.Unicode.GetString(buf, 0, buf.Length))
                'sb.Append(Encoding.ASCII.GetString(buf, 0, buf.Length)) 'waiwai edit
                .Close()
            End With
            receiveStatus = True
        Catch ex As TimeoutException
            MsgBox("doAcceptData timeout: " & ex.Message, MsgBoxStyle.Exclamation)
            receiveStatus = False
            clientSocket.Close()
            Exit Sub
        Catch ex As Exception
            MsgBox("doAcceptData: " & ex.Message, MsgBoxStyle.Exclamation)
            receiveStatus = False
            clientSocket.Close()
            Exit Sub
        Finally
            RaiseEvent recOK(receiveStatus)
        End Try
        ' post event
        sc.Post(New SendOrPostCallback(AddressOf OnDatareceived), sb.ToString)
        ' start new read
        m_ServerListener.BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), m_ServerListener)
    End Sub
    '
    ' now data to calling class and UI thread
    '
    Private Sub OnDatareceived(ByVal state As Object)
        RaiseEvent Datareceived(state.ToString)
    End Sub
End Class
