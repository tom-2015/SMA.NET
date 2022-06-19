Imports System.IO
Imports System.IO.Ports
Imports System.Threading

Public Class SMAConnection

    Public Property EnableDebugging As Boolean = False

    Protected m_Stream As Stream
    Protected m_ReadThread As Thread
    Protected m_LocalAddress As UInt16 = 1 'local network address

    Protected m_IncompletePackets As New Dictionary(Of SMATelegram.SMACommands, SMADataPacket)


    Public Event RequestNextPacket(ByVal Connection As SMAConnection, ByVal PacketCount As Byte, ByVal Command As SMATelegram.SMACommands)
    Public Event ReceivedDataPacket(ByVal Connection As SMAConnection, ByVal DataPacket As SMADataPacket)
    Public Event RS485TXOn(ByVal Connection As SMAConnection, ByVal DataTransmitted As Byte())
    Public Event RS485TXOff(ByVal Connection As SMAConnection, ByVal DataTransmitted As Byte())

    Public Event DebugRxData(ByVal Connection As SMAConnection, ByVal Data As Byte)
    Public Event DebugRxPacket(ByVal Connection As SMAConnection, ByVal Packet As PPPPacket)
    Public Event DebugTxPacket(ByVal Connection As SMAConnection, ByVal Packet As PPPPacket)

    Protected m_NetAddressAck As SMAConfigAddressPacket
    Protected m_DeviceInfoPacket As SMADeviceInfoPacket
    Protected m_UnconfiguredDevices As List(Of SMAGetNetpacket)
    Protected m_NetstartDevices As List(Of SMANetStartpacket)

    Protected m_SearchDeviceMutex As New Object
    Protected m_NetStartMutex As New Object
    Protected m_UnconfigredMutex As New Object



    Public Sub New(ByVal Port As Stream)
        m_Stream = Port
        m_ReadThread = New Thread(AddressOf ReadDataHelper)
        m_ReadThread.Start(m_Stream)
    End Sub

    Public Sub New(ByVal Port As Stream, ByVal LocalNetworkAddress As UInt16)
        m_Stream = Port ' Port.BaseStream
        m_ReadThread = New Thread(AddressOf ReadDataHelper)
        m_ReadThread.Start(m_Stream)
        m_LocalAddress = LocalNetworkAddress
    End Sub


    ''' <summary>
    ''' Runs is other thread reading all data from the port stream
    ''' </summary>
    ''' <param name="Port"></param>
    Protected Sub ReadDataHelper(ByVal Port As Stream)
        Dim CurrentPacket As New MemoryStream

        Try
            While True

                Dim Value As Integer = m_Stream.ReadByte()

                If EnableDebugging Then RaiseEvent DebugRxData(Me, Value)

                If Value >= 0 AndAlso Value <= 255 Then
                    If CurrentPacket.Length = 0 Then
                        If Value = PPPPacket.PPPStartEnd Then
                            CurrentPacket.WriteByte(Value)
                        End If
                    Else
                        CurrentPacket.WriteByte(Value)

                        If Value = PPPPacket.PPPStartEnd Then
                            NewIncomingPPP(CurrentPacket.ToArray())
                            CurrentPacket.SetLength(0)
                        End If

                    End If
                End If
            End While
        Catch e As ThreadAbortException

        End Try
    End Sub

    ''' <summary>
    ''' New PPP packet discovered in the network stream
    ''' </summary>
    ''' <param name="Data"></param>
    Protected Sub NewIncomingPPP(ByVal Data As Byte())
        Dim PPP As New PPPPacket(Data)

        If EnableDebugging Then RaiseEvent DebugRxPacket(Me, PPP)

        If PPP.IsValid() Then
            Dim Telegram As New SMATelegram(PPP.UserData())
            Dim SMADataPacket As SMADataPacket

            SyncLock m_IncompletePackets
                If m_IncompletePackets.ContainsKey(Telegram.Cmd) Then 'check if incomplete packet already exists
                    SMADataPacket = m_IncompletePackets(Telegram.Cmd)
                Else
                    SMADataPacket = SMADataPacket.CreateDatapacket(Telegram) 'no, create a new data packet
                End If
            End SyncLock

            If SMADataPacket IsNot Nothing Then

                If Telegram.PacketCount > 0 Then 'if more partial packets are waiting, send cmd to get them
                    Thread.Sleep(100)
                    RaiseEvent RequestNextPacket(Me, Telegram.PacketCount, Telegram.Cmd)
                    SendSMATelegram(New SMATelegram(Telegram.Destination, Telegram.Source, SMATelegram.SMAPacketFlags.SMAF_None, Telegram.Cmd, {}, Telegram.PacketCount))
                End If

                SMADataPacket.AddTelegram(Telegram) 'process new incoming telegram

                If SMADataPacket.isComplete Then 'packet is complete received
                    SyncLock m_IncompletePackets
                        If m_IncompletePackets.ContainsKey(Telegram.Cmd) Then m_IncompletePackets.Remove(Telegram.Cmd) 'remove from the incomplete buffer
                    End SyncLock

                    RaiseEvent ReceivedDataPacket(Me, SMADataPacket)

                    Select Case Telegram.Cmd
                        Case SMATelegram.SMACommands.CMD_CFG_NETADR
                            m_NetAddressAck = SMADataPacket
                        Case SMATelegram.SMACommands.CMD_SEARCH_DEV
                            m_DeviceInfoPacket = SMADataPacket
                        Case SMATelegram.SMACommands.CMD_GET_NET_START
                            m_NetstartDevices.Add(SMADataPacket)
                        Case SMATelegram.SMACommands.CMD_GET_NET
                            m_UnconfiguredDevices.Add(SMADataPacket)
                    End Select
                Else
                    SyncLock m_IncompletePackets
                        If Not m_IncompletePackets.ContainsKey(Telegram.Cmd) Then
                            m_IncompletePackets.Add(Telegram.Cmd, SMADataPacket)
                        End If
                    End SyncLock
                End If
            End If

        End If

    End Sub

    ''' <summary>
    ''' Gets / sets the local network address
    ''' </summary>
    ''' <returns></returns>
    Public Property LocalNetworkAddress() As UInt16
        Get
            Return m_LocalAddress
        End Get
        Set(value As UInt16)
            m_LocalAddress = value
        End Set
    End Property

    ''' <summary>
    ''' Sends a telegram over RS485
    ''' </summary>
    ''' <param name="Telegram"></param>
    Public Sub SendSMATelegram(ByVal Telegram As SMATelegram)
        Dim PPP As New PPPPacket(&HFF, &H3, &H4140, Telegram.GetData())
        Dim Buff As Byte() = PPP.GetData()

        SyncLock m_IncompletePackets 'check if this is the initial request packet
            If Telegram.PacketCount = 0 AndAlso m_IncompletePackets.ContainsKey(Telegram.Cmd) Then m_IncompletePackets.Remove(Telegram.Cmd)
        End SyncLock

        If EnableDebugging Then RaiseEvent DebugTxPacket(Me, PPP)
        RaiseEvent RS485TXOn(Me, Buff)
        m_Stream.Write(Buff, 0, Buff.Length)
        m_Stream.Flush()
        RaiseEvent RS485TXOff(Me, Buff)
    End Sub

    ''' <summary>
    ''' Sends a telegram to search for serialnumber on the network (CMD_SEARCH_DEV)
    ''' use the ReceivedDataPacket event to check for incoming device answer
    ''' </summary>
    ''' <param name="SerialNumber"></param>
    ''' <returns>the device packet with the serial and network address if waitforresponse is true, else returns nothing</returns>
    Public Function SearchDevice(ByVal SerialNumber As UInt32, Optional ByVal WaitForResponse As Boolean = True, Optional ByVal Timeout As Integer = 4096) As SMADeviceInfoPacket
        SyncLock m_SearchDeviceMutex

            Dim Telegram As New SMATelegram(1, 0, SMATelegram.SMAPacketFlags.SMAF_Broadcast, SMATelegram.SMACommands.CMD_SEARCH_DEV, BitConverter.GetBytes(SerialNumber))

            m_DeviceInfoPacket = Nothing
            SendSMATelegram(Telegram)

            While WaitForResponse AndAlso m_DeviceInfoPacket Is Nothing AndAlso Timeout > 0
                Thread.Sleep(100)
                Timeout -= 100
            End While


            If WaitForResponse Then
                Return m_DeviceInfoPacket
            End If

            Return Nothing
        End SyncLock
    End Function

    ''' <summary>
    ''' Resets all network configuration for all devices (CMD_GET_NET_START)
    ''' all connected devices will respond with their serial number
    ''' After this you will need to assign a new network address
    ''' </summary>
    Public Function StartNetConfiguration(Optional ByVal WaitForResponse As Boolean = True, Optional ByVal Timeout As Integer = 4096) As List(Of SMANetStartpacket)
        SyncLock m_NetStartMutex
            Dim Telegram As New SMATelegram(m_LocalAddress, 0, SMATelegram.SMAPacketFlags.SMAF_Broadcast, SMATelegram.SMACommands.CMD_GET_NET_START, {})

            m_NetstartDevices = New List(Of SMANetStartpacket)

            SendSMATelegram(Telegram)

            If WaitForResponse Then
                Thread.Sleep(Timeout)
                Return m_NetstartDevices
            End If

            Return Nothing
        End SyncLock
    End Function

    ''' <summary>
    ''' Sends a CMD_GET_NET, all devices without assigned network address will respond
    ''' </summary>
    Public Function GetUnconfiguredDevices(Optional ByVal WaitForResponse As Boolean = True, Optional ByVal Timeout As Integer = 4096) As List(Of SMAGetNetpacket)
        SyncLock m_UnconfigredMutex
            Dim Telegram As New SMATelegram(m_LocalAddress, 0, SMATelegram.SMAPacketFlags.SMAF_Broadcast, SMATelegram.SMACommands.CMD_GET_NET, {})

            m_UnconfiguredDevices = New List(Of SMAGetNetpacket)

            SendSMATelegram(Telegram)

            If WaitForResponse Then
                Thread.Sleep(Timeout)
                Return m_UnconfiguredDevices
            End If

            Return Nothing

        End SyncLock
    End Function

    ''' <summary>
    ''' Sets the network address for device with the given serial numnber (CMD_CFG_NETADR)
    ''' </summary>
    ''' <param name="SerialNumber"></param>
    ''' <param name="NetworkAddress"></param>
    ''' <returns>if wait for response=true SMAConfigAddressPacket or nothing if failure. returns nothing if waitforresponse is false</returns>
    Public Function SetNetworkAddress(ByVal SerialNumber As UInt32, ByVal NetworkAddress As UInt16, Optional ByVal WaitForResponse As Boolean = True, Optional ByVal Timeout As Integer = 4096) As SMAConfigAddressPacket
        Dim Data As New MemoryStream
        Data.Write(BitConverter.GetBytes(SerialNumber), 0, 4)
        Data.Write(BitConverter.GetBytes(NetworkAddress), 0, 2)

        m_NetAddressAck = Nothing

        Dim Telegram As New SMATelegram(m_LocalAddress, 0, SMATelegram.SMAPacketFlags.SMAF_Broadcast, SMATelegram.SMACommands.CMD_GET_NET_START, Data.ToArray())
        SendSMATelegram(Telegram)

        While WaitForResponse AndAlso m_NetAddressAck Is Nothing AndAlso Timeout > 0
            Thread.Sleep(100)
            Timeout -= 100
        End While

        If WaitForResponse Then
            Return m_NetAddressAck
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Synchronizes all attached devices to current time and stores channel values
    ''' </summary>
    Public Sub Synchronize()
        Dim Time As UInt32 = (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
        SendSMATelegram(New SMATelegram(m_LocalAddress, 0, SMATelegram.SMAPacketFlags.SMAF_Broadcast, SMATelegram.SMACommands.CMD_SYN_ONLINE, BitConverter.GetBytes(Time)))
    End Sub

    Public Sub Close()
        m_Stream.Close()
        m_ReadThread.Abort()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Close()
    End Sub
End Class
