Imports System.Threading

Public Class SunnyBoy


    Protected WithEvents m_Connection As SMAConnection
    Protected m_NetworkAddress As UInt16 = 201
    Protected m_ChannelValues As SMAChannelValuesPacket
    Protected m_LastErrNr As Byte
    Protected m_ReceivedChannelValues As Boolean


    Public Event ReceivedChannelValues(ByVal Boy As SunnyBoy, ByVal Values As SMAChannelValuesPacket)
    Public Event ErrorDetected(ByVal Boy As SunnyBoy, ByVal Err As Byte)

    ''' <summary>
    ''' Initializes a new sunnyboy given its network address and a physical connection
    ''' </summary>
    ''' <param name="SMAConnection"></param>
    ''' <param name="NetworkAddress"></param>
    Public Sub New(ByVal SMAConnection As SMAConnection, ByVal NetworkAddress As UInt16)
        m_Connection = SMAConnection
        m_NetworkAddress = NetworkAddress
    End Sub

    ''' <summary>
    ''' Sends command to read the channel values
    ''' </summary>
    ''' <param name="Synchronize">If true broadcast a time synchronization packet first</param>
    ''' 
    Public Function GetChannelValues(Optional ByVal Synchronize As Boolean = True, Optional ByVal WaitForResponse As Boolean = True, Optional ByVal Timeout As Integer = 4000) As SMAChannelValuesPacket
        If Synchronize Then
            m_Connection.Synchronize()
            Thread.Sleep(100)
        End If

        m_ReceivedChannelValues = False
        Dim Telegram As New SMATelegram(m_Connection.LocalNetworkAddress, m_NetworkAddress, SMATelegram.SMAPacketFlags.SMAF_None, SMATelegram.SMACommands.CMD_GET_DATA, {&HF, &H9, &H0})
        m_Connection.SendSMATelegram(Telegram)

        While WaitForResponse AndAlso Timeout > 0 AndAlso Not m_ReceivedChannelValues
            Thread.Sleep(100)
            Timeout -= 100
        End While

        If WaitForResponse AndAlso m_ReceivedChannelValues Then
            Return m_ChannelValues
        End If

        Return Nothing
    End Function


    Private Sub m_Connection_ReceivedDataPacket(Connection As SMAConnection, DataPacket As SMADataPacket) Handles m_Connection.ReceivedDataPacket
        If DataPacket.Telegram.Destination = m_Connection.LocalNetworkAddress AndAlso DataPacket.Telegram.Source = m_NetworkAddress Then
            Select Case DataPacket.Cmd
                Case SMATelegram.SMACommands.CMD_GET_DATA
                    m_ChannelValues = DataPacket
                    m_ReceivedChannelValues = True
                    RaiseEvent ReceivedChannelValues(Me, DataPacket)
                    If m_ChannelValues.ErrNr <> m_LastErrNr Then
                        RaiseEvent ErrorDetected(Me, m_ChannelValues.ErrNr)
                        m_LastErrNr = m_ChannelValues.ErrNr
                    End If
            End Select
        End If
    End Sub

    ''' <summary>
    ''' Returns the last received channel values
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ChannelValues() As SMAChannelValuesPacket
        Get
            Return m_ChannelValues
        End Get
    End Property

End Class
