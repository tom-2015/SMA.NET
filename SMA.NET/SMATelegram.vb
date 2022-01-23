Imports System.IO

Public Class SMATelegram

    Public Enum SMAPacketFlags
        SMAF_Broadcast = 128
        SMAF_Acknowledge = 64
        SMAF_GatewayBlocking = 16
        SMAF_None = 0
    End Enum

    Public Enum SMACommands
        CMD_GET_NET_START = 6
        CMD_GET_NET = 1
        CMD_SEARCH_DEV = 2
        CMD_CFG_NETADR = 3
        CMD_GET_CINFO = 9
        CMD_SET_GRPADR = 5
        CMD_SET_MPARA = 15
        CMD_TNR_VERIFY = 50
        CMD_SYN_ONLINE = 10
        CMD_GET_DATA = 11
        CMD_SET_DATA = 12
        CMD_GET_SINFO = 13
        CMD_GET_MTIME = 20
        CMD_SET_MTIME = 21
        CMD_GET_BINFO = 30
        CMD_GET_BIN = 31
        CMD_SET_BIN = 32
        CMD_VAR_VALUE = 51
        CMD_VAR_FIND = 52
        CMD_VAR_STATUS_OUT = 53
        CMD_VAR_DEFINE_OUT = 54
        CMD_VAR_DEFINE_IN = 56
        CMD_PDELIMIT = 40
        CMD_TEAM_FUNCTION = 60
    End Enum

    Protected m_Source As UInt16
    Protected m_Destination As UInt16
    Protected m_Flags As SMAPacketFlags
    Protected m_Cmd As SMACommands
    Protected m_UserData As Byte()
    Protected m_PacketCnt As Byte



    Public Sub New(ByVal Data As Byte())
        Dim Packet As New BinaryReader(New MemoryStream(Data))
        m_Source = Packet.ReadUInt16()
        m_Destination = Packet.ReadUInt16()
        m_Flags = Packet.ReadByte()
        m_PacketCnt = Packet.ReadByte()
        m_Cmd = Packet.ReadByte()
        Dim UserDataSize As Integer = Data.Length - 2 - 2 - 1 - 1 - 1
        If UserDataSize > 0 Then
            m_UserData = Packet.ReadBytes(UserDataSize)
        Else
            m_UserData = {}
        End If
    End Sub


    Public Sub New(ByVal Source As UInt16, ByVal Destination As UInt16, ByVal Flags As SMAPacketFlags, ByVal Cmd As SMACommands, ByVal UserData As Byte(), Optional ByVal PacketCount As Byte = 0)
        m_Source = Source
        m_Destination = Destination
        m_Flags = Flags
        m_Cmd = Cmd
        m_UserData = UserData
        m_PacketCnt = PacketCount
    End Sub


    Public Function GetData() As Byte()
        Dim Packet As New MemoryStream()
        Packet.Write(BitConverter.GetBytes(m_Source), 0, 2)
        Packet.Write(BitConverter.GetBytes(m_Destination), 0, 2)
        Packet.WriteByte(m_Flags)
        Packet.WriteByte(m_PacketCnt)
        Packet.WriteByte(m_Cmd)
        Packet.Write(m_UserData, 0, m_UserData.Length)
        Return Packet.ToArray()
    End Function

    Public Property Source() As UInt16
        Get
            Return m_Source
        End Get
        Set(value As UInt16)
            m_Source = value
        End Set
    End Property

    Public Property Destination() As UInt16
        Get
            Return m_Destination
        End Get
        Set(value As UInt16)
            m_Destination = value
        End Set
    End Property

    Public Property Flags() As SMAPacketFlags
        Get
            Return m_Flags
        End Get
        Set(value As SMAPacketFlags)
            m_Flags = value
        End Set
    End Property

    Public Property Cmd() As SMACommands
        Get
            Return m_Cmd
        End Get
        Set(value As SMACommands)
            m_Cmd = value
        End Set
    End Property

    Public Property PacketCount() As Byte
        Get
            Return m_PacketCnt
        End Get
        Set(value As Byte)
            m_PacketCnt = value
        End Set
    End Property

    Public Property UserData() As Byte()
        Get
            Return m_UserData
        End Get
        Set(value As Byte())
            m_UserData = value
        End Set
    End Property


End Class
