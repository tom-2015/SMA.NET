Imports System.IO

''' <summary>
''' General data packet
''' </summary>
Public MustInherit Class SMADataPacket

    Protected m_Data As New MemoryStream
    Protected m_Complete As Boolean = False 'set true when packet receive completed
    Protected m_Telegram As SMATelegram

    Public Sub New(ByVal Telegram As SMATelegram)
        m_Telegram = Telegram
    End Sub

    Public Overridable Sub AddTelegram(ByVal Telegram As SMATelegram)
        AddTelegramData(Telegram.UserData)
    End Sub

    Public Overridable Sub AddTelegramData(ByVal Data As Byte())
        m_Data.Write(Data, 0, Data.Length)
    End Sub


    ''' <summary>
    ''' Creates a new datapacket from a telegram
    ''' </summary>
    ''' <param name="Telegram"></param>
    ''' <returns></returns>
    Public Shared Function CreateDatapacket(ByVal Telegram As SMATelegram) As SMADataPacket
        Dim Packet As SMADataPacket = Nothing
        Select Case Telegram.Cmd
            Case SMATelegram.SMACommands.CMD_SEARCH_DEV
                Packet = New SMADeviceInfoPacket(Telegram)
            Case SMATelegram.SMACommands.CMD_GET_DATA
                Packet = New SMAChannelValuesPacket(Telegram)
            Case SMATelegram.SMACommands.CMD_GET_NET
                Packet = New SMAGetNetpacket(Telegram)
            Case SMATelegram.SMACommands.CMD_GET_NET_START
                Packet = New SMANetStartpacket(Telegram)
            Case SMATelegram.SMACommands.CMD_CFG_NETADR
                Packet = New SMAConfigAddressPacket(Telegram)
        End Select

        Return Packet
    End Function

    ''' <summary>
    ''' Returns true if all data is received, false if waiting for more packets to arrive
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property isComplete() As Boolean
        Get
            Return m_Complete
        End Get
    End Property

    Public ReadOnly Property Cmd() As SMATelegram.SMACommands
        Get
            Return m_Telegram.Cmd
        End Get
    End Property

    Public ReadOnly Property Telegram() As SMATelegram
        Get
            Return m_Telegram
        End Get
    End Property

End Class

''' <summary>
''' Decodes the read channel values
''' </summary>
Public Class SMAChannelValuesPacket

    Inherits SMADataPacket


    Public Property ChannelType As UInt16
    Public Property Channelindex As Byte
    Public Property Records As UInt16
    Public Property Time As UInt32
    Public Property Timebase As UInt32
    Public Property Upv As UInt16
    Public Property Upvn As UInt16
    Public Property Iac As UInt16
    Public Property Iacn As UInt16
    Public Property Uac As UInt16
    Public Property Fac As UInt16
    Public Property Pac As UInt16
    Public Property Zac As UInt16
    Public Property DZac As UInt16
    Public Property Riso As UInt16
    Public Property UacSrr As UInt16
    Public Property FacSrr As UInt16
    Public Property ZacSrr As UInt32
    Public Property IZac As UInt16
    Public Property Ipv As UInt16
    Public Property unknown3 As UInt32
    Public Property unknown4 As UInt32
    Public Property unknown5 As UInt32
    Public Property unknown6 As UInt32
    Public Property RippleCtrlFreq As UInt32
    Public Property ETotal As UInt32
    Public Property HTotal As UInt32
    Public Property HourOn As UInt32
    Public Property PowerOn As UInt32
    Public Property unknown10 As UInt32
    Public Property Serial As UInt32
    Public Property Status As Byte
    Public Property Phase As Byte
    Public Property ErrNr As Byte

    Public Sub New(ByVal Telegram As SMATelegram)
        MyBase.New(Telegram)
    End Sub

    Public Overrides Sub AddTelegramData(ByVal Data As Byte())
        MyBase.AddTelegramData(Data)
        If m_Data.Length >= 92 Then
            m_Complete = True
            m_Data.Position = 0
            Dim Reader As New BinaryReader(m_Data)

            ChannelType = Reader.ReadUInt16()
            Channelindex = Reader.ReadByte()
            Records = Reader.ReadUInt16()
            Time = Reader.ReadUInt32()
            Timebase = Reader.ReadUInt32()
            Upv = Reader.ReadUInt16()
            Upvn = Reader.ReadUInt16()
            Iac = Reader.ReadUInt16()
            Iacn = Reader.ReadUInt16()
            Uac = Reader.ReadUInt16()
            Fac = Reader.ReadUInt16()
            Pac = Reader.ReadUInt16()
            Zac = Reader.ReadUInt16()
            DZac = Reader.ReadUInt16()
            Riso = Reader.ReadUInt16()
            UacSrr = Reader.ReadUInt16()
            FacSrr = Reader.ReadUInt16()
            ZacSrr = Reader.ReadUInt32()
            IZac = Reader.ReadUInt16()
            Ipv = Reader.ReadUInt16()
            unknown3 = Reader.ReadUInt32()
            unknown4 = Reader.ReadUInt32()
            unknown5 = Reader.ReadUInt32()
            unknown6 = Reader.ReadUInt32()
            RippleCtrlFreq = Reader.ReadUInt32()
            ETotal = Reader.ReadUInt32()
            HTotal = Reader.ReadUInt32()
            HourOn = Reader.ReadUInt32()
            PowerOn = Reader.ReadUInt32()
            unknown10 = Reader.ReadUInt32()
            Serial = Reader.ReadUInt32()
            Status = Reader.ReadByte()
            Phase = Reader.ReadByte()
            ErrNr = Reader.ReadByte()

        End If

    End Sub


    Public Overrides Function ToString() As String

        Return "{""channel_type"": " & ChannelType &
                ", ""channel_index"": " & Channelindex &
                ", ""records"": " & Records &
                ", ""time"":" & Time &
                ", ""time_base"": " & Timebase &
                ", ""upv"": " & Upv &
                ", ""upvn"": " & Upvn &
                ", ""iac"": " & Iac &
                ", ""iacn"": " & Iacn &
                ", ""fac"": " & Fac &
                ", ""pac"": " & Pac &
                ", ""zac"": " & Zac &
                ", ""dzac"": " & DZac &
                ", ""riso"": " & Riso &
                ", ""uacsrr"": " & UacSrr &
                ", ""facsrr"": " & FacSrr &
                ", ""zacsrr"": " & ZacSrr &
                ", ""izac"": " & IZac &
                ", ""ipv"": " & Ipv &
                ", ""ripple_ctrl_freq"": " & RippleCtrlFreq &
                ", ""etotal"": " & ETotal &
                ", ""htotal"": " & HTotal &
                ", ""houron"": " & HourOn &
                ", ""poweron"": " & PowerOn &
                ", ""serial"": " & Serial &
                ", ""status"": " & Status &
                ", ""phase"": " & Phase &
                ", ""errnr"": " & ErrNr & "}"

    End Function

End Class

''' <summary>
''' Decodes a discovered device
''' </summary>
Public Class SMADeviceInfoPacket
    Inherits SMADataPacket


    Public Name As String
    Public Serial As UInt32
    Public Address As UInt16

    Public Sub New(ByVal Telegram As SMATelegram)
        MyBase.New(Telegram)
        Address = Telegram.Source
    End Sub

    Public Overrides Sub AddTelegramData(ByVal Data As Byte())
        MyBase.AddTelegramData(Data)

        If m_Data.Length = 12 Then

            m_Complete = True
            m_Data.Position = 0
            Dim Reader As New BinaryReader(m_Data)

            Serial = Reader.ReadUInt32()
            Name = ""
            For i As Integer = 0 To 7
                Dim Value As Byte = Reader.ReadByte()
                If Value <> 0 Then
                    Name = Name & Chr(Value)
                End If
            Next

        End If

    End Sub

    Public Overrides Function ToString() As String
        Return "{ ""serial"":"" " & Serial & ", ""address"": """ & Address & """}"
    End Function


End Class

Public Class SMANetStartpacket
    Inherits SMADeviceInfoPacket


    Public Sub New(ByVal Telegram As SMATelegram)
        MyBase.New(Telegram)
    End Sub

End Class

Public Class SMAGetNetpacket
    Inherits SMADeviceInfoPacket


    Public Sub New(ByVal Telegram As SMATelegram)
        MyBase.New(Telegram)
    End Sub

End Class

Public Class SMAConfigAddressPacket
    Inherits SMADataPacket


    Public Serial As UInt32

    Public Sub New(ByVal Telegram As SMATelegram)
        MyBase.New(Telegram)
    End Sub

    Public Overrides Sub AddTelegramData(ByVal Data As Byte())
        MyBase.AddTelegramData(Data)

        If m_Data.Length = 4 Then

            m_Complete = True
            m_Data.Position = 0
            Dim Reader As New BinaryReader(m_Data)

            Serial = Reader.ReadUInt32()

        End If

    End Sub

    Public Overrides Function ToString() As String
        Return "{ ""serial"":"" " & Serial & "}"
    End Function


End Class