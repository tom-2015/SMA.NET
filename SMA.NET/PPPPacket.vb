Imports System.IO

Public Class PPPPacket

    Protected m_Valid As Boolean 'true if received data is a valid PPP frame (checksum and headers OK)

    Protected m_Address As Byte
    Protected m_Control As Byte
    Protected m_Protocol As UInt16
    Protected m_UserData As Byte()
    Protected m_CheckSum As UInt16

    Public Const PPPStartEnd As Byte = &H7E



    ''' <summary>
    ''' Creates a new PPP frame given received PPP packet
    ''' </summary>
    ''' <param name="PPPFrame"></param>
    Public Sub New(ByVal PPPFrame As Byte())

        If PPPFrame.Length > 2 Then
            If PPPFrame(0) = PPPStartEnd AndAlso PPPFrame(PPPFrame.Length - 1) = PPPStartEnd Then

                Dim EscapedData As Byte()
                ReDim EscapedData(0 To PPPFrame.Length - 3)
                Array.Copy(PPPFrame, 1, EscapedData, 0, PPPFrame.Length - 2)

                Dim UnescapedData As Byte() = PPPUnEscape(EscapedData)
                Dim Reader As New BinaryReader(New MemoryStream(UnescapedData))

                If UnescapedData.Length > 6 Then
                    m_Address = Reader.ReadByte()
                    m_Control = Reader.ReadByte()
                    m_Protocol = Reader.ReadUInt16()
                    Dim UserDataLength As Integer = UnescapedData.Length - 1 - 1 - 2 - 2
                    If UserDataLength > 0 Then
                        m_UserData = Reader.ReadBytes(UserDataLength)
                    Else
                        m_UserData = {}
                    End If
                    m_CheckSum = BitConverter.ToUInt16(Reader.ReadBytes(2).Reverse().ToArray(), 0)
                    m_Valid = m_CheckSum = CalculateChecksum()
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Creates a new PPP frame given address, protocol and user data
    ''' </summary>
    ''' <param name="Address"></param>
    ''' <param name="Control"></param>
    ''' <param name="Protocol"></param>
    ''' <param name="UserData"></param>
    Public Sub New(ByVal Address As Byte, ByVal Control As Byte, ByVal Protocol As UInt16, ByVal UserData As Byte())
        m_Address = Address
        m_Control = Control
        m_Protocol = Protocol
        m_UserData = UserData

    End Sub


    ''' <summary>
    ''' Returns PPP frame
    ''' </summary>
    ''' <returns></returns>
    Public Function GetData() As Byte()
        Dim Packet As New MemoryStream()
        Dim EscapedPacket As New MemoryStream
        Dim EscapedData As Byte()
        Dim Checksum As UInt16 = CalculateChecksum()


        Packet.WriteByte(m_Address)
        Packet.WriteByte(m_Control)
        Packet.Write(BitConverter.GetBytes(m_Protocol), 0, 2)
        Packet.Write(m_UserData, 0, m_UserData.Length)
        Packet.Write(BitConverter.GetBytes(Checksum).Reverse().ToArray(), 0, 2)


        EscapedData = PPPEscape(Packet.ToArray())
        EscapedPacket.WriteByte(PPPStartEnd)
        EscapedPacket.Write(EscapedData, 0, EscapedData.Length)
        EscapedPacket.WriteByte(PPPStartEnd)


        Return EscapedPacket.ToArray()
    End Function


    Public Property Address() As Byte
        Get
            Return m_Address
        End Get
        Set(value As Byte)
            m_Address = value
        End Set
    End Property

    Public Property Control() As Byte
        Get
            Return m_Control
        End Get
        Set(value As Byte)
            m_Control = value
        End Set
    End Property

    Public Property Protocol() As UInt16
        Get
            Return m_Protocol
        End Get
        Set(value As UInt16)
            m_Protocol = value
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

    Public ReadOnly Property Checksum() As UInt16
        Get
            Return m_CheckSum
        End Get
    End Property

    Public ReadOnly Property IsValid() As Boolean
        Get
            Return m_Valid
        End Get
    End Property


    Protected Function CalculateChecksum() As UInt16
        Dim PPContent As New MemoryStream()
        PPContent.WriteByte(m_Address) 'address
        PPContent.WriteByte(m_Control)  'control
        PPContent.Write(BitConverter.GetBytes(m_Protocol), 0, 2) 'protocol header
        PPContent.Write(m_UserData, 0, m_UserData.Length)
        Return fcs16(PPContent.ToArray())
    End Function

    Private Function PPPEscape(ByVal data As Byte()) As Byte()
        Dim Result As New MemoryStream()

        For i As Integer = 0 To data.Length - 1
            Select Case data(i)
                Case &H7D, &H7E, &H11, &H12, &H13
                    Result.WriteByte(&H7D)
                    Result.WriteByte(data(i) Xor &H20)
                Case Else
                    Result.WriteByte(data(i))
            End Select
        Next

        Return Result.ToArray()
    End Function

    Private Function PPPUnEscape(ByVal data As Byte()) As Byte()
        Dim Result As New MemoryStream()
        Dim Escaped As Boolean = False

        For i As Integer = 0 To data.Length - 1
            If Escaped Then
                Result.WriteByte(data(i) Xor &H20)
                Escaped = False
            Else
                If data(i) = &H7D Then
                    Escaped = True
                Else
                    Result.WriteByte(data(i))
                End If
            End If
        Next

        Return Result.ToArray()
    End Function

End Class
