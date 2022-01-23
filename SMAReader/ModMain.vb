Imports System.IO
Imports System.Net
Imports System.Threading
'Imports WiringPiNet
Imports SMADOTNET

Module ModMain

    Dim WithEvents Connection As SMAConnection

    Dim TxPin As GPIOPin
    Public Port As IO.Ports.SerialPort


    Sub SetNetworkAddress(ByVal Serial As String, ByVal Address As String)

        Dim Response As SMAConfigAddressPacket = Connection.SetNetworkAddress(UInt32.Parse(Serial), UInt16.Parse(Address), True)
        If Response IsNot Nothing Then
            Console.WriteLine("Network address for serial " & Serial & " set to: " & Address)
        End If
    End Sub

    Sub FindDevice(ByVal Serial As String)

        Dim Response As SMADeviceInfoPacket = Connection.SearchDevice(UInt32.Parse(Serial))
        If Response IsNot Nothing Then
            Console.WriteLine("Device serial=" & Serial & " address= " & Response.Address)
        End If
    End Sub

    Sub FindUnConfigured()

        Dim Response As List(Of SMAGetNetpacket) = Connection.GetUnconfiguredDevices()

        Console.WriteLine("Found " & Response.Count & " devices.")
        For Each Packet As SMAGetNetpacket In Response
            Console.WriteLine("Found device: serial=" & Packet.Serial & " name=" & Packet.Name & " address=" & Packet.Address)
        Next


    End Sub

    Sub ResetConfiguration()

        Dim Response As List(Of SMANetStartpacket) = Connection.StartNetConfiguration()

        Console.WriteLine("Reset all network address.")
        For Each Packet As SMANetStartpacket In Response
            Console.WriteLine("Found device: serial=" & Packet.Serial & " name=" & Packet.Name & " address=" & Packet.Address)
        Next

    End Sub

    Sub GetChannelValues(ByVal Address As String, ByVal Human As Boolean)
        Dim SunnyBoy As SunnyBoy

        If UInt16.Parse(Address) <= 0 Then
            Console.WriteLine("Error missing or invalid parameter: SMA Address.")
            Return
        End If

        SunnyBoy = New SunnyBoy(Connection, UInt16.Parse(Address))

        Dim Values As SMAChannelValuesPacket = SunnyBoy.GetChannelValues(True, True)

        If Values IsNot Nothing Then
            If Human Then
                Console.WriteLine(" Time: " & (New DateTime(1970, 1, 1, 0, 0, 0).AddSeconds(Values.Time)).ToLocalTime().ToString())
                Console.WriteLine(" Serial: " & Values.Serial)
                Console.WriteLine(" Upv: " & Values.Upv & " V")
                Console.WriteLine(" Ipv: " & PrintFixedPoint(Values.Ipv, 3) & " A")
                Console.WriteLine(" Uac: " & PrintFixedPoint(Values.Uac, 3) & " V")
                Console.WriteLine(" Iac: " & PrintFixedPoint(Values.Iac, 1) & " A")
                Console.WriteLine(" Fac:" & PrintFixedPoint(Values.Fac, 2) & " Hz")
                Console.WriteLine(" Pac: " & Values.Pac & " W")
                Console.WriteLine(" hon: " & Values.HourOn & " H")
                Console.WriteLine(" E-total: " & PrintFixedPoint(Values.ETotal, 3) & " kWh")
                Console.WriteLine(" Status: " & Values.Status)
                Console.WriteLine(" Error: " & Values.ErrNr)
            Else
                Console.WriteLine(Values.ToString())
            End If

        Else
            Console.WriteLine("Failed to get channel values.")
        End If

    End Sub

    Private Function PrintFixedPoint(ByVal Value As Integer, ByVal DecPoint As Integer) As String
        Dim Div As Integer = 1
        For i As Integer = 0 To DecPoint - 1
            Div = Div * 10
        Next
        Return Math.Round(Value / Div, DecPoint)
    End Function

    Sub Main(ByVal Args As String())

        'Dim GPIO As New Gpio(Gpio.NumberingMode.Internal)
        Dim ExitProgram As Boolean = False

        Dim Command As String = ""
        Dim Parameter As String = ""
        Dim ReadValue As Boolean = False
        Dim Parameters As New Dictionary(Of String, String)
        Dim NoValueParameters As String() = {"-h"}

        If Args.Count > 0 Then
            Command = Args(0).ToLower()
        End If
        Parameters.Add("-d", IIf(Environment.OSVersion.Platform = PlatformID.Unix, "/dev/serial0", "COM1"))
        Parameters.Add("-b", "1200")
        Parameters.Add("-p", IO.Ports.Parity.None)
        Parameters.Add("-a", 201)
        Parameters.Add("-s", 1)
        Parameters.Add("-gpio", 4)

        For i As Integer = 1 To Args.Count - 1
            If ReadValue Then
                If Parameters.ContainsKey(Parameter) Then Parameters.Remove(Parameter)
                Parameters.Add(Parameter, Args(i))
                ReadValue = False
            Else
                Parameter = Args(i)
                If NoValueParameters.Contains(Parameter) Then
                    If Not Parameters.ContainsKey(Parameter) Then Parameters.Add(Parameter, "")
                    ReadValue = False
                Else
                    ReadValue = True
                End If
            End If
        Next


        Select Case Command
            Case "config"
                OpenConnection(Parameters)
                SetNetworkAddress(Parameters("-s"), Parameters("-a"))
            Case "find"
                OpenConnection(Parameters)
                FindDevice(Parameters("-s"))
            Case "get"
                OpenConnection(Parameters)
                GetChannelValues(Parameters("-a"), Parameters.ContainsKey("-h"))
            Case "unconfigured"
                OpenConnection(Parameters)
                FindUnConfigured()
            Case "reset"
                OpenConnection(Parameters)
                ResetConfiguration()
            Case Else
                Console.WriteLine("Arguments:")
                Console.WriteLine("SMAReader.exe [command] [options]")
                Console.WriteLine(" config -s [serial_nr] -a [address]    sets network address of device with given serial number.")
                Console.WriteLine(" find -s [serial_nr]                   finds device on the bus with given serial number and returns the network address.")
                Console.WriteLine(" get -a [address] [-h]                 get and print the channel values in json for the sunnyboy device with network address, -h is optional for human readable data.")
                Console.WriteLine(" unconfigured                          finds all serial numbers that do not have a network address set.")
                Console.WriteLine(" reset                                 finds all serial numbers on the RS485 bus and reset their network address.")
                Console.WriteLine("Optional options for all commands:")
                Console.WriteLine(" -d      Serial port device to use, default /dev/serial0")
                Console.WriteLine(" -b      Baud rate, default 1200")
                Console.WriteLine(" -p      Parity check, default 0=None, values: 0, 1: odd, 2: even")
                Console.WriteLine(" -gpio   The GPIO pin to use for turn on TX, default 4, set to -1 to not use a pin.")
        End Select

        If Connection IsNot Nothing Then Connection.Close()

    End Sub

    Private Sub OpenConnection(ByVal Parameters As Dictionary(Of String, String))

        If Environment.OSVersion.Platform = PlatformID.Unix AndAlso Parameters("-gpio") <> "-1" Then
            TxPin = New GPIOPin(Parameters("-gpio"))

            TxPin.SetMode(PinMode.Output)
            TxPin.Write(PinValue.Low)
        End If

        Port = New IO.Ports.SerialPort(Parameters("-d"), Integer.Parse(Parameters("-b")), Integer.Parse(Parameters("-p")), 8, 1) With {
            .NewLine = vbLf
        }
        Port.Open()

        Connection = New SMAConnection(Port.BaseStream)
    End Sub

    Private Sub Connection_ReceivedDataPacket(Connection As SMAConnection, DataPacket As SMADataPacket) Handles Connection.ReceivedDataPacket

    End Sub

    Private Sub Connection_RS485TXOff(Connection As SMAConnection, ByVal BytesTransmitted As Byte()) Handles Connection.RS485TXOff
        If TxPin IsNot Nothing Then
            Dim Delay As Integer = BytesTransmitted.Length * 9 / 1.2
            While Port.BytesToWrite > 0 'wait until bytes written
                Thread.Sleep(10)
            End While

            Thread.Sleep(90) 'add some delay to make sure the last byte is transferred (seems BytesToWrite becomes 0 when the last byte is still transferring)
            TxPin.Write(PinValue.Low)
            Port.DiscardInBuffer()
        End If
    End Sub

    Private Sub Connection_RS485TXOn(Connection As SMAConnection, ByVal BytesTransmitted As Byte()) Handles Connection.RS485TXOn
        If TxPin IsNot Nothing Then
            TxPin.Write(PinValue.High)
        End If
    End Sub

    Public Function BytesToString(ByVal Input As Byte()) As String
        Dim Result As New System.Text.StringBuilder(Input.Length * 2)
        Dim Part As String
        For Each b As Byte In Input
            Part = Conversion.Hex(b)
            If Part.Length = 1 Then Part = "0" & Part
            Result.Append(Part)
        Next
        Return Result.ToString()
    End Function

End Module
