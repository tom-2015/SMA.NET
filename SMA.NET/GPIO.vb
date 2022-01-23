Imports System.IO.File
Imports System.IO

Public Enum PinMode
    Output = 0
    Input = 1
End Enum

Public Enum PinValue As Integer
    Low = 0
    High = 1
End Enum

Public Class GPIOPin

    Protected m_Pin As Integer
    Protected m_fd As Stream

    Public Const GPIOPath As String = "/sys/class/gpio"

    Public Sub New(ByVal Pin As Integer)
        m_Pin = Pin

        If Not File.Exists(GPIOPath & Path.DirectorySeparatorChar & "gpio" & m_Pin & Path.DirectorySeparatorChar & "value") Then
            WriteAllText(GPIOPath & Path.DirectorySeparatorChar & "export", Pin.ToString())
        End If

    End Sub

    Public Sub SetMode(ByVal Mode As PinMode)
        WriteAllText(GPIOPath & Path.DirectorySeparatorChar & "gpio" & m_Pin & Path.DirectorySeparatorChar & "direction", IIf(Mode = PinMode.Output, "out", "in"))
    End Sub

    Public Sub Write(ByVal Value As PinValue)
        WriteAllText(GPIOPath & Path.DirectorySeparatorChar & "gpio" & m_Pin & Path.DirectorySeparatorChar & "value", IIf(Value, "1", "0"))
    End Sub

    Public Function Read() As PinValue
        If ReadAllText(GPIOPath & Path.DirectorySeparatorChar & "gpio" & m_Pin & Path.DirectorySeparatorChar & "value", Text.Encoding.ASCII) = "1" Then
            Return PinValue.High
        Else
            Return PinValue.Low
        End If
    End Function

    Public Sub Close()
        WriteAllText(GPIOPath & Path.DirectorySeparatorChar & "unexport", m_Pin.ToString())
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

    End Sub
End Class
