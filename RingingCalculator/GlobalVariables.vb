Public Class GlobalVariables
    Public Shared bells As New List(Of Bell)
    Public Shared COM_ports As New List(Of IO.Ports.SerialPort)
    Public Shared COM_ports_configured As Boolean = False
    Public Shared switch As New Switch
    Public Shared start_time As DateTime

    ' Function to fill the Global Variable COM port list with the number of ports specified.
    Public Shared Sub generate_COM_ports(ports As Integer)
        Dim port As IO.Ports.SerialPort

        GlobalVariables.COM_ports.Clear()
        For ii As Integer = 1 To ports
            port = New IO.Ports.SerialPort
            AddHandler port.PinChanged, AddressOf port_pin_changed
            GlobalVariables.COM_ports.Add(New IO.Ports.SerialPort())
        Next
    End Sub

    ' Function to populate the bells list with the number of bells passed in.
    Public Shared Sub generate_bells(ByVal bells As Integer)
        GlobalVariables.bells.Clear()
        For ii As Integer = 1 To bells
            GlobalVariables.bells.Add(New Bell("bell" + ii.ToString))
        Next
    End Sub

End Class
