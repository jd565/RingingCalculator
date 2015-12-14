Public Class GlobalVariables
    Public Shared bells As New List(Of Bell)
    Public Shared COM_ports As New List(Of IO.Ports.SerialPort)
    Public Shared COM_ports_configured As Boolean = False
    Public Shared switch As New Switch("switch1")
    Public Shared start_time As DateTime
    Public Shared recording As Boolean = False
    Public Shared debounce_time As Integer = 25
    Public Shared changes_per_peal As Integer = 5040
    Public Shared changes_per_lead As Integer = 40
    Public Shared changes_per_course As Integer = 360
    Public Shared leads_per_course As Integer = 9

    ' Function to fill the Global Variable COM port list with the number of ports specified.
    Public Shared Sub generate_COM_ports(ByVal ports As Integer)
        Dim port As IO.Ports.SerialPort

        ' Make sure all the COM ports are closed
        For Each port In GlobalVariables.COM_ports
            If port.IsOpen Then
                port.Close()
            End If
        Next

        GlobalVariables.COM_ports.Clear()
        For ii As Integer = 1 To ports
            port = New IO.Ports.SerialPort
            AddHandler port.PinChanged, AddressOf port_pin_changed_wrapper
            GlobalVariables.COM_ports.Add(port)
        Next
    End Sub

    ' Function to populate the bells list with the number of bells passed in.
    Public Shared Sub generate_bells(ByVal bells As Integer)
        GlobalVariables.bells.Clear()
        For ii As Integer = 1 To bells
            GlobalVariables.bells.Add(New Bell(ii))
        Next
    End Sub

    ' Function to handle the debounce time changing
    Public Shared Sub debounce_time_changed(debounce_time As Integer)
        GlobalVariables.debounce_time = debounce_time
    End Sub

    ' Function to update changes per course
    Public Shared Sub update_changes_per_course()
        GlobalVariables.changes_per_course = GlobalVariables.changes_per_lead * GlobalVariables.leads_per_course
    End Sub

End Class
