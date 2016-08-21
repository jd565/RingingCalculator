Public Class GlobalVariables
    Public Shared bells As New List(Of Bell)
    Public Shared COM_ports As New List(Of COMPort)
    Public Shared config_loaded As Boolean = False
    Public Shared switch As New Switch("switch1")
    Public Shared method_started As Boolean = False
    Public Shared debounce_time As Integer = 25
    Public Shared changes_per_peal As Integer = 5040
    Public Shared changes_per_lead As Integer = 40
    Public Shared changes_per_course As Integer = 360
    Public Shared leads_per_course As Integer = 9
    Public Shared cpm_string_format As String = "##0.0"
    Public Shared hours_and_mins As String = "hh\:mm"
    Public Shared hours_mins_seconds As String = "hh\:mm\:ss"
    Public Shared full_time As String = "hh\:mm\:ss\.f"
    Public Shared bell_light_time As Integer = 200
    Public Shared start_index As Integer = 0
    Public Shared start_time As DateTime
    Public Shared stop_index As Integer = 0
    Public Shared statistics_init As Boolean = False
    Public Shared input_tracer As Boolean = False

    ' Function to reset the global variables, and clear all lists.
    Public Shared Sub reset()
        GlobalVariables.bells.Clear()
        GlobalVariables.close_COM_ports()
        GlobalVariables.COM_ports.Clear()
        GlobalVariables.config_loaded = False
        GlobalVariables.switch = New Switch("switch1")
        GlobalVariables.method_started = False
    End Sub

    Public Shared Sub wait(time As Integer)
        Dim timer As New Timer
        timer.Interval = time
        timer.Enabled = True
        AddHandler timer.Tick, AddressOf GlobalVariables.stop_timer
        While timer.Enabled = True
            Application.DoEvents()
        End While
    End Sub

    Private Shared Sub stop_timer(t As Timer, e As EventArgs)
        t.Enabled = False
    End Sub

    ' Function to close all COM ports
    Public Shared Sub close_COM_ports()
        For Each port In GlobalVariables.COM_ports
            If port.IsOpen Then
                port.Close()
            End If
        Next
    End Sub

    ' Function to fill the Global Variable COM port list with the number of ports specified.
    Public Shared Sub generate_COM_ports(ByVal ports As Integer)
        Dim port As COMPort

        ' Make sure all the COM ports are closed
        GlobalVariables.close_COM_ports()

        GlobalVariables.COM_ports.Clear()
        For ii As Integer = 1 To ports
            port = New COMPort
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

    ' Function to connect the COM ports.
    ' This should only be called after the ports have been filled in.
    ' This function returns a boolean of whether it has succeeded in connecting all ports.
    Public Shared Function connect_COM_ports() As Boolean
        Dim success As Boolean = True

        ' Connect the ports one at a time.
        For Each port In GlobalVariables.COM_ports

            ' Try to open a port. If it fails, run the code in the Catch part.
            Try
                port.Open()
                port.RtsEnable = True
            Catch ex As Exception
                RcDebug.debug_print("Tried to open port " & port.PortName & ", but failed.")
                MsgBox("Failed to open port " & port.PortName & ".",, "Error")
                success = False
            End Try
        Next

        ' If any of the ports have failed then we should make sure all ports are closed.
        If Not success Then
            GlobalVariables.close_COM_ports()
        End If

        Return success
    End Function

End Class
