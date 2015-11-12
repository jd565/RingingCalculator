Public Class Switch
    Private serial_port As String
    Private port_pin As Integer
    Private switch_running As Boolean
    Public can_be_configured As Boolean

    ' Function to set the port and pin
    Public Sub set_port(port As String, pin As Integer)
        Me.serial_port = port
        Me.port_pin = pin
    End Sub

    ' Function for checking the port and pin
    Public Function check_port(port As String, pin As Integer) As Boolean
        Return (port.Equals(Me.serial_port) And pin = Me.port_pin)
    End Function

    ' Function to initialize an object of this class.
    ' This just makes sure that the switch is initialized not running
    Public Sub New()
        Me.switch_running = False
    End Sub


    Public ReadOnly Property isRunning() As Boolean
        Get
            Return Me.switch_running
        End Get
    End Property

    ' Function for changing the state of the switch
    Public Sub trigger_switch()
        If Not Me.isRunning Then
            Me.start_recording()
        Else
            Me.stop_recording()
        End If
    End Sub

    ' Function to handle when the program starts recording
    Private Sub start_recording()
        For Each bell In GlobalVariables.bells
            bell.reset()
        Next
        Me.switch_running = True
        GlobalVariables.start_time = DateTime.Now()
    End Sub

End Class
