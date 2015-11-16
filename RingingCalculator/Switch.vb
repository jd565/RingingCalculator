Public Class Switch
    Private port_pin_value As PortPin
    Private switch_running As Boolean
    Public can_be_configured As Boolean
    Public name As String
    Public button As New Button
    Private debounce_state_value As Boolean

    Public ReadOnly Property debounce_state As Boolean
        Get
            Return Me.debounce_state_value
        End Get
    End Property

    Public Property port_pin As PortPin
        Get
            Return Me.port_pin_value
        End Get
        Set(value As PortPin)
            Me.port_pin_value = value
        End Set
    End Property

    ' Function to initialize an object of this class.
    ' This just makes sure that the switch is initialized not running
    Public Sub New(name As String)
        Me.name = name
        Me.switch_running = False

        button.Name = Me.name + "_btn"
        button.Text = "Configure switch"
        AddHandler button.Click, AddressOf configure_button_pressed
    End Sub

    Private Sub debounce_timer_tick(timer As Timer, e As EventArgs)
        Console.WriteLine("debounce timer popped")
        timer.Enabled = False
        Me.debounce_state_value = False
    End Sub

    Public ReadOnly Property isRunning() As Boolean
        Get
            Return Me.switch_running
        End Get
    End Property

    ' Function for changing the state of the switch
    Public Sub trigger_switch()
        If Me.debounce_state_value Then
            GoTo EXIT_LABEL
        End If
        Me.debounce_state_value = True
        start_debounce_timer(AddressOf debounce_timer_tick)
        If Not Me.isRunning Then
            Me.start_running()
        Else
            Me.stop_running()
        End If
EXIT_LABEL:
    End Sub

    ' Function to handle when the program starts recording
    Private Sub start_running()
        For Each bell In GlobalVariables.bells
            bell.reset()
        Next
        Me.switch_running = True
        GlobalVariables.start_time = DateTime.Now()
    End Sub

    ' Function to handle when the program stops recording
    Private Sub stop_running()
        Me.switch_running = False
        GlobalVariables.recording = False
    End Sub

    ' Function to handle when the configure button is pressed.
    Public Sub configure_button_pressed(btn As Button, e As EventArgs)
        Me.can_be_configured = True

        ' We only want 1 configuration happening at a time, so pop out another form to stop this happening
        generate_frmConfigure(btn.Parent)
    End Sub

End Class
