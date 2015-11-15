Module BellRingingFunctions

    ' Wrapper function for a pin on a port changing
    Public Sub port_pin_changed_wrapper(port As IO.Ports.SerialPort, e As IO.Ports.SerialPinChangedEventArgs)
        Dim port_pin As New PortPin(port.PortName, e.EventType)
        port_pin_changed(port_pin)
    End Sub

    ' Function to handle a pin on a COM port changing.
    Public Sub port_pin_changed(port_pin As PortPin)

        ' A pin on a port changing can mean a few things:
        ' A bell is waiting to be configured and this event sets the configuration
        ' Bell is being rung to check delays and this event triggers the light
        ' The switch is being configured and this event sets the config
        ' The switch is being pressed to start or stop recording
        ' We are recording and a bell has just been rung
        If Not GlobalVariables.switch.isRunning Then
            If event_configures_bell(port_pin) Then
                GoTo EXIT_LABEL
            End If
            If event_rings_bell(port_pin) Then
                GoTo EXIT_LABEL
            End If
            If event_configures_switch(port_pin) Then
                GoTo EXIT_LABEL
            End If
            If event_triggers_switch(port_pin) Then
                GoTo EXIT_LABEL
            End If
        Else
            If event_rings_bell(port_pin) Then
                GoTo EXIT_LABEL
            End If
            If event_triggers_switch(port_pin) Then
                GoTo EXIT_LABEL
            End If
        End If

EXIT_LABEL:
    End Sub

    ' Function to check if a bell can be configured, and if it can to configure it.
    Private Function event_configures_bell(port_pin As PortPin) As Boolean
        For Each bell In GlobalVariables.bells
            If bell.can_be_configured Then
                Console.WriteLine("Configuring bell {0}", bell.bell_number)
                bell.port_pin = port_pin
                item_has_been_configured()
                Return True
            End If
        Next

        ' If we have got here then we haven't found any bells to configure.
        Return False
    End Function

    ' Function to check if this port and pin are set to any of the bells, and if they are to ring them
    Private Function event_rings_bell(port_pin As PortPin) As Boolean
        For Each bell In GlobalVariables.bells
            If port_pin.Equals(bell.port_pin) Then
                Console.WriteLine("Ringing bell {0}", bell.bell_number)
                bell.trigger_bell()
                Return True
            End If
        Next
        Return False
    End Function

    ' Function to check if we are configuring the switch.
    Private Function event_configures_switch(port_pin As PortPin) As Boolean
        If GlobalVariables.switch.can_be_configured Then
            Console.WriteLine("Configuring switch")
            GlobalVariables.switch.port_pin = port_pin
            item_has_been_configured()
            Return True
        End If
        Return False
    End Function

    ' Function to check if the port matches the switch.
    Private Function event_triggers_switch(port_pin As PortPin) As Boolean
        If port_pin.Equals(GlobalVariables.switch.port_pin) Then
            Console.WriteLine("Triggering switch.")
            GlobalVariables.switch.trigger_switch()
            Return True
        End If
        Return False
    End Function

    Public Sub start_new_timer(interval As Integer, proc As EventHandler)
        Dim timer As New Timer
        timer.Interval = interval
        AddHandler timer.Tick, proc
        timer.Enabled = True
    End Sub

    Public Sub start_debounce_timer(proc As EventHandler)
        start_new_timer(GlobalVariables.debounce_time, proc)
    End Sub

End Module
