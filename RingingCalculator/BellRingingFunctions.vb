Module BellRingingFunctions

    ' Function to handle a pin on a COM port changing.
    Public Sub port_pin_changed(port As IO.Ports.SerialPort, e As IO.Ports.SerialPinChangedEventArgs)

        ' If a pin changed while we aren't running, check for new configuration
        If Not GlobalVariables.switch.isRunning Then
            Console.WriteLine("Not running, check bell to configure.")
            For Each bell In GlobalVariables.bells
                If bell.can_be_configured Then
                    Console.WriteLine("Found bell. Configuring " & bell.name)
                    bell.set_port(port.PortName, e.EventType)
                    bell_has_been_configured()
                    GoTo EXIT_LABEL
                End If
            Next

            ' It is possible that we are configuring or starting the switch.
            If GlobalVariables.switch.can_be_configured Then
                Console.WriteLine("Configure the switch")
                GlobalVariables.switch.set_port(port.PortName, e.EventType)
                GlobalVariables.switch.can_be_configured = False
                GoTo EXIT_LABEL
            End If

            If GlobalVariables.switch.check_port(port.PortName, e.EventType) Then
                Console.WriteLine("switch has been pressed")
                GlobalVariables.switch.trigger_switch()
                GoTo EXIT_LABEL
            End If

        Else
            ' We are running, so check the bells to see which one we matched.
            Console.WriteLine("Running, check which bell rung.")
            For Each bell In GlobalVariables.bells
                If bell.check_port(port.PortName, e.EventType) Then
                    Console.WriteLine("Found bell. " & bell.name & " rung.")
                    bell.trigger_bell()
                    GoTo EXIT_LABEL
                End If
            Next
        End If
EXIT_LABEL:
    End Sub

End Module
