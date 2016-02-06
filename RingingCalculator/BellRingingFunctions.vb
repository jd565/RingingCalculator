Module BellRingingFunctions

    ' Wrapper function for a pin on a port changing
    Public Sub port_pin_changed_wrapper(port As COMPort, e As IO.Ports.SerialPinChangedEventArgs)
        Dim port_pin As New PortPin(port.PortName, e.EventType)
        Dim frm As Form
        frm = Form.ActiveForm
        If frm Is Nothing Then
            Exit Sub
        End If
        If frm.InvokeRequired Then
            Try
                frm.Invoke(New port_pin_changed_del(AddressOf port_pin_changed_wrapper), New Object() {port, e})
            Catch ex As System.InvalidOperationException
                Console.WriteLine("Hit an InvalidOperationException")
            End Try
        Else
            ' Even if we are on the active form this still may throw an exception, so catch it here
            Try
                port_pin_changed(port_pin)
            Catch ex As System.InvalidOperationException
                Console.WriteLine("Hit an InvalidOperationException")
            End Try
        End If
    End Sub

    Delegate Sub port_pin_changed_del(port As IO.Ports.SerialPort, e As IO.Ports.SerialPinChangedEventArgs)

    ' Function to handle a pin on a COM port changing.
    Public Sub port_pin_changed(port_pin As PortPin)
        Dim frm As frmInputTracer

        Console.WriteLine("Port pin has changed")

        ' A pin on a port changing can mean a few things:
        ' A bell is waiting to be configured and this event sets the configuration
        ' Bell is being rung to check delays and this event triggers the light
        ' The switch is being configured and this event sets the config
        ' The switch is being pressed to start or stop recording
        ' We are recording and a bell has just been rung
        If GlobalVariables.input_tracer Then
            frm = find_form(GetType(frmInputTracer))
            If frm IsNot Nothing Then
                frm.port_pin_triggered(port_pin)
            End If
        End If
        If event_rings_bell(port_pin) Then
            Exit Sub
        End If
        If event_triggers_switch(port_pin) Then
            Exit Sub
        End If
        If event_configures_input(port_pin) Then
            Exit Sub
        End If

    End Sub

    ' Function to check if a bell can be configured, and if it can to configure it.
    Private Function event_configures_input(port_pin As PortPin) As Boolean
        Dim frm As frmConfigure
        frm = find_form(GetType(frmConfigure))
        If frm IsNot Nothing Then
            Console.WriteLine("Configuring {0}", frm.input.name)
            frm.input.configure(port_pin)
            frm.item_has_been_configured()
            Return True
        End If

        ' If we have got here then we haven't found any bells to configure.
        Return False
    End Function

    ' Function to check if this port and pin are set to any of the bells, and if they are to ring them
    Private Function event_rings_bell(port_pin As PortPin) As Boolean
        For Each bell In GlobalVariables.bells
            If port_pin.Equals(bell.port_pin) Then
                Console.WriteLine("Ringing bell {0}", bell.bell_number)
                bell.trigger_input()
                Return True
            End If
        Next
        Return False
    End Function

    ' Function to check if the port matches the switch.
    Private Function event_triggers_switch(port_pin As PortPin) As Boolean
        If port_pin.Equals(GlobalVariables.switch.port_pin) Then
            Console.WriteLine("Triggering switch.")
            GlobalVariables.switch.trigger_input()
            Return True
        End If
        Return False
    End Function

    ' Function to start a timer with the specified interval in ms
    ' calls proc when it finishes
    Public Sub start_new_timer(interval As Integer, proc As EventHandler)
        Dim timer As New Timer
        timer.Interval = interval
        AddHandler timer.Tick, proc
        timer.Enabled = True
    End Sub

    ' Function wrapper for debounce timers
    Public Sub start_debounce_timer(proc As EventHandler)
        start_new_timer(GlobalVariables.debounce_time, proc)
    End Sub

    ' Function to see if we should begin recording
    ' This checks for either:
    '   There is only 1 bell, so start now
    '   We have hit 21.
    Public Function has_method_started(change_id As Integer) As Boolean
        Dim row As Row = Statistics.rows(change_id)

        If GlobalVariables.bells.Count = 1 Then
            Console.WriteLine("Method has started")
            GlobalVariables.method_started = True
            Statistics.changes = 0
            GlobalVariables.start_index = change_id
            GlobalVariables.start_time = Statistics.rows(change_id).time
            Return True
        End If

        If row.bells(0).bell = 2 And row.bells(1).bell = 1 Then
            Console.WriteLine("Method has started")
            GlobalVariables.method_started = True
            Statistics.changes = 0
            GlobalVariables.start_index = change_id
            GlobalVariables.start_time = Statistics.rows(change_id - 1).time
            Return True
        End If
        Return False
    End Function

    ' Function to call when a bell has just rung
    ' We add the bell straight to the row we expect it to be in
    ' Need to make sure that the row does actually exist
    Public Sub bell_has_just_rung(bell As Bell)
        Dim change_id As Integer = bell.change_times.Count - 1

        ' This is checking that we have a row to put our bell in.
        ' If this is the first bell in a row, we need to add a new row to
        ' the list.
        If Statistics.rows.Count <= change_id Then
            Statistics.rows.Add(New Row)
        End If

        ' This is sufficient as we should never get in the state of trying to add
        ' a change that is more than 1 ahead of the size of the rows.
        ' If we do, this will raise an exception and everything will fail.
        Statistics.rows(change_id).add(bell.change_times.Last)

        If Statistics.rows(change_id).row_is_full Then
            row_is_full(change_id)
        End If
    End Sub

End Module
