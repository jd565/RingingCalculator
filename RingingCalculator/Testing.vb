﻿' This module contains functions to run tests on the program.
' Select the test you want to run in the run_tests function.
Module Testing

    ' Main function for running the tests
    Public Sub run_tests(parent As Form)
        frmBells_tests(parent)
        test_wait(2000)
        'test_print_big_row()
        test_ring_hunt_mini()
    End Sub

    Private Sub test_ring_hunt_mini()
        Dim bells As Integer = 4
        Dim ports As Integer = 1

        global_variables_test(bells, ports)

        generate_frmBells(frmMain)

        Dim switch_port_pin As New PortPin("COM1", 0)
        configure_switch_test(switch_port_pin)

        Dim bell_port_pin As New List(Of PortPin)
        For ii As Integer = 1 To bells
            bell_port_pin.Add(New PortPin("COM2", ii))
        Next
        configure_bells_test(bell_port_pin)

        start_timer_test(switch_port_pin)

        Dim hunt_mini As New List(Of Array)
        hunt_mini.Add({2, 1, 4, 3})
        hunt_mini.Add({2, 4, 1, 3})
        hunt_mini.Add({4, 2, 3, 1})
        hunt_mini.Add({4, 3, 2, 1})
        hunt_mini.Add({3, 4, 1, 2})
        hunt_mini.Add({3, 1, 4, 2})
        hunt_mini.Add({1, 3, 2, 4})
        hunt_mini.Add({1, 2, 3, 4})

        For Each row In hunt_mini
            test_ring_this_row(row, bell_port_pin)
        Next


    End Sub

    Private Sub test_print_big_row()
        Dim bells As Integer = 20
        Dim ports As Integer = 1

        global_variables_test(bells, ports)

        generate_frmBells(frmMain)

        Dim switch_port_pin As New PortPin("COM1", 8)
        configure_switch_test(switch_port_pin)

        Dim bell_port_pin As New List(Of PortPin)
        For ii As Integer = 1 To bells
            bell_port_pin.Add(New PortPin("COM2", ii))
        Next
        configure_bells_test(bell_port_pin)

        start_timer_test(switch_port_pin)

        test_ring_this_row({2, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20}, bell_port_pin)

        test_cleanup()
    End Sub

    Private Sub frmBells_tests(parent As Form)
        Dim bells As Integer = 6
        Dim ports As Integer = 3
        ' Generate the global variables
        global_variables_test(bells, ports)

        ' Open the form
        generate_frmBells(parent)

        ' Configure the switch
        Dim switch_port_pin As New PortPin("COM1", 8)
        configure_switch_test(switch_port_pin)

        ' Configure the bells (6 bells)
        Dim bell_port_pin As New List(Of PortPin)
        For ii As Integer = 1 To bells
            bell_port_pin.Add(New PortPin("COM2", ii))
        Next
        configure_bells_test(bell_port_pin)

        ' Test changing the delay value of a bell
        configure_bell_delay_test(GlobalVariables.bells(0))

        ' Now test ringing the bell and check that the light changes
        bell_state_change_test_not_running(GlobalVariables.bells(1), bell_port_pin(1))

        ' Start the timer
        start_timer_test(switch_port_pin)

        ' Test that the state of bells doesn't change
        test_bell_does_not_change_state(GlobalVariables.bells(2), bell_port_pin(2))

        ' ring a change, check that the row is correct
        test_ring_this_row({2, 1, 3, 4, 5, 6}, bell_port_pin)

        ' Before closing the forms cleanup the state of the program
        test_cleanup()

    End Sub

    Private Sub test_cleanup()

        ' Clear the bells from the global variables list
        GlobalVariables.bells.Clear()

        ' Reset the switch
        GlobalVariables.switch = New Switch("switch1")

        ' Set all global variables to their inital values
        GlobalVariables.recording = False
        GlobalVariables.debounce_time = 25

        ' Close all forms apart from the main one
        For Each form In frmMain.OwnedForms
            form.Close()
        Next
    End Sub

    Private Sub test_wait(time As Integer)
        Dim timer As New Timer
        timer.Interval = time
        timer.Enabled = True
        AddHandler timer.Tick, AddressOf disable_timer
        While timer.Enabled = True
            Application.DoEvents()
        End While
    End Sub

    Private Sub test_ring_this_row(row As Array, bell_ports As List(Of PortPin))
        Dim gap_between_bells As Integer = 150
        Dim output_string As String
        For Each ii In row
            port_pin_changed(bell_ports(ii - 1))
            test_wait(gap_between_bells)
        Next
        For Each ii In row
            port_pin_changed(bell_ports(ii - 1))
            test_wait(gap_between_bells)
        Next
        output_string = print_change(GlobalVariables.bells(0).change_times.Count)
        Debug.WriteLine(output_string)
    End Sub

    Private Sub test_bell_does_not_change_state(bell As Bell, port_pin As PortPin)
        Dim old_state As Integer = bell.get_state

        port_pin_changed(port_pin)
        test_wait(50)
        Debug.Assert(old_state = bell.get_state)
        test_wait(50)
    End Sub

    Private Sub disable_timer(Timer As Timer, e As EventArgs)
        Timer.Enabled = False
    End Sub

    Private Sub start_timer_test(port_pin As PortPin)
        port_pin_changed(port_pin)
        Debug.Assert(GlobalVariables.switch.isRunning)
    End Sub

    ' Tests changing state on a bell and checks everything works.
    ' This is a mess...
    Private Sub bell_state_change_test_not_running(bell As Bell, port_pin As PortPin)
        Dim original_state As Integer = bell.get_state

        test_bell_changes_state(bell, port_pin)

        test_bell_delay_works(bell, port_pin)

        test_bell_color_changes(bell, port_pin)

        test_set_bell_state(bell, port_pin, original_state)
    End Sub

    Private Sub test_bell_changes_state(bell As Bell, port_pin As PortPin)
        Dim old_state As Integer = bell.get_state
        port_pin_changed(port_pin)
        test_wait(100)
        Debug.Assert(old_state <> bell.get_state)
    End Sub

    Private Sub test_bell_delay_works(bell As Bell, port_pin As PortPin)
        Dim orig_h_delay As Integer = bell.handstroke_delay
        Dim orig_b_delay As Integer = bell.backstroke_delay
        Dim start_time As DateTime
        Dim end_time As DateTime
        Dim time_diff1 As Integer
        Dim time_diff2 As Integer
        Dim old_state As Integer

        ' Test the delays work.
        old_state = bell.get_state
        bell.fields.handstroke_delay.Value = 100
        bell.fields.backstroke_delay.Value = 100
        start_time = DateTime.Now()
        port_pin_changed(port_pin)
        While old_state = bell.get_state
            Application.DoEvents()
        End While
        end_time = DateTime.Now()
        time_diff1 = (end_time.Ticks - start_time.Ticks) / TimeSpan.TicksPerMillisecond
        Debug.WriteLine("Time diff 1: {0}", time_diff1)

        bell.fields.backstroke_delay.Value = 200
        bell.fields.handstroke_delay.Value = 200
        old_state = bell.get_state
        start_time = DateTime.Now()
        port_pin_changed(port_pin)
        While old_state = bell.get_state
            Application.DoEvents()
        End While
        end_time = DateTime.Now()
        time_diff2 = (end_time.Ticks - start_time.Ticks) / TimeSpan.TicksPerMillisecond
        Debug.WriteLine("Time diff 2: {0}", time_diff2)
        Debug.Assert(time_diff2 > time_diff1)
        bell.fields.backstroke_delay.Value = orig_b_delay
        bell.fields.handstroke_delay.Value = orig_h_delay
    End Sub

    Private Sub test_bell_color_changes(bell As Bell, port_pin As PortPin)
        Dim original_state As Integer = bell.get_state
        Dim old_color As Color
        test_set_bell_state(bell, port_pin, 0)

        ' Test that the colour of the light changes.
        old_color = bell.fields.blob.BackColor
        port_pin_changed(port_pin)
        test_wait(1000)
        Debug.Assert(Not old_color.Equals(bell.fields.blob.BackColor))

        test_set_bell_state(bell, port_pin, original_state)
    End Sub

    Private Sub test_set_bell_state(bell As Bell, port_pin As PortPin, state As Integer)
        While bell.get_state <> state
            port_pin_changed(port_pin)
            test_wait(100)
        End While
    End Sub

    Private Sub configure_bell_delay_test(bell As Bell)
        Dim old_value As Integer
        old_value = bell.fields.handstroke_delay.Value
        bell.fields.handstroke_delay.Value = 100
        Debug.Assert(bell.handstroke_delay = 100)
        bell.fields.handstroke_delay.Value = old_value
        old_value = bell.fields.backstroke_delay.Value
        bell.fields.backstroke_delay.Value = 100
        Debug.Assert(bell.backstroke_delay = 100)
        bell.fields.backstroke_delay.Value = old_value
    End Sub

    Private Sub configure_bells_test(bell_port_pins As List(Of PortPin))
        For ii As Integer = 0 To bell_port_pins.Count - 1
            configure_bell_test(GlobalVariables.bells(ii), bell_port_pins(ii))
        Next
    End Sub

    Private Sub configure_bell_test(bell As Bell, port_pin As PortPin)
        bell.fields.button.PerformClick()
        port_pin_changed(port_pin)
        Debug.Assert(bell.port_pin.Equals(port_pin))
        Debug.Assert(bell.can_be_configured = False)
    End Sub

    Private Sub configure_switch_test(test_port_pin As PortPin)
        GlobalVariables.switch.button.PerformClick()
        Debug.Assert(GlobalVariables.switch.can_be_configured = True)
        ' Send in a PortPin signal
        port_pin_changed(test_port_pin)
        ' Now check that the switch has been configured
        Debug.Assert(GlobalVariables.switch.port_pin.Equals(test_port_pin))
        Debug.Assert(GlobalVariables.switch.can_be_configured = False)
    End Sub

    Private Sub global_variables_test(bells As Integer, ports As Integer)
        ' Generate bells
        GlobalVariables.generate_bells(bells)
        Debug.Assert(GlobalVariables.bells.Count = bells)

        ' Generate COMs
        GlobalVariables.generate_COM_ports(ports)
        Debug.Assert(GlobalVariables.COM_ports.Count = ports)
        Debug.Assert(GlobalVariables.COM_ports_configured = False)
    End Sub

End Module
