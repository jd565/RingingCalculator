#If DEBUG Then
' This module contains functions to run tests on the program.
' Select the test you want to run in the run_tests function.
Module Testing

    Public test_mode As Boolean = False

    ' Main function for running the tests
    Public Sub run_tests(parent As Form)
        Testing.test_mode = True
        'frmBells_tests(parent)
        'test_wait(2000)
        'test_print_big_row()
        'test_ring_hunt_mini()
        'test_notation()
        'test_method_gen()
        'test_input_tracer(parent)
        'test_place_stats()
        'test_composition()
        'test_composed_method()
        'test_peal(parent)
        test_ring_cambridge_minor()
        Testing.test_mode = False
    End Sub

    Private Sub test_ring_cambridge_minor()
        Dim method As New Method("&-36-14-12-36-14-56 le12", 6)
        method.generate()
        Dim cam_minor As New List(Of Array)
        Dim row_array(5) As Integer
        Dim ii As Integer
        cam_minor.Add({1, 2, 3, 4, 5, 6})
        cam_minor.Add({1, 2, 3, 4, 5, 6})
        For Each row In method.rows
            cam_minor.Add({row.bells(0).bell, row.bells(1).bell, row.bells(2).bell, row.bells(3).bell, row.bells(4).bell, row.bells(5).bell})
        Next
        cam_minor.Add({1, 2, 3, 4, 5, 6})
        cam_minor.Add({1, 2, 3, 4, 5, 6})

        Dim bells As Integer = 6
        Dim ports As Integer = 2
        Dim pp As PortPin
        Dim frm As frmStats

        global_variables_test(bells, ports)

        frmPerf.changes_per_lead.Text = "24"
        frmPerf.leads_per_course.Text = "5"

        Dim switch_port_pin As New PortPin("COM1", 0)
        GlobalVariables.switch.port_pin = switch_port_pin

        Dim bell_port_pin As New List(Of PortPin)
        For ii = 1 To bells
            pp = New PortPin("COM2", ii)
            bell_port_pin.Add(pp)
            GlobalVariables.bells(ii - 1).port_pin = pp
        Next

        GlobalVariables.config_loaded = True

        test_start_timer(switch_port_pin)

        For Each row In cam_minor
            test_ring_this_row(row, bell_port_pin, 25)
        Next

        'test_stop_timer(switch_port_pin)

        frm = find_form(GetType(frmStats))
        frm.btn_view_stats.PerformClick()

        test_save_rows()
    End Sub

    Private Sub test_peal(parent As Form)
        Dim frm As New frmMethod(parent)
        frm.place_notation.Text = "b &-3-4-25-36-4-5-6-7"
        frm.bells_text.Text = "8"

        ' This is the composition for a peal of Cambridge Major
        ' found from http://ringing.org/main/pages/peals/major/single/cambridge
        '5,090 Cambridge Surprise Major
        'Robert D S Brown
        '
        '234567   V  B  I  M  W  F  H
        '34256                      2 
        '45362             2  2     3 
        '56423             2  2     3 
        '62534             2  2     3 
        '23645       -              3 
        '63542             -          
        '34625       -                
        '756324                  -  - 
        '(324567)  -  -  -     s        
        frm.composition.Text = "5  3  2  M  W  4  H" & vbCrLf & "                  2 " & vbCrLf & "         2  2     3 " & vbCrLf & "         2  2     3 " & vbCrLf & "         2  2     3 " & vbCrLf & "   -              3 " & vbCrLf & "         -          " & vbCrLf & "   -                " & vbCrLf & "               -  - " & vbCrLf & "-  -  -     s  "
    End Sub

    Private Sub test_composed_method()
        Dim frm As frmPerfStats
        Dim full_comp As String = "W H" & Chr(13) & Chr(10) & "  3"
        Dim method As New Method("b &-3-4-25-36-4-5-6-7", 8)
        method.add_composition(full_comp)
        method.generate()
        Debug.WriteLine(method.rows.Count)
        For Each row In method.rows
            Debug.WriteLine(row.print())
        Next
        frm = New frmPerfStats(frmPerf, method)
    End Sub

    Private Sub test_composition()
        Dim comp As Composition
        Dim full_comp As String = "W M  H" & Chr(13) & Chr(10) & "- 2" & vbCrLf & "  -  2s"
        comp = New Composition(full_comp)
        For Each c In comp.composition
            Debug.WriteLine(c.call_to_make & " at " & c.location)
        Next
    End Sub

    Private Sub test_place_stats()
        Dim place_stats As New List(Of PlaceStats)
        Dim place_stats1 As New PlaceStats
        Dim place_stats2 As New PlaceStats

        place_stats1.add(1000, True)
        place_stats1.add(2000, True)
        place_stats1.add(3000, True)
        place_stats1.add(4000, True)
        place_stats1.add(5000, False)
        place_stats1.add(6000, False)
        place_stats1.add(7000, False)
        place_stats1.add(8000, False)
        Debug.Assert(place_stats1.h_average = 2500)
        Debug.Assert(place_stats1.b_average = 6500)
        Debug.Assert(place_stats1.average = 4500)
        place_stats2.add(8000, True)
        place_stats2.add(7000, True)
        place_stats2.add(6000, True)
        place_stats2.add(5000, True)
        place_stats2.add(4000, False)
        place_stats2.add(3000, False)
        place_stats2.add(2000, False)
        place_stats2.add(1000, False)
        Debug.Assert(place_stats2.h_average = 6500)
        Debug.Assert(place_stats2.b_average = 2500)
        Debug.Assert(place_stats2.average = 4500)

        place_stats.Add(place_stats1)
        place_stats.Add(place_stats2)

        Debug.Assert(place_stats.Average(Function(s) s.h_average) = 4500)
        Debug.Assert(place_stats.Average(Function(s) s.b_average) = 4500)
        Debug.Assert(place_stats.Average(Function(s) s.average) = 4500)

    End Sub

    'Tests the input tracer method
    Private Sub test_input_tracer(p As Form)
        Dim frm As New frmInputTracer(p)
        Dim real_pp As New PortPin("COM1", 2)
        Dim fake_pp As New PortPin("COM1", 1)
        Dim state As Integer

        state = frm.status
        port_pin_changed(real_pp)
        Debug.Assert(state <> frm.status)
        state = frm.status
        port_pin_changed(fake_pp)
        Debug.Assert(state = frm.status)
        test_wait(300)
        port_pin_changed(real_pp)
        Debug.Assert(state <> frm.status)
        test_wait(250)
        frm.reset.PerformClick()
        Debug.Assert(frm.status = 0)
        frm.debounce_enable.Checked = True
        state = frm.status
        port_pin_changed(real_pp)
        port_pin_changed(real_pp)
        Debug.Assert(state <> frm.status)
        frm.reset.PerformClick()
        frm.hold_enable.Checked = True
        test_wait(100)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
        test_wait(5)
        port_pin_changed(real_pp)
        test_wait(25)
        port_pin_changed(real_pp)
        test_wait(30)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
        test_wait(15)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
        frm.debounce_enable.Checked = True
        frm.debounce_value.Value = 100
        test_wait(5500)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
        test_wait(5)
        port_pin_changed(real_pp)
        test_wait(25)
        port_pin_changed(real_pp)
        test_wait(30)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
        test_wait(15)
        port_pin_changed(real_pp)
        test_wait(10)
        port_pin_changed(real_pp)
    End Sub

    ' Tests generating a method and prints all the rows to screen
    Private Sub test_method_gen()
        Dim frm As frmPerfStats
        Dim method As New Method("b &-3-4-25-36-4-5-6-7", 8)
        method.generate()
        Debug.WriteLine(method.rows.Count)
        For Each row In method.rows
            Debug.WriteLine(row.print())
        Next
        frm = New frmPerfStats(frmPerf, method)
    End Sub

    ' Tests generating a set of rows from place notation
    Private Sub test_notation()
        Dim n As New PlaceNotation("b &-3-4-25-36-4-5-6-7", 8)
        n.parse()
        For Each s In n.main_block
            Debug.Write(s.notation & ",")
        Next
        Debug.WriteLine("")

        For Each s In n.main_block
            Debug.Write(s.fill_notation(8) & ",")
        Next
        Debug.WriteLine("")

        For Each s In n.main_block
            For Each i In s.change_hash(8)
                Debug.Write(i.ToString & ", ")
            Next
            Debug.WriteLine("")
        Next
    End Sub

    ' Rings plain hunt on 4
    Private Sub test_ring_hunt_mini()
        Dim bells As Integer = 4
        Dim ports As Integer = 1
        Dim pp As PortPin
        Dim frm As frmStats

        global_variables_test(bells, ports)

        frmPerf.changes_per_lead.Text = "2"
        frmPerf.leads_per_course.Text = "200"

        Dim switch_port_pin As New PortPin("COM1", 0)
        GlobalVariables.switch.port_pin = switch_port_pin

        Dim bell_port_pin As New List(Of PortPin)
        For ii As Integer = 1 To bells
            pp = New PortPin("COM2", ii)
            bell_port_pin.Add(pp)
            GlobalVariables.bells(ii - 1).port_pin = pp
        Next

        GlobalVariables.config_loaded = True

        test_start_timer(switch_port_pin)

        Dim hunt_mini As New List(Of Array)
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({2, 1, 4, 3})
        hunt_mini.Add({2, 4, 1, 3})
        hunt_mini.Add({4, 2, 3, 1})
        hunt_mini.Add({4, 3, 2, 1})
        hunt_mini.Add({3, 4, 1, 2})
        hunt_mini.Add({3, 1, 4, 2})
        hunt_mini.Add({1, 3, 2, 4})
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({1, 2, 3, 4})
        hunt_mini.Add({1, 2, 3, 4})

        For Each row In hunt_mini
            test_ring_this_row(row, bell_port_pin)
        Next

        test_stop_timer(switch_port_pin)

        frm = find_form(GetType(frmStats))
        frm.btn_view_stats.PerformClick()

        test_save_rows()

    End Sub

    ' Function to test saving the rows to a file
    Private Sub test_save_rows()
        save_statistics("test_save.txt")
    End Sub

    ' Function to test saving the config
    Private Sub test_save_config()
        save_config()
    End Sub

    Private Sub test_print_big_row()
        Dim bells As Integer = 20
        Dim ports As Integer = 1

        global_variables_test(bells, ports)

        'generate_frmBells(frmMain)

        Dim switch_port_pin As New PortPin("COM1", 8)
        configure_switch_test(switch_port_pin)

        Dim bell_port_pin As New List(Of PortPin)
        For ii As Integer = 1 To bells
            bell_port_pin.Add(New PortPin("COM2", ii))
        Next
        configure_bells_test(bell_port_pin)

        test_start_timer(switch_port_pin)

        test_ring_this_row({2, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20}, bell_port_pin)

        test_cleanup()
    End Sub

    Private Sub frmBells_tests(parent As Form)
        Dim bells As Integer = 6
        Dim ports As Integer = 3
        ' Generate the global variables
        global_variables_test(bells, ports)

        ' Open the form
        'generate_frmBells(parent)

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
        test_start_timer(switch_port_pin)

        ' Test that the state of bells doesn't change
        test_bell_does_not_change_state(GlobalVariables.bells(2), bell_port_pin(2))

        ' ring a change, check that the row is correct
        test_ring_this_row({2, 1, 3, 4, 5, 6}, bell_port_pin)

        test_stop_timer(switch_port_pin)

        ' Before closing the forms cleanup the state of the program
        test_cleanup()

    End Sub

    Private Sub test_stop_timer(pp As PortPin)
        ' Only stop the switch if it is running
        If GlobalVariables.switch.is_running Then
            port_pin_changed(pp)
            ' Unpress the switch
            test_wait(50)
            port_pin_changed(pp)
        End If
        Debug.Assert(Not GlobalVariables.switch.is_running)
    End Sub

    Private Sub test_cleanup()

        ' Clear the bells from the global variables list
        GlobalVariables.bells.Clear()

        ' Reset the switch
        GlobalVariables.switch = New Switch("switch1")

        ' Set all global variables to their inital values
        GlobalVariables.method_started = False
        GlobalVariables.debounce_time = 25

        ' Close all forms apart from the main one
        For Each form In frmPerf.OwnedForms
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

    Private Sub test_ring_this_row(row As Array, bell_ports As List(Of PortPin), Optional delay As Integer = 155)
        Randomize()
        Dim output_string As String
        For ii = 0 To row.Length - 1
            port_pin_changed(bell_ports(row(ii) - 1))

            ' This sets the delay of bell 1 to be larger than the others
            If ii <> row.Length - 1 AndAlso row(ii + 1) = 1 Then
                test_wait(Math.Floor(10 * Rnd()) + delay)
            Else
                test_wait(Math.Floor(10 * Rnd()) + delay - 10)
            End If
        Next
        For ii = 0 To row.Length - 1
            port_pin_changed(bell_ports(row(ii) - 1))
            test_wait(Math.Floor(10 * Rnd()) + delay)
        Next
        output_string = Statistics.rows.Last.print
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

    Private Sub test_start_timer(port_pin As PortPin)
        port_pin_changed(port_pin)
        Debug.Assert(GlobalVariables.switch.is_running)
        ' Unpress switch
        test_wait(50)
        port_pin_changed(port_pin)
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
        ' Now move the bell back to state 0
        test_set_bell_state(bell, port_pin, 0)
    End Sub

    Private Sub configure_switch_test(test_port_pin As PortPin)
        GlobalVariables.switch.button.PerformClick()
        Debug.Assert(GlobalVariables.switch.can_be_configured = True)
        ' Send in a PortPin signal
        port_pin_changed(test_port_pin)
        ' Now check that the switch has been configured
        Debug.Assert(GlobalVariables.switch.port_pin.Equals(test_port_pin))
        Debug.Assert(GlobalVariables.switch.can_be_configured = False)
        ' Now unpress the switch
        test_wait(50)
        port_pin_changed(test_port_pin)
    End Sub

    Private Sub global_variables_test(bells As Integer, ports As Integer)
        ' Generate bells
        GlobalVariables.generate_bells(bells)
        Debug.Assert(GlobalVariables.bells.Count = bells)

        ' Generate COMs
        GlobalVariables.generate_COM_ports(ports)
        Debug.Assert(GlobalVariables.COM_ports.Count = ports)
    End Sub

End Module
#End If
