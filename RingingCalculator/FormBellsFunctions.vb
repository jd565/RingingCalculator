Module FormBellsFunctions

    Const BELL_FIELD_GAP As Integer = 20
    Const BELL_FIELD_WIDTH As Integer = 70

    ' Function to return the x coordinate (in pixels) of the column index ii
    Private Function column_width(ii As Integer) As Integer
        Return BELL_FIELD_GAP + ii * (BELL_FIELD_GAP + BELL_FIELD_WIDTH)
    End Function

    ' Function to generate the bells form with the required number of bells
    ' We decide a required width per bell, and set the overall form width.
    ' We then generate the required boxes for each bell.
    Public Sub generate_frmbells()
        Dim frm As New Form
        Dim ii As Integer

        ii = 0
        For Each bell In GlobalVariables.bells
            add_bell_field(column_width(ii), bell, frm)
            ii += 1
        Next

        frm.Text = "Ringing Simulator"
        frm.Name = "frmBells"
        frm.ClientSize = New Size(column_width(ii), 300)
        frmMain.AddOwnedForm(frm)
        AddHandler frm.FormClosing, AddressOf dispose_of_form
        frm.Show()

    End Sub

    ' Function to add one bell. The bell fields will be placed at the coordinate specified,
    ' and named using bell_name
    Public Sub add_bell_field(coordinate As Integer, bell As Bell, form As Form)
        bell.button = New Button
        bell.button.Name = bell.name + "_button"
        bell.button.Text = "Configure bell on serial port"
        bell.button.Size = New Size(BELL_FIELD_WIDTH, 40)
        bell.button.Location = New Point(coordinate, BELL_FIELD_GAP)
        AddHandler bell.button.Click, AddressOf configure_bell_button_pressed

        form.Controls.Add(bell.button)

    End Sub

    ' Function to handle when the configure bell button is pressed.
    ' This should find the bell object that owns the button and set
    ' it to expect being configured.
    Public Sub configure_bell_button_pressed(btn As Button, e As EventArgs)
        For Each bell In GlobalVariables.bells
            If btn.Equals(bell.button) Then

                ' We want to set the bell to be configured.
                bell.can_be_configured = True

                ' We only want 1 configuration happening at a time, so pop out another form to stop this happening
                create_config_waiting_form(btn.Parent)
                Exit For
            End If
        Next
    End Sub

    'Function to create a form to show while waiting for the config to be set
    Private Sub create_config_waiting_form(frm As Form)
        Dim waiting_form As New Form
        Dim label As New Label
        Dim cancel As New Button

        label.Text = "Please ring a bell."
        label.Size = New Size(70, 40)
        label.Location = New Point(35, 30)

        cancel.Text = "Cancel"
        cancel.Size = New Size(70, 40)
        cancel.Location = New Point(35, 70)
        AddHandler cancel.Click, AddressOf close_parent_form

        waiting_form.Text = "Configure bell port."
        waiting_form.Name = "frmConfigure"
        waiting_form.ClientSize = New Size(140, 140)
        waiting_form.Controls.Add(cancel)
        waiting_form.Controls.Add(label)
        frm.AddOwnedForm(waiting_form)
        AddHandler waiting_form.FormClosing, AddressOf cancel_config_waiting_form

        frm.Hide()
        waiting_form.Show()
    End Sub

    ' Function to cancel the config waiting form and show the previous form
    ' We first set all bells to be not configurable,
    ' then delete this form.
    Private Sub cancel_config_waiting_form(frm As Form, e As EventArgs)
        For Each bell In GlobalVariables.bells
            If bell.can_be_configured Then
                bell.can_be_configured = False
            End If
        Next

        dispose_of_form(frm, e)
    End Sub

    ' Function to be called when a configure event is hit on a serial port.
    ' We expect the waiting form to be open, but check this.
    ' If the waiting form is open, then close it, as the bell has been configured.
    Public Sub bell_has_been_configured()

        ' Loop through all open forms, in case the waiting form has been sent to the background.
        For Each frm As Form In My.Application.OpenForms
            If frm.Name.Equals("frmConfigure") Then
                Console.WriteLine("Found form to close")
                frm.Close()
                Exit For
            End If
        Next
    End Sub

End Module
