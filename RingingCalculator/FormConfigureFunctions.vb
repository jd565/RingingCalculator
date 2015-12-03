Module FormConfigureFunctions

    'Function to create a form to show while waiting for the config to be set
    Public Sub generate_frmConfigure(parent As Form)
        Dim waiting_form As New Form
        Dim label As New Label
        Dim cancel As New Button

        label.Text = "Please ring bell or press switch"
        label.Size = New Size(70, 40)
        label.Location = New Point(35, 30)

        cancel.Text = "Cancel"
        cancel.Size = New Size(70, 40)
        cancel.Location = New Point(35, 70)
        AddHandler cancel.Click, AddressOf close_parent_form

        waiting_form.Text = "Configure"
        waiting_form.Name = "frmConfigure"
        waiting_form.Font = DEFAULT_FONT
        waiting_form.ClientSize = New Size(140, 140)
        waiting_form.Controls.Add(cancel)
        waiting_form.Controls.Add(label)
        parent.AddOwnedForm(waiting_form)
        AddHandler waiting_form.FormClosing, AddressOf cancel_config_waiting_form

        parent.Hide()
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

        GlobalVariables.switch.can_be_configured = False

        dispose_of_form(frm, e)
    End Sub

    ' Function to be called when a configure event is hit on a serial port.
    ' We expect the waiting form to be open, but check this.
    ' If the waiting form is open, then close it, as the bell has been configured.
    Public Sub item_has_been_configured()
        Dim frm As Form

        frm = find_form("frmConfigure")
        If frm.InvokeRequired Then
            frm.Invoke(New MethodInvoker(AddressOf item_has_been_configured))
        Else
            frm.Close()
        End If
    End Sub

End Module
