Module FormComFunctions

    Const COM_FIELD_GAP As Integer = 20
    Const COM_FIELD_HEIGHT As Integer = 20
    Const COM_FIELD_WIDTH As Integer = 140
    Const COM_LABEL_WIDTH As Integer = 50
    Const COM_COLUMN_1 As Integer = COM_FIELD_GAP
    Const COM_COLUMN_2 As Integer = 2 * COM_FIELD_GAP + COM_FIELD_WIDTH
    Const COM_COLUMN_3 As Integer = 3 * COM_FIELD_GAP + COM_FIELD_WIDTH + COM_LABEL_WIDTH
    Const COM_MIN_HEIGHT As Integer = 2 * COM_FIELD_GAP + 5 * COM_FIELD_HEIGHT

    ' Simple function to return the y coordinate (in pixels) of the row index ii
    Private Function row_height(ii As Integer) As Integer
        Return (COM_FIELD_GAP + ii * (COM_FIELD_GAP + COM_FIELD_HEIGHT))
    End Function

    ' Function to initialise the COM ports
    ' Create a form showing a list of the available COM ports,
    ' and then have text boxes for input for as many COM ports as they selected.
    Public Sub generate_frmCOMs(parent As Form)
        Dim frm As New Form
        Dim list_box As New ListBox
        Dim label As New Label
        Dim button As New Button
        Dim form_height As Integer

        ' ii counts how many rows of our grid we have used
        Dim ii As Integer

        list_box.Size = New Size(COM_FIELD_WIDTH, 5 * COM_FIELD_HEIGHT)
        list_box.Location = New Point(COM_COLUMN_1, COM_FIELD_GAP)
        For Each port_name In IO.Ports.SerialPort.GetPortNames()
            list_box.Items.Add(port_name)
        Next

        label.Text = "Please choose a port on the left for each connection shown below"
        label.Size = New Size(COM_FIELD_WIDTH + COM_FIELD_GAP + COM_LABEL_WIDTH, 2 * COM_FIELD_HEIGHT)
        label.Location = New Point(COM_COLUMN_2, COM_FIELD_GAP)

        ii = 1
        For Each port In GlobalVariables.COM_ports
            generate_COM_field(frm, row_height(ii), ii)
            ii += 1
        Next

        button.Text = "Done"
        button.Size = New Size(COM_FIELD_WIDTH, COM_FIELD_HEIGHT)
        button.Location = New Point(COM_COLUMN_3, row_height(ii))
        AddHandler button.Click, AddressOf frmCOM_done

        ii += 1

        frm.Controls.Add(list_box)
        frm.Controls.Add(label)
        frm.Controls.Add(button)

        form_height = row_height(ii)
        If form_height < COM_MIN_HEIGHT Then
            form_height = COM_MIN_HEIGHT
        End If

        frm.ClientSize = New Size(COM_COLUMN_3 + COM_FIELD_WIDTH + COM_FIELD_GAP, form_height)
        parent.AddOwnedForm(frm)
        frm.Text = "COM port settings"
        frm.Name = "frmCOM"
        frm.Font = DEFAULT_FONT
        AddHandler frm.FormClosing, AddressOf dispose_of_form
        frm.Show()

    End Sub

    ' Function that is called when the button on the COM form is pressed.
    Private Sub frmCOM_done(sender As Button, e As EventArgs)
        Dim frm As Form

        ' Get the parent form of the button that caused this event.
        frm = sender.Parent

        ' Load the COM ports.
        load_COM_port_names(frm)

        ' We have now (hopefully) filled each port with a COM port name.
        ' Try and connect all the COM ports.
        If GlobalVariables.connect_COM_ports() Then
            ' We have successfully opened all the COM ports.
            ' close this form and return to the main form.
            GlobalVariables.COM_ports_configured = True
            frm.Close()
        Else
            ' We haven't opened all the ports successfully.
            ' Post an error message that this failed.
            MessageBox.Show("Failed to open COM ports.", "Error")
        End If
    End Sub

    ' Function to find load all the input values into the COM port list.
    Private Sub load_COM_port_names(frm As Form)
        Dim ii As Integer = 0

        ' Look at every item on this form
        For Each ctrl As Control In frm.Controls

            ' If the name of the item starts with 'port' then match
            ' This means we can't name any other names on this form starting with port.
            If ctrl.Name.StartsWith("port") Then
                ' We have found a port box, add it's text to a port in the list, as long as there is text in there.
                If Not ctrl.Text.Equals("") Then
                    GlobalVariables.COM_ports(ii).PortName = ctrl.Text
                    ii += 1
                End If
            End If
        Next


    End Sub

    ' Function to put in a text box in the specified location for creating frmCOMs
    Private Sub generate_COM_field(frm As Form, coordinate As Integer, index As Integer)
        Dim text_box As New TextBox
        Dim label As New Label

        text_box.Name = "port" + index.ToString() + "_txt"
        text_box.Size = New Size(COM_FIELD_WIDTH, COM_FIELD_HEIGHT)
        text_box.Location = New Point(COM_COLUMN_3, coordinate)

        label.Text = "Port " + index.ToString() + ": "
        label.Location = New Point(COM_COLUMN_2, coordinate)
        label.Size = New Size(COM_LABEL_WIDTH, COM_FIELD_HEIGHT)

        frm.Controls.Add(text_box)
        frm.Controls.Add(label)
    End Sub

End Module
