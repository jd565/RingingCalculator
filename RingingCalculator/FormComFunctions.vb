Partial Class frmCom
    Inherits Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    Const COM_FIELD_GAP As Integer = 20
    Const COM_FIELD_HEIGHT As Integer = 20
    Const COM_FIELD_WIDTH As Integer = 140
    Const COM_LABEL_WIDTH As Integer = 50
    Const COM_LIST_HEIGHT As Integer = 5 * COM_FIELD_HEIGHT
    Const COM_TITLE_HEIGHT As Integer = 2 * COM_FIELD_HEIGHT
    Const COM_TITLE_WIDTH As Integer = COM_FIELD_WIDTH + COM_FIELD_GAP + COM_LABEL_WIDTH

    Private COM_LIST_SIZE As New Size(COM_FIELD_WIDTH, COM_LIST_HEIGHT)
    Private COM_LABEL_SIZE As New Size(COM_LABEL_WIDTH, COM_FIELD_HEIGHT)
    Private COM_FIELD_SIZE As New Size(COM_FIELD_WIDTH, COM_FIELD_HEIGHT)
    Private COM_TITLE_SIZE As New Size(COM_TITLE_WIDTH, COM_TITLE_HEIGHT)

    Private list_box As ListBox
    Private label As Label
    Private button As Button
    Private form_height As Integer
    Private parent_frm As Form

    ' Function to initialise the COM ports
    ' Create a form showing a list of the available COM ports,
    ' and then have text boxes for input for as many COM ports as they selected.
    Public Sub generate(parent As Form)
        Me.list_box = New ListBox
        Me.label = New Label
        Me.button = New Button

        Dim ii As Integer

        Me.list_box.Size = COM_LIST_SIZE
        Me.list_box.Location = Me.coordinate(0, 1)
        For Each port_name In IO.Ports.SerialPort.GetPortNames()
            Me.list_box.Items.Add(port_name)
        Next

        Me.label.Text = "Please choose a port on the left for each connection shown below"
        Me.label.Size = COM_TITLE_SIZE
        Me.label.Location = Me.coordinate(0, 0)

        ii = 1
        For Each port In GlobalVariables.COM_ports
            Me.generate_COM_field(port, ii, ii + 1)
            ii += 1
        Next

        Me.button.Text = "Done"
        Me.button.Size = New Size(COM_FIELD_WIDTH, COM_FIELD_HEIGHT)
        Me.button.Location = coordinate(2, ii)
        AddHandler Me.button.Click, AddressOf Me.frmCOM_done

        ii += 1

        Me.Controls.Add(Me.list_box)
        Me.Controls.Add(Me.label)
        Me.Controls.Add(Me.button)

        Me.ClientSize = New Size(coordinate(3, Math.Max(ii, 5)))
        parent.AddOwnedForm(Me)
        Me.parent_frm = parent
        Me.Text = "COM port settings"
        Me.Name = "frmCOM"
        Me.Font = DEFAULT_FONT
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        Me.Show()
        parent.Hide()

    End Sub

    ' Function that is called when the button on the COM form is pressed.
    Private Sub frmCOM_done(sender As Button, e As EventArgs)

        ' Load the COM ports.
        Me.load_COM_port_names()

        ' We have now (hopefully) filled each port with a COM port name.
        ' Try and connect all the COM ports.
        If GlobalVariables.connect_COM_ports() Then
            ' We have successfully opened all the COM ports.
            ' close this form and show the bell screen.
            If Me.parent_frm.GetType = GetType(frmNewConf) Then
                Dim frm As frmNewConf = Me.parent_frm
                frm.generate_frm_bells()
            End If
            Me.Close()
        Else
                ' We haven't opened all the ports successfully.
                ' Post an error message that this failed.
                MessageBox.Show("Failed to open COM ports.", "Error")
        End If
    End Sub

    ' Function to return the coordinate of the point on the grid
    Private Function coordinate(x As Integer, y As Integer) As Point
        Dim ii As Integer
        Dim jj As Integer

        If x <= 1 Then
            ii = COM_FIELD_GAP + x * (COM_FIELD_GAP + COM_FIELD_WIDTH)
        Else
            ii = COM_FIELD_GAP + x * COM_FIELD_GAP + (x - 1) * COM_FIELD_WIDTH + COM_LABEL_WIDTH
        End If

        If x = 0 Then
            If y <= 1 Then
                jj = COM_FIELD_GAP + y * (COM_FIELD_GAP + COM_TITLE_HEIGHT)
            Else
                jj = 3 * COM_FIELD_GAP + COM_TITLE_HEIGHT + COM_LIST_HEIGHT + (y - 2) * (COM_FIELD_GAP + COM_FIELD_HEIGHT)
            End If
        Else
            If y > 0 Then
                jj = 2 * COM_FIELD_GAP + COM_FIELD_HEIGHT + (y - 1) * (COM_FIELD_GAP + COM_FIELD_HEIGHT)
            Else
                jj = COM_FIELD_GAP
            End If
        End If

        Return New Point(ii, jj)
    End Function

    ' Function to load all the input values into the COM port list.
    Private Sub load_COM_port_names()
        For Each port In GlobalVariables.COM_ports
            If Not port.txt_field.Text.Equals("") Then
                port.PortName = port.txt_field.Text
            End If
        Next
    End Sub

    ' Function to put in a text box in the specified location for creating frmCOMs
    Private Sub generate_COM_field(port As COMPort, index As Integer, c As Integer)
        port.new_fields()

        port.txt_field.Name = "port" + index.ToString() + "_txt"
        port.txt_field.Size = COM_FIELD_SIZE
        port.txt_field.Location = coordinate(1, c)

        port.label.Text = "Port " + index.ToString() + ": "
        port.label.Location = coordinate(2, c)
        port.label.Size = COM_LABEL_SIZE

        Me.Controls.Add(port.txt_field)
        Me.Controls.Add(port.label)
    End Sub

End Class
