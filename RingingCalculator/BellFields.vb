Public Class BellFields
    Public button As New Button
    Public handstroke_delay As New NumericUpDown
    Public handstroke_label As New Label
    Public backstroke_delay As New NumericUpDown
    Public backstroke_label As New Label
    Private parent_reference As Bell

    ' A shape that is displayed on forms to show the state of the bell
    Public blob As New PowerPacks.OvalShape

    ' Canvas to hold the shape
    Public canvas As New PowerPacks.ShapeContainer

    Public ReadOnly Property parent As Bell
        Get
            Return Me.parent_reference
        End Get
    End Property

    ' When an object of this class is initialized then assign all the names to
    ' the controls.
    Public Sub New(ByRef bell As Bell)
        Me.parent_reference = bell

        handstroke_delay.Maximum = 1000
        handstroke_delay.Name = bell.name + "_hnd"
        handstroke_delay.Value = bell.handstroke_delay
        AddHandler handstroke_delay.ValueChanged, AddressOf handstroke_delay_changed_wrapper

        handstroke_label.Text = "Handstroke delay:"

        backstroke_delay.Maximum = 1000
        backstroke_delay.Name = bell.name + "_bck"
        backstroke_delay.Value = bell.backstroke_delay
        AddHandler backstroke_delay.ValueChanged, AddressOf backstroke_delay_changed_wrapper

        backstroke_label.Text = "Backstroke delay:"

        button.Name = bell.name + "_button"
        button.Text = "Configure bell on serial port"
        AddHandler button.Click, AddressOf configure_button_pressed

        blob = New PowerPacks.OvalShape
        blob.BackStyle = PowerPacks.BackStyle.Opaque
        blob.BackColor = Color.Gray
        blob.BorderWidth = 2
        blob.BorderColor = Color.Black
        blob.Parent = canvas
    End Sub

    ' Function to handle when the configure button is pressed.
    ' This could be triggered by a bell or the switch
    Private Sub configure_button_pressed(btn As Button, e As EventArgs)
        Dim frm As New frmConfigure(Me.parent)
        Me.parent.can_be_configured = True

        ' We only want 1 configuration happening at a time, so pop out another form to stop this happening
        frm.generate(btn.Parent)
    End Sub

    Private Sub handstroke_delay_changed_wrapper(updown As NumericUpDown, e As EventArgs)
        Me.parent.handstroke_delay = updown.Value
    End Sub

    Private Sub backstroke_delay_changed_wrapper(updown As NumericUpDown, e As EventArgs)
        Me.parent.backstroke_delay = updown.Value
    End Sub

End Class
