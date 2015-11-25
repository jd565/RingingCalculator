Module FormBellsFunctions

    Const BELL_FIELD_GAP As Integer = 35
    Const BELL_FIELD_HEIGHT As Integer = 50
    Const BELL_UPDOWN_WIDTH As Integer = 40
    Const BELL_LABEL_WIDTH As Integer = 65
    Const BELL_FIELD_WIDTH As Integer = BELL_LABEL_WIDTH + BELL_UPDOWN_WIDTH
    Const BELL_LIGHT_DIAMETER As Integer = 65
    Const BELL_ROW_1 As Integer = BELL_FIELD_GAP
    Const BELL_ROW_2 As Integer = BELL_ROW_1 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_3 As Integer = BELL_ROW_2 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_4 As Integer = BELL_ROW_3 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_5 As Integer = BELL_ROW_4 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_FORM_HEIGHT As Integer = BELL_ROW_5 + BELL_FIELD_GAP + BELL_LIGHT_DIAMETER

    ' Function to return the x coordinate (in pixels) of the column index ii
    Private Function column_width(ii As Integer) As Integer
        Return BELL_FIELD_GAP + ii * (BELL_FIELD_GAP + BELL_FIELD_WIDTH)
    End Function

    ' Function to generate the bells form with the required number of bells
    ' We decide a required width per bell, and set the overall form width.
    ' We then generate the required boxes for each bell.
    Public Sub generate_frmBells(parent As Form)
        Dim frm As New Form
        Dim ii As Integer = 0
        Dim debounce_time As New NumericUpDown
        Dim debounce_label As New Label

        GlobalVariables.switch.new_button()
        GlobalVariables.bells.ForEach(Sub(bell) bell.new_fields())

        GlobalVariables.switch.button.Size = New Size(BELL_FIELD_WIDTH, BELL_FIELD_HEIGHT)
        GlobalVariables.switch.button.Location = New Point(column_width(ii), BELL_ROW_1)
        ii += 1

        debounce_time.Name = "debounce_time_updown"
        debounce_time.Value = 25
        debounce_time.Size = New Size(BELL_UPDOWN_WIDTH, BELL_FIELD_HEIGHT)
        debounce_time.Location = New Point(column_width(ii) + BELL_FIELD_WIDTH - BELL_UPDOWN_WIDTH, BELL_ROW_1)
        AddHandler debounce_time.ValueChanged, AddressOf debounce_time_changed

        debounce_label.Name = "debounce_time_label"
        debounce_label.Text = "Debounce delay:"
        debounce_label.Size = New Size(BELL_LABEL_WIDTH, BELL_FIELD_HEIGHT)
        debounce_label.Location = New Point(column_width(ii), BELL_ROW_1)

        ' Reset the column width as the bells go on the row below these configs
        ii = 0
        For Each bell In GlobalVariables.bells
            add_bell_field(column_width(ii), bell, frm)
            ii += 1
        Next

        frm.Controls.Add(GlobalVariables.switch.button)
        frm.Controls.Add(debounce_time)
        frm.Controls.Add(debounce_label)

        frm.Text = "Ringing Simulator"
        frm.Name = "frmBells"
        frm.Font = DEFAULT_FONT
        frm.ClientSize = New Size(column_width(ii), BELL_FORM_HEIGHT)
        parent.AddOwnedForm(frm)
        AddHandler frm.FormClosing, AddressOf dispose_of_form
        frm.Show()

    End Sub

    ' Function to add one bell. The bell fields will be placed at the coordinate specified,
    ' and named using bell_name
    Private Sub add_bell_field(coordinate As Integer, bell As Bell, form As Form)

        bell.fields.handstroke_delay.Size = New Size(BELL_UPDOWN_WIDTH, BELL_FIELD_HEIGHT)
        bell.fields.handstroke_delay.Location = New Point(coordinate + BELL_FIELD_WIDTH - BELL_UPDOWN_WIDTH, BELL_ROW_2)

        bell.fields.handstroke_label.Size = New Size(BELL_LABEL_WIDTH, BELL_FIELD_HEIGHT)
        bell.fields.handstroke_label.Location = New Point(coordinate, BELL_ROW_2)

        bell.fields.backstroke_delay.Size = New Size(BELL_UPDOWN_WIDTH, BELL_FIELD_HEIGHT)
        bell.fields.backstroke_delay.Location = New Point(coordinate + BELL_FIELD_WIDTH - BELL_UPDOWN_WIDTH, BELL_ROW_3)

        bell.fields.backstroke_label.Size = New Size(BELL_LABEL_WIDTH, BELL_FIELD_HEIGHT)
        bell.fields.backstroke_label.Location = New Point(coordinate, BELL_ROW_3)

        bell.fields.button.Size = New Size(BELL_FIELD_WIDTH, BELL_FIELD_HEIGHT)
        bell.fields.button.Location = New Point(coordinate, BELL_ROW_4)

        bell.fields.blob.Size = New Size(BELL_LIGHT_DIAMETER, BELL_LIGHT_DIAMETER)
        bell.fields.blob.Location = New Point(coordinate + (BELL_FIELD_WIDTH - BELL_LIGHT_DIAMETER) / 2, BELL_ROW_5)

        form.Controls.Add(bell.fields.button)
        form.Controls.Add(bell.fields.handstroke_delay)
        form.Controls.Add(bell.fields.handstroke_label)
        form.Controls.Add(bell.fields.backstroke_delay)
        form.Controls.Add(bell.fields.backstroke_label)
        form.Controls.Add(bell.fields.canvas)

    End Sub

    Private Sub debounce_time_changed(debounce_time As NumericUpDown, e As EventArgs)
        GlobalVariables.debounce_time_changed(debounce_time.Value)
    End Sub

End Module
