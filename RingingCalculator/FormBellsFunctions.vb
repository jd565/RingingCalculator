Module FormBellsFunctions

    Const BELL_FIELD_GAP As Integer = 40
    Const BELL_LABEL_HEIGHT As Integer = 25
    Const BELL_UPDOWN_HEIGHT As Integer = 25
    Const BELL_FIELD_HEIGHT As Integer = BELL_LABEL_HEIGHT + BELL_UPDOWN_HEIGHT
    Const BELL_FIELD_WIDTH As Integer = 100
    Const BELL_LIGHT_DIAMETER As Integer = 50
    Const BELL_ROW_1 As Integer = BELL_FIELD_GAP
    Const BELL_ROW_2 As Integer = BELL_ROW_1 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_3 As Integer = BELL_ROW_2 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_4 As Integer = BELL_ROW_3 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    Const BELL_ROW_5 As Integer = BELL_ROW_4 + BELL_FIELD_GAP + BELL_FIELD_HEIGHT
    const BELL_ROW_6 As Integer = BELL_ROW_5 + BELL_FIELD_GAP + BELL_LIGHT_DIAMETER
    Const BELL_FORM_HEIGHT As Integer = BELL_ROW_6

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
        Dim changes_per_lead As New TextBox
        Dim lbl_changes_per_lead As New Label
        Dim leads_per_course As New TextBox
        Dim lbl_leads_per_course As New Label
        Dim changes_per_Peal As New TextBox
        Dim lbl_changes_per_Peal As New Label


        lbl_changes_per_lead.Text = "Changes per lead:"
        lbl_changes_per_lead.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
        lbl_changes_per_lead.Location = New Point(column_width(2), BELL_ROW_1)

        changes_per_lead.Text = GlobalVariables.changes_per_lead.ToString
        changes_per_lead.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        changes_per_lead.Location = New Point(column_width(2), BELL_ROW_1 + BELL_LABEL_HEIGHT)
        AddHandler changes_per_lead.TextChanged, AddressOf changes_per_lead_changed

        lbl_leads_per_course.Text = "Leads per course:"
        lbl_leads_per_course.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
        lbl_leads_per_course.Location = New Point(column_width(3), BELL_ROW_1)

        leads_per_course.Text = GlobalVariables.leads_per_course.ToString
        leads_per_course.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        leads_per_course.Location = New Point(column_width(3), BELL_ROW_1 + BELL_LABEL_HEIGHT)
        AddHandler leads_per_course.TextChanged, AddressOf leads_per_course_changed

        lbl_changes_per_Peal.Text = "Peal length:"
        lbl_changes_per_Peal.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
        lbl_changes_per_Peal.Location = New Point(column_width(4), BELL_ROW_1)

        changes_per_Peal.Text = GlobalVariables.changes_per_peal.ToString
        changes_per_Peal.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        changes_per_Peal.Location = New Point(column_width(4), BELL_ROW_1 + BELL_LABEL_HEIGHT)
        AddHandler changes_per_Peal.TextChanged, AddressOf changes_per_peal_changed

        GlobalVariables.switch.new_button()
        GlobalVariables.bells.ForEach(Sub(bell) bell.new_fields())

        GlobalVariables.switch.button.Size = New Size(BELL_FIELD_WIDTH, BELL_FIELD_HEIGHT)
        GlobalVariables.switch.button.Location = New Point(column_width(ii), BELL_ROW_1)
        ii += 1

        debounce_time.Name = "debounce_time_updown"
        debounce_time.Value = 25
        debounce_time.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        debounce_time.Location = New Point(column_width(ii), BELL_ROW_1 + BELL_LABEL_HEIGHT)
        AddHandler debounce_time.ValueChanged, AddressOf debounce_time_changed

        debounce_label.Name = "debounce_time_label"
        debounce_label.Text = "Debounce delay:"
        debounce_label.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
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
        frm.Controls.Add(lbl_changes_per_lead)
        frm.Controls.Add(changes_per_lead)
        frm.Controls.Add(lbl_changes_per_Peal)
        frm.Controls.Add(changes_per_Peal)
        frm.Controls.Add(lbl_leads_per_course)
        frm.Controls.Add(leads_per_course)

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

        bell.fields.handstroke_delay.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        bell.fields.handstroke_delay.Location = New Point(coordinate, BELL_ROW_2 + BELL_LABEL_HEIGHT)

        bell.fields.handstroke_label.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
        bell.fields.handstroke_label.Location = New Point(coordinate, BELL_ROW_2)

        bell.fields.backstroke_delay.Size = New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
        bell.fields.backstroke_delay.Location = New Point(coordinate, BELL_ROW_3 + BELL_LABEL_HEIGHT)

        bell.fields.backstroke_label.Size = New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
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

    Private Sub changes_per_lead_changed(txt As textbox, e As eventargs)
        If GlobalVariables.switch.isRunning Then
            Console.WriteLine("Tried to change changes per lead, but we are running")
            txt.Text = GlobalVariables.changes_per_lead
        Else
            GlobalVariables.changes_per_lead = Val(txt.Text)
            GlobalVariables.update_changes_per_course()
        End If
    End Sub

    Private Sub leads_per_course_changed(txt As TextBox, e As EventArgs)
        If GlobalVariables.switch.isRunning Then
            Console.WriteLine("Tried to change leads per course, but we are running")
            txt.Text = GlobalVariables.leads_per_course
        Else
            GlobalVariables.leads_per_course = Val(txt.Text)
            GlobalVariables.update_changes_per_course()
        End If
    End Sub

    Private Sub changes_per_peal_changed(txt As TextBox, e As EventArgs)
        If GlobalVariables.switch.isRunning Then
            Console.WriteLine("Tried to change peal length, but we are running")
            txt.Text = GlobalVariables.changes_per_peal
        Else
            GlobalVariables.changes_per_peal = Val(txt.Text)
        End If
    End Sub

End Module
