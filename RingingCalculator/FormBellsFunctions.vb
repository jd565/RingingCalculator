Module FormBellsFunctions

    Const BELL_FIELD_GAP As Integer = 40
    Const BELL_LABEL_HEIGHT As Integer = 25
    Const BELL_UPDOWN_HEIGHT As Integer = 25
    Const BELL_FIELD_HEIGHT As Integer = BELL_LABEL_HEIGHT + BELL_UPDOWN_HEIGHT
    Const BELL_FIELD_WIDTH As Integer = 100
    Const BELL_LIGHT_DIAMETER As Integer = 50
    Private BELL_LABEL_SIZE As New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
    Private BELL_UPDOWN_SIZE As New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
    Private BELL_UPDOWN_OFFSET As New Point(0, BELL_LABEL_HEIGHT)
    Private BELL_GENERAL_SIZE As New Size(BELL_FIELD_WIDTH, BELL_FIELD_HEIGHT)
    Private BELL_LIGHT_SIZE As New Size(Math.Min(Math.Min(BELL_LIGHT_DIAMETER, BELL_FIELD_HEIGHT), BELL_FIELD_WIDTH),
                                        Math.Min(Math.Min(BELL_LIGHT_DIAMETER, BELL_FIELD_HEIGHT), BELL_FIELD_WIDTH))
    Private BELL_LIGHT_OFFSET As New Point((BELL_FIELD_WIDTH - BELL_LIGHT_SIZE.Width) / 2,
                                           (BELL_FIELD_HEIGHT - BELL_LIGHT_SIZE.Height) / 2)
    Private debounce_time As New NumericUpDown
    Private debounce_label As New Label
    Private changes_per_lead As New TextBox
    Private lbl_changes_per_lead As New Label
    Private leads_per_course As New TextBox
    Private lbl_leads_per_course As New Label
    Private changes_per_Peal As New TextBox
    Private lbl_changes_per_Peal As New Label
    Private save_file As New SaveFileDialog
    Private save As New Button
    Private frm As New Form

    ' Function to return the coordinate of a point on the grid
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(BELL_FIELD_GAP + x * (BELL_FIELD_GAP + BELL_FIELD_WIDTH),
                         BELL_FIELD_GAP + y * (BELL_FIELD_GAP + BELL_FIELD_HEIGHT))
    End Function

    ' Function to generate the bells form with the required number of bells
    ' We decide a required width per bell, and set the overall form width.
    ' We then generate the required boxes for each bell.
    Public Sub generate_frmBells(parent As Form)
        Dim ii As Integer = 0

        ' Switch
        GlobalVariables.switch.new_button()
        GlobalVariables.switch.button.Size = BELL_GENERAL_SIZE
        GlobalVariables.switch.button.Location = coordinate(0, 0)

        ' Debounce time
        debounce_label.Name = "debounce_time_label"
        debounce_label.Text = "Debounce delay:"
        debounce_label.Size = BELL_LABEL_SIZE
        debounce_label.Location = coordinate(1, 0)
        debounce_time.Name = "debounce_time_updown"
        debounce_time.Value = 25
        debounce_time.Size = BELL_UPDOWN_SIZE
        debounce_time.Location = coordinate(1, 0) + BELL_UPDOWN_OFFSET
        AddHandler debounce_time.ValueChanged, AddressOf debounce_time_changed

        ' Changes per lead
        lbl_changes_per_lead.Text = "Changes per lead:"
        lbl_changes_per_lead.Size = BELL_LABEL_SIZE
        lbl_changes_per_lead.Location = coordinate(2, 0)
        changes_per_lead.Text = GlobalVariables.changes_per_lead.ToString
        changes_per_lead.Size = BELL_UPDOWN_SIZE
        changes_per_lead.Location = coordinate(2, 0) + BELL_UPDOWN_OFFSET
        AddHandler changes_per_lead.TextChanged, AddressOf changes_per_lead_changed

        ' leads per course
        lbl_leads_per_course.Text = "Leads per course:"
        lbl_leads_per_course.Size = BELL_LABEL_SIZE
        lbl_leads_per_course.Location = coordinate(3, 0)
        leads_per_course.Text = GlobalVariables.leads_per_course.ToString
        leads_per_course.Size = BELL_UPDOWN_SIZE
        leads_per_course.Location = coordinate(3, 0) + BELL_UPDOWN_OFFSET
        AddHandler leads_per_course.TextChanged, AddressOf leads_per_course_changed

        ' changes per peal
        lbl_changes_per_Peal.Text = "Peal length:"
        lbl_changes_per_Peal.Size = BELL_LABEL_SIZE
        lbl_changes_per_Peal.Location = coordinate(4, 0)
        changes_per_Peal.Text = GlobalVariables.changes_per_peal.ToString
        changes_per_Peal.Size = BELL_UPDOWN_SIZE
        changes_per_Peal.Location = coordinate(4, 0) + BELL_UPDOWN_OFFSET
        AddHandler changes_per_Peal.TextChanged, AddressOf changes_per_peal_changed

        ' Bells
        ' Add new fields to the bell to stop errors where fields are on more than 1 form
        GlobalVariables.bells.ForEach(Sub(bell) bell.new_fields())
        ' Reset the column width as the bells go on the row below these configs
        ii = 0
        For Each bell In GlobalVariables.bells
            add_bell_field(ii, bell, frm)
            ii += 1
        Next

        ' Config file saving
        save.Text = "Save configuration"
        save.Size = BELL_GENERAL_SIZE
        save.Location = coordinate(0, 5)
        AddHandler save.Click, AddressOf save_button_pressed
        save_file.FileName = "ringingcalculator.conf"
        save_file.Filter = "Conf files|*.conf|All files|*.*"
        save_file.Title = "Configuration File"
        save_file.DefaultExt = "conf"

        frm.Controls.Add(GlobalVariables.switch.button)
        frm.Controls.Add(debounce_time)
        frm.Controls.Add(debounce_label)
        frm.Controls.Add(lbl_changes_per_lead)
        frm.Controls.Add(changes_per_lead)
        frm.Controls.Add(lbl_changes_per_Peal)
        frm.Controls.Add(changes_per_Peal)
        frm.Controls.Add(lbl_leads_per_course)
        frm.Controls.Add(leads_per_course)
        frm.Controls.Add(save)

        frm.Text = "Ringing Simulator"
        frm.Name = "frmBells"
        frm.Font = DEFAULT_FONT
        frm.ClientSize = New Size(coordinate(ii, 6))
        parent.AddOwnedForm(frm)
        AddHandler frm.FormClosing, AddressOf dispose_of_form
        frm.Show()

    End Sub

    ' Function to add one bell. The bell fields will be placed at the coordinate specified,
    ' and named using bell_name
    Private Sub add_bell_field(x As Integer, bell As Bell, form As Form)

        ' Hnadstroke delay
        bell.fields.handstroke_delay.Size = BELL_UPDOWN_SIZE
        bell.fields.handstroke_delay.Location = coordinate(x, 1) + BELL_UPDOWN_OFFSET
        bell.fields.handstroke_label.Size = BELL_LABEL_SIZE
        bell.fields.handstroke_label.Location = coordinate(x, 1)

        ' Backstroke delay
        bell.fields.backstroke_delay.Size = BELL_UPDOWN_SIZE
        bell.fields.backstroke_delay.Location = coordinate(x, 2) + BELL_UPDOWN_OFFSET
        bell.fields.backstroke_label.Size = BELL_LABEL_SIZE
        bell.fields.backstroke_label.Location = coordinate(x, 2)

        ' configure port pin button
        bell.fields.button.Size = BELL_GENERAL_SIZE
        bell.fields.button.Location = coordinate(x, 3)

        ' Light
        bell.fields.blob.Size = BELL_LIGHT_SIZE
        bell.fields.blob.Location = coordinate(x, 4) + BELL_LIGHT_OFFSET

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

    Private Sub save_button_pressed(obj As Object, e As EventArgs)
        get_file_to_save_to()
        save_config(save_file.FileName)
        MsgBox("Configuration saved to " & save_file.FileName & ".",, "File Saved")
    End Sub

    Private Sub get_file_to_save_to()
        save_file.ShowDialog()
    End Sub

End Module
