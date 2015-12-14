Module FormStatsFunctions

    Const STATS_FONT_SIZE As Integer = 14
    Const STATS_FIELD_WIDTH As Integer = 200
    Const STATS_KEY_HEIGHT As Integer = 40
    Const STATS_VALUE_HEIGHT As Integer = 40
    Const STATS_PAIR_GAP As Integer = 15
    Const STATS_FIELD_HEIGHT As Integer = STATS_VALUE_HEIGHT + STATS_KEY_HEIGHT + STATS_PAIR_GAP
    Const STATS_FIELD_GAP As Integer = 30
    Private STATS_KEY_SIZE As New Size(STATS_FIELD_WIDTH, STATS_KEY_HEIGHT)
    Private STATS_VALUE_SIZE As New Size(STATS_FIELD_WIDTH, STATS_VALUE_HEIGHT)
    Private STATS_VALUE_OFFSET As New Point(0, STATS_KEY_HEIGHT + STATS_PAIR_GAP)

    Public Sub generate_frmStats(parent As Form)
        Dim frm As New Form
        Dim button As New Button
        Dim x As Integer = 0
        Dim y As Integer = 0

        Statistics.reset_stats_fields()

        add_key_value_labels(Statistics.changes_key, Statistics.changes_value, "Changes", frm, x, y)
        x += 1

        add_key_value_labels(Statistics.leads_key, Statistics.leads_value, "Leads", frm, x, y)
        x += 1

        add_key_value_labels(Statistics.time_key, Statistics.time_value, "Time", frm, x, y)
        x = 0
        y += 1

        add_key_value_labels(Statistics.peal_speed_key, Statistics.peal_speed_value, "Peal Speed", frm, x, y)
        x += 1

        add_key_value_labels(Statistics.lead_end_row_key, Statistics.lead_end_row_value, "Lead End Row", frm, x, y)
        x = 0
        y += 1

        add_key_value_labels(Statistics.last_course_peal_speed_key, Statistics.last_course_peal_speed_value,
                             "Last Course Peal Speed", frm, x, y)
        x += 1

        add_key_value_labels(Statistics.last_course_time_key, Statistics.last_course_time_value, "Last Course Time", frm, x, y)

        button.Size = New Size(STATS_FIELD_WIDTH, STATS_FIELD_HEIGHT)
        button.Location = coordinate(1, 3)
        button.Text = "Close"
        AddHandler button.Click, AddressOf close_parent_form

        frm.Controls.Add(button)


        frm.Text = "Statistics"
        frm.Name = "frmStats"
        frm.Font = New Font(DEFAULT_FONT.OriginalFontName, STATS_FONT_SIZE, DEFAULT_FONT.Style)
        frm.ClientSize = Point.op_Explicit(coordinate(3, 4))
        parent.AddOwnedForm(frm)
        AddHandler frm.FormClosing, AddressOf dispose_of_form
        frm.Show()

    End Sub

    ' Function to generate a key/value pair of labels
    Private Sub add_key_value_labels(key As Label, value As Label, name As String, frm As Form, x As Integer, y As Integer)
        key.Text = name + ":"
        key.Size = STATS_KEY_SIZE
        key.Location = coordinate(x, y)
        value.Text = ""
        value.Size = STATS_VALUE_SIZE
        value.Location = coordinate(x, y) + STATS_VALUE_OFFSET
        value.Name = "stats_" + name.ToLower().Replace(" ", "_")

        frm.Controls.Add(key)
        frm.Controls.Add(value)
    End Sub

    ' Function to return the coordinate of a certain row and column
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(STATS_FIELD_GAP + x * (STATS_FIELD_GAP + STATS_FIELD_WIDTH),
                        STATS_FIELD_GAP + y * (STATS_FIELD_GAP + STATS_FIELD_HEIGHT))
    End Function

End Module
