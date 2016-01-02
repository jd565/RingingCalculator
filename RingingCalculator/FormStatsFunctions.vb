Partial Class frmStats
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

    Const STATS_FONT_SIZE As Integer = 11
    Const STATS_FIELD_WIDTH As Integer = 200
    Const STATS_KEY_HEIGHT As Integer = 40
    Const STATS_VALUE_HEIGHT As Integer = 40
    Const STATS_PAIR_GAP As Integer = 15
    Const STATS_FIELD_HEIGHT As Integer = STATS_VALUE_HEIGHT + STATS_KEY_HEIGHT + STATS_PAIR_GAP
    Const STATS_FIELD_GAP As Integer = 30
    Private STATS_KEY_SIZE As New Size(STATS_FIELD_WIDTH, STATS_KEY_HEIGHT)
    Private STATS_VALUE_SIZE As New Size(STATS_FIELD_WIDTH, STATS_VALUE_HEIGHT)
    Private STATS_VALUE_OFFSET As New Point(0, STATS_KEY_HEIGHT + STATS_PAIR_GAP)

    Friend button As Button

    Public Sub generate(parent As Form)
        Me.button = New Button
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim frm_lights As New frmLights

        ' We want to generate the lights to go alongside this form
        frm_lights.generate(Me)

        Statistics.reset_stats_fields()

        Me.add_key_value_labels(Statistics.changes_key, Statistics.changes_value, "Changes", x, y)
        x += 1

        Me.add_key_value_labels(Statistics.leads_key, Statistics.leads_value, "Leads", x, y)
        x += 1

        Me.add_key_value_labels(Statistics.time_key, Statistics.time_value, "Time", x, y)
        x = 0
        y += 1

        Me.add_key_value_labels(Statistics.peal_speed_key, Statistics.peal_speed_value, "Peal Speed", x, y)
        x += 1

        Me.add_key_value_labels(Statistics.lead_end_row_key, Statistics.lead_end_row_value, "Lead End Row", x, y)
        x = 0
        y += 1

        Me.add_key_value_labels(Statistics.last_course_peal_speed_key, Statistics.last_course_peal_speed_value,
                             "Last Course Peal Speed", x, y)
        x += 1

        Me.add_key_value_labels(Statistics.last_course_time_key, Statistics.last_course_time_value, "Last Course Time", x, y)

        x = 0
        y += 1

        Me.add_key_value_labels(Statistics.changes_per_minute_key, Statistics.changes_per_minute_value, "Changes per minute", x, y)

        x += 1

        Me.add_key_value_labels(Statistics.last_minute_changes_key, Statistics.last_minute_changes_value, "Changes in the last minute", x, y)

        x += 1
        y += 1

        Me.button.Size = New Size(STATS_FIELD_WIDTH, STATS_FIELD_HEIGHT)
        Me.button.Location = coordinate(1, y)
        Me.button.Text = "Close"
        AddHandler Me.button.Click, AddressOf close_parent_form

        y += 1

        Me.Controls.Add(button)


        Me.Text = "Statistics"
        Me.Name = "frmStats"
        Me.Font = New Font(DEFAULT_FONT.OriginalFontName, STATS_FONT_SIZE, DEFAULT_FONT.Style)
        Me.ClientSize = New Size(coordinate(3, y))
        parent.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form
        Me.Show()

    End Sub

    ' Function to generate a key/value pair of labels
    Private Sub add_key_value_labels(key As Label, value As Label, name As String, x As Integer, y As Integer)
        key.Text = name + ":"
        key.Size = STATS_KEY_SIZE
        key.Location = coordinate(x, y)
        value.Text = ""
        value.Size = STATS_VALUE_SIZE
        value.Location = coordinate(x, y) + STATS_VALUE_OFFSET
        value.Name = "stats_" + name.ToLower().Replace(" ", "_")

        Me.Controls.Add(key)
        Me.Controls.Add(value)
    End Sub

    ' Function to return the coordinate of a certain row and column
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(STATS_FIELD_GAP + x * (STATS_FIELD_GAP + STATS_FIELD_WIDTH),
                        STATS_FIELD_GAP + y * (STATS_FIELD_GAP + STATS_FIELD_HEIGHT))
    End Function

End Class
