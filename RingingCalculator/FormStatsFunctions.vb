﻿Partial Class frmStats
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

    Const STATS_MAX_COLUMNS As Integer = 5

    Private STATS_FONT_SIZE As Integer = 11
    Private STATS_FIELD_WIDTH As Integer = 200
    Private STATS_KEY_HEIGHT As Integer = 40
    Private STATS_VALUE_HEIGHT As Integer = 40
    Private STATS_PAIR_GAP As Integer = 15
    Private STATS_FIELD_HEIGHT As Integer = STATS_VALUE_HEIGHT + STATS_KEY_HEIGHT + STATS_PAIR_GAP
    Private STATS_FIELD_GAP As Integer = 30
    Private STATS_MENU_HEIGHT As Integer = 25
    Private STATS_KEY_SIZE As New Size(STATS_FIELD_WIDTH, STATS_KEY_HEIGHT)
    Private STATS_VALUE_SIZE As New Size(STATS_FIELD_WIDTH, STATS_VALUE_HEIGHT)
    Private STATS_VALUE_OFFSET As New Point(0, STATS_KEY_HEIGHT + STATS_PAIR_GAP)

    Private rows As Integer

    Friend btn_close As ToolStripButton
    Friend btn_view_stats As ToolStripButton
    Friend btn_stop As Button
    Friend main_menu As MenuStrip
    Friend kvl_lpc As KeyValueLabel
    Friend kvl_cpl As KeyValueLabel

    Public Sub New(frm As Form)
        If Not GlobalVariables.statistics_init Then
            Statistics.init()
        End If
        Me.generate(frm)
    End Sub

    Public Sub generate(parent As Form)
        Me.btn_close = New ToolStripButton
        Me.btn_stop = New Button
        Me.btn_view_stats = New ToolStripButton
        Me.main_menu = New MenuStrip
        Me.kvl_cpl = New KeyValueLabel(True, "Changes Per Lead")
        Me.kvl_lpc = New KeyValueLabel(True, "Leads per Course")
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim frm_lights As frmLights

        Statistics.reset_stats_fields()

        For Each kvp In Statistics.key_vals
            If kvp.Value.enabled Then
                Me.add_key_value_labels(kvp.Value)
            End If
        Next

        Me.add_key_value_labels(Me.kvl_lpc)
        Me.kvl_lpc.value.Text = GlobalVariables.leads_per_course.ToString
        Me.add_key_value_labels(Me.kvl_cpl)
        Me.kvl_cpl.value.Text = GlobalVariables.changes_per_lead.ToString

        Me.btn_close.Text = "Close"
        AddHandler Me.btn_close.Click, AddressOf Me.close_form

        Me.btn_view_stats.Text = "View performance"
        AddHandler Me.btn_view_stats.Click, AddressOf Me.view_perf_stats

        Me.btn_stop.Text = "Stop"
        AddHandler Me.btn_stop.Click, AddressOf Me.stop_rec

        Me.set_sizes()

        'Menu
        Me.main_menu.Items.Add(Me.btn_close)
        Me.main_menu.Items.Add(Me.btn_view_stats)
        Me.main_menu.Height = STATS_MENU_HEIGHT

        Me.Controls.Add(btn_stop)
        Me.Controls.Add(Me.main_menu)

        Me.Text = "Statistics"
        Me.Name = "frmStats"
        Me.ClientSize = New Size(coordinate(STATS_MAX_COLUMNS, Me.rows))
        parent.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form
        AddHandler Me.SizeChanged, AddressOf Me.size_changed

        frm_lights = New frmLights(Me, Me.Size.Height)
        Me.Activate()

        Me.Show()

        Me.Location = New Point(0, 0)

    End Sub

    ' Function to generate a key/value pair of labels
    Private Sub add_key_value_labels(kvl As KeyValueLabel)
        kvl.key.Text = kvl.name + ":"
        kvl.value.Text = ""
        kvl.value.Name = "stats_" + kvl.name.ToLower().Replace(" ", "_")

        Me.Controls.Add(kvl.key)
        Me.Controls.Add(kvl.value)
    End Sub

    Private Sub set_key_value_size_loc(kvl As KeyValueLabel, c As Point)
        kvl.key.Size = STATS_KEY_SIZE
        kvl.key.Location = c
        kvl.value.Size = STATS_VALUE_SIZE
        kvl.value.Location = c + STATS_VALUE_OFFSET
    End Sub

    ' Function to handle the size of the form changing
    ' This should resize everything on the form to be arranged nicely on the new size
    Private Sub size_changed(o As Object, e As EventArgs)
        ' First work out the new values for the height and width values
        ' There should be a number of rows as stored in me.rows and that many + 1 gaps.
        ' The gap should be 1/3rd the height of the row.
        Me.STATS_FIELD_HEIGHT = Me.ClientSize.Height / (Me.rows + (Me.rows + 1) / 3)
        Me.STATS_FIELD_GAP = Me.STATS_FIELD_HEIGHT / 3
        Me.STATS_KEY_HEIGHT = (Me.STATS_FIELD_HEIGHT - Me.STATS_PAIR_GAP) / 2
        Me.STATS_VALUE_HEIGHT = Me.STATS_KEY_HEIGHT

        ' The width is big enough to allow 5 'columns' and 6 gaps.
        Me.STATS_FIELD_WIDTH = (Me.ClientSize.Width - 6 * Me.STATS_FIELD_GAP) / 5

        Me.STATS_KEY_SIZE = New Size(STATS_FIELD_WIDTH, STATS_KEY_HEIGHT)
        Me.STATS_VALUE_SIZE = New Size(STATS_FIELD_WIDTH, STATS_VALUE_HEIGHT)
        Me.STATS_VALUE_OFFSET = New Point(0, STATS_KEY_HEIGHT + STATS_PAIR_GAP)

        ' Now set the sizes of everything on the form
        Me.set_sizes

    End Sub

    Private Sub set_sizes()
        Dim x As Integer = 0
        Dim y As Integer = 0

        For Each kvp In Statistics.key_vals
            If kvp.Value.enabled Then
                Me.set_key_value_size_loc(kvp.Value, coordinate(x, y))
                x += 1
                If x = STATS_MAX_COLUMNS Then
                    x = 0
                    y += 1
                End If
            End If
        Next

        If x <> 0 Then
            y += 1
        End If

        Me.set_key_value_size_loc(Me.kvl_cpl, coordinate(1, y))
        Me.set_key_value_size_loc(Me.kvl_lpc, coordinate(2, y))

        Me.btn_stop.Size = New Size(STATS_FIELD_WIDTH, STATS_FIELD_HEIGHT)
        Me.btn_stop.Location = coordinate(0, y)

        Me.set_max_font()

        Me.rows = y + 1
    End Sub

    ' Function to find the largest font we can use for the current size of form.
    Private Sub set_max_font()
        ' The longest string on the form is "Last Course Peal Speed", so check against this.
        Dim long_string As String = Statistics.key_vals("Last Course Peal Speed").name
        Dim font_ck As Font = DEFAULT_FONT

        ' Start by checking this font and going larger until we no longer fit on 1 line.
        While TextRenderer.MeasureText(long_string, font_ck).Width < STATS_FIELD_WIDTH
            Me.STATS_FONT_SIZE += 1
            font_ck = New Font(font_ck.Name, Me.STATS_FONT_SIZE)
        End While

        ' Then go smaller
        While TextRenderer.MeasureText(long_string, font_ck).Width > STATS_FIELD_WIDTH
            If Me.STATS_FONT_SIZE = 1 Then Exit While
            Me.STATS_FONT_SIZE -= 1
            font_ck = New Font(font_ck.Name, Me.STATS_FONT_SIZE)
        End While

        Me.Font = font_ck

    End Sub

    ' Function to return the coordinate of a certain row and column
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(STATS_FIELD_GAP + STATS_MENU_HEIGHT + x * (STATS_FIELD_GAP + STATS_FIELD_WIDTH),
                        STATS_FIELD_GAP + y * (STATS_FIELD_GAP + STATS_FIELD_HEIGHT))
    End Function

    ' Function to save the data when the button is pressed
    Private Sub view_perf_stats(o As Object, e As EventArgs)
        Dim frm As frmPerfStats = New frmPerfStats(Me)
    End Sub

    ' Function to stop recording
    Private Sub stop_rec(s As Object, e As EventArgs)
        GlobalVariables.switch.stop_running()
        If Not GlobalVariables.switch.is_running Then
            Me.btn_stop.Enabled = False
            Me.btn_stop.Visible = False
        End If
    End Sub

    Private Sub close_form(o As Object, e As EventArgs)
        Me.Close()
    End Sub

End Class
