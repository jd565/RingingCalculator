Partial Class frmPerfStats
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

    Private PSTATS_TB_WIDTH As Integer = 600
    Private PSTATS_TB_HEIGHT As Integer = 500
    Private PSTATS_FIELD_HEIGHT As Integer = 40
    Private PSTATS_FIELD_WIDTH As Integer = 100
    Private PSTATS_FIELD_GAP As Integer = 15
    Private PSTATS_VAL_WIDTH As Integer = 100
    Private PSTATS_CHART_SIZE As Size = New Size(400, 300)
    Private PSTATS_TB_SIZE As Size = New Size(PSTATS_TB_WIDTH, PSTATS_TB_HEIGHT)
    Private PSTATS_VAL_SIZE As Size = New Size(PSTATS_VAL_WIDTH, PSTATS_FIELD_HEIGHT)
    Private PSTATS_BUTTON_SIZE As Size = New Size(150, 100)
    Private PSTATS_FIELD_SIZE As Size = New Size(PSTATS_FIELD_WIDTH, PSTATS_FIELD_HEIGHT)
    Private PSTATS_FRM_SIZE As Size = New Size(PSTATS_TB_WIDTH + 600, PSTATS_TB_HEIGHT + 150)

    Private parent_frm

    Friend perf_details As TextBox
    Friend save_perf As Button
    Friend save_file As SaveFileDialog
    Friend choose_bell As ComboBox
    Friend average_delay As Label
    Friend average_delay_val As Label
    Friend handstroke_delay As Label
    Friend handstroke_delay_val As Label
    Friend backstroke_delay As Label
    Friend backstroke_delay_val As Label
    Friend handstroke_lead As Label
    Friend handstroke_lead_val As Label
    Friend backstroke_lead As Label
    Friend backstroke_lead_val As Label
    Friend average_tdelay_val As Label
    Friend handstroke_tdelay_val As Label
    Friend backstroke_tdelay_val As Label
    Friend handstroke_tlead_val As Label
    Friend backstroke_tlead_val As Label
    Friend average_ddelay_val As Label
    Friend handstroke_ddelay_val As Label
    Friend backstroke_ddelay_val As Label
    Friend handstroke_dlead_val As Label
    Friend backstroke_dlead_val As Label
    Friend title1 As Label
    Friend title2 As Label
    Friend title3 As Label
    Friend bell_trace As DataVisualization.Charting.Series
    Friend av_trace As DataVisualization.Charting.Series
    Friend chart As DataVisualization.Charting.Chart
    Friend chart_area As DataVisualization.Charting.ChartArea
    Friend choose_bell_label As Label

    ' Note above that there are labels for the bell,
    ' the totals (tdelay) and the difference (ddelay)

    Public Sub New(frm As Form, Optional method As Method = Nothing)
        Me.generate(frm, method)
    End Sub

    Private Sub generate(frm As Form, method As Method)
        Me.perf_details = New TextBox
        Me.save_perf = New Button
        Me.save_file = New SaveFileDialog
        Me.choose_bell = New ComboBox
        Me.average_delay = New Label
        Me.average_delay_val = New Label
        Me.handstroke_delay = New Label
        Me.handstroke_delay_val = New Label
        Me.backstroke_delay = New Label
        Me.backstroke_delay_val = New Label
        Me.handstroke_lead = New Label
        Me.handstroke_lead_val = New Label
        Me.backstroke_lead = New Label
        Me.backstroke_lead_val = New Label
        Me.average_ddelay_val = New Label
        Me.handstroke_ddelay_val = New Label
        Me.backstroke_ddelay_val = New Label
        Me.handstroke_dlead_val = New Label
        Me.backstroke_dlead_val = New Label
        Me.average_tdelay_val = New Label
        Me.handstroke_tdelay_val = New Label
        Me.backstroke_tdelay_val = New Label
        Me.handstroke_tlead_val = New Label
        Me.backstroke_tlead_val = New Label
        Me.title1 = New Label
        Me.title2 = New Label
        Me.title3 = New Label
        Me.chart = New DataVisualization.Charting.Chart
        Me.bell_trace = New DataVisualization.Charting.Series
        Me.av_trace = New DataVisualization.Charting.Series
        Me.chart_area = New DataVisualization.Charting.ChartArea("ca")
        Me.choose_bell_label = New Label

        If method Is Nothing Then
            Statistics.generate_place_delays()
        End If

        'perf_details
        Me.perf_details.Size = PSTATS_TB_SIZE
        Me.perf_details.Font = New Font("Courier New", DefaultFont.Size)
        Me.perf_details.Location = coordinate(0, 0)
        Me.perf_details.Multiline = True
        Me.perf_details.Text = statistics_string(, method)
        Me.perf_details.Name = "perf_details"
        Me.perf_details.ScrollBars = ScrollBars.Both

        'save_perf
        Me.save_perf.Size = PSTATS_BUTTON_SIZE
        Me.save_perf.Location = coordinate(0, 1)
        Me.save_perf.Text = "Save Performance"
        Me.save_perf.Name = "save_perf"
        AddHandler Me.save_perf.Click, AddressOf Me.save_perf_click

        'save_dialog
        Me.save_file.FileName = "ringingcalculator.txt"
        Me.save_file.Filter = "txt files|*.txt|All files|*.*"
        Me.save_file.Title = "Save Performance"
        Me.save_file.DefaultExt = "txt"

        If method Is Nothing Then

            'choose_bell_label
            Me.choose_bell_label.Name = "choose_bell_label"
            Me.choose_bell_label.Text = "Show stats for bell:"
            Me.choose_bell_label.Location = coordinate(1, 0)
            Me.choose_bell_label.Size = PSTATS_FIELD_SIZE

            'choose_bell
            Me.choose_bell.Name = "choose_bell"
            Me.choose_bell.Size = PSTATS_FIELD_SIZE
            Me.choose_bell.Location = coordinate(2, 0)
            For ii = 1 To Statistics.rows(0).size
                Me.choose_bell.Items.Add(ii.ToString)
            Next
            Me.choose_bell.SelectedIndex = 0
            AddHandler Me.choose_bell.SelectedIndexChanged, AddressOf Me.bell_stats_changed

            'average_delay
            Me.average_delay.Name = "average_delay"
            Me.average_delay.Text = "Average delay (ms):"
            Me.average_delay.Size = PSTATS_FIELD_SIZE
            Me.average_delay.Location = coordinate(1, 2)
            Me.average_delay_val.Name = "average_delay_val"
            Me.average_delay_val.Size = PSTATS_VAL_SIZE
            Me.average_delay_val.Location = coordinate(2, 2)

            'handstroke_delay
            Me.handstroke_delay.Name = "handstroke_delay"
            Me.handstroke_delay.Text = "Handstroke delay (ms):"
            Me.handstroke_delay.Size = PSTATS_FIELD_SIZE
            Me.handstroke_delay.Location = coordinate(1, 3)
            Me.handstroke_delay_val.Name = "handstroke_delay_val"
            Me.handstroke_delay_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_delay_val.Location = coordinate(2, 3)

            'backstroke_delay
            Me.backstroke_delay.Name = "backstroke_delay"
            Me.backstroke_delay.Text = "Backstroke delay (ms):"
            Me.backstroke_delay.Size = PSTATS_FIELD_SIZE
            Me.backstroke_delay.Location = coordinate(1, 4)
            Me.backstroke_delay_val.Name = "backstroke_delay_val"
            Me.backstroke_delay_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_delay_val.Location = coordinate(2, 4)

            'handstroke_lead
            Me.handstroke_lead.Name = "handstroke_lead"
            Me.handstroke_lead.Text = "Handstroke lead delay (ms):"
            Me.handstroke_lead.Size = PSTATS_FIELD_SIZE
            Me.handstroke_lead.Location = coordinate(1, 5)
            Me.handstroke_lead_val.Name = "handstroke_lead_val"
            Me.handstroke_lead_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_lead_val.Location = coordinate(2, 5)

            'backstroke_lead
            Me.backstroke_lead.Name = "backstroke_lead"
            Me.backstroke_lead.Text = "Backstroke lead delay (ms):"
            Me.backstroke_lead.Size = PSTATS_FIELD_SIZE
            Me.backstroke_lead.Location = coordinate(1, 6)
            Me.backstroke_lead_val.Name = "backstroke_lead_val"
            Me.backstroke_lead_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_lead_val.Location = coordinate(2, 6)

            'average_tdelay
            Me.average_tdelay_val.Text = Statistics.bell_stats.Average(Function(s) s.average).ToString("####0")
            Me.average_tdelay_val.Name = "average_tdelay_val"
            Me.average_tdelay_val.Size = PSTATS_VAL_SIZE
            Me.average_tdelay_val.Location = coordinate(3, 2)

            'handstroke_tdelay
            Me.handstroke_tdelay_val.Text = Statistics.bell_stats.Average(Function(s) s.h_average).ToString("####0")
            Me.handstroke_tdelay_val.Name = "handstroke_tdelay_val"
            Me.handstroke_tdelay_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_tdelay_val.Location = coordinate(3, 3)

            'backstroke_tdelay
            Me.backstroke_tdelay_val.Text = Statistics.bell_stats.Average(Function(s) s.b_average).ToString("####0")
            Me.backstroke_tdelay_val.Name = "backstroke_tdelay_val"
            Me.backstroke_tdelay_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_tdelay_val.Location = coordinate(3, 4)

            'handstroke_tlead
            Me.handstroke_tlead_val.Text = Statistics.bell_lead_stats.Average(Function(s) s.h_average).ToString("####0")
            Me.handstroke_tlead_val.Name = "handstroke_tlead_val"
            Me.handstroke_tlead_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_tlead_val.Location = coordinate(3, 5)

            'backstroke_tlead
            Me.backstroke_tlead_val.Text = Statistics.bell_lead_stats.Average(Function(s) s.b_average).ToString("####0")
            Me.backstroke_tlead_val.Name = "backstroke_tlead_val"
            Me.backstroke_tlead_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_tlead_val.Location = coordinate(3, 6)

            'average_ddelay
            Me.average_ddelay_val.Name = "average_ddelay_val"
            Me.average_ddelay_val.Size = PSTATS_VAL_SIZE
            Me.average_ddelay_val.Location = coordinate(4, 2)

            'handstroke_ddelay
            Me.handstroke_ddelay_val.Name = "handstroke_ddelay_val"
            Me.handstroke_ddelay_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_ddelay_val.Location = coordinate(4, 3)

            'backstroke_ddelay
            Me.backstroke_ddelay_val.Name = "backstroke_ddelay_val"
            Me.backstroke_ddelay_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_ddelay_val.Location = coordinate(4, 4)

            'handstroke_dlead
            Me.handstroke_dlead_val.Name = "handstroke_dlead_val"
            Me.handstroke_dlead_val.Size = PSTATS_VAL_SIZE
            Me.handstroke_dlead_val.Location = coordinate(4, 5)

            'backstroke_dlead
            Me.backstroke_dlead_val.Name = "backstroke_dlead_val"
            Me.backstroke_dlead_val.Size = PSTATS_VAL_SIZE
            Me.backstroke_dlead_val.Location = coordinate(4, 6)

            'title1
            Me.title1.Name = "title1"
            Me.title1.Size = PSTATS_VAL_SIZE
            Me.title1.Location = coordinate(2, 1)
            Me.title1.Text = "Bell delays"

            'title2
            Me.title2.Name = "title2"
            Me.title2.Size = PSTATS_VAL_SIZE
            Me.title2.Location = coordinate(3, 1)
            Me.title2.Text = "Total delays"

            'title3
            Me.title3.Name = "title3"
            Me.title3.Size = PSTATS_VAL_SIZE
            Me.title3.Location = coordinate(4, 1)
            Me.title3.Text = "Difference"

            'chart_area
            Me.chart_area.AxisX.MajorTickMark.Enabled = False
            Me.chart_area.AxisX.MajorGrid.Enabled = False
            Me.chart_area.AxisY.LabelStyle.Enabled = False
            Me.chart_area.AxisY.Minimum = 0
            Me.chart_area.AxisY.Maximum = 1
            Me.chart_area.AxisY.MajorGrid.Enabled = False
            Me.chart_area.AxisX.Interval = 5

            Me.bell_trace.ChartArea = "ca"
            Me.bell_trace.ChartType = DataVisualization.Charting.SeriesChartType.Line
            Me.bell_trace.Name = "Bell Trace"
            Me.bell_trace.Color = Color.Red

            Me.av_trace.ChartArea = "ca"
            Me.av_trace.ChartType = DataVisualization.Charting.SeriesChartType.Line
            Me.av_trace.Name = "Avg Trace"
            Me.av_trace.Color = Color.Blue

            Me.chart.Size = PSTATS_CHART_SIZE
            Me.chart.Location = coordinate(1, 7)
            Me.chart.ChartAreas.Clear()
            Me.chart.ChartAreas.Add(Me.chart_area)
            Me.chart.Series.Add(Me.bell_trace)
            Me.chart.Series.Add(Me.av_trace)

            'Generate the average trace
            Me.generate_trace()

            'Generate the inital values
            Me.generate_delay_vals()

            Me.Controls.Add(Me.choose_bell)
            Me.Controls.Add(Me.average_delay)
            Me.Controls.Add(Me.average_delay_val)
            Me.Controls.Add(Me.handstroke_delay)
            Me.Controls.Add(Me.handstroke_delay_val)
            Me.Controls.Add(Me.backstroke_delay)
            Me.Controls.Add(Me.backstroke_delay_val)
            Me.Controls.Add(Me.handstroke_lead)
            Me.Controls.Add(Me.handstroke_lead_val)
            Me.Controls.Add(Me.backstroke_lead)
            Me.Controls.Add(Me.backstroke_lead_val)
            Me.Controls.Add(Me.average_tdelay_val)
            Me.Controls.Add(Me.handstroke_tdelay_val)
            Me.Controls.Add(Me.backstroke_tdelay_val)
            Me.Controls.Add(Me.handstroke_tlead_val)
            Me.Controls.Add(Me.backstroke_tlead_val)
            Me.Controls.Add(Me.average_ddelay_val)
            Me.Controls.Add(Me.handstroke_ddelay_val)
            Me.Controls.Add(Me.backstroke_ddelay_val)
            Me.Controls.Add(Me.handstroke_dlead_val)
            Me.Controls.Add(Me.backstroke_dlead_val)
            Me.Controls.Add(Me.choose_bell_label)
            Me.Controls.Add(Me.chart)

        End If

        'frm
        Me.Controls.Add(Me.perf_details)
        Me.Controls.Add(Me.save_perf)
        Me.Controls.Add(Me.title1)
        Me.Controls.Add(Me.title2)
        Me.Controls.Add(Me.title3)

        Me.AutoScroll = True
        Me.Text = "Performance Statistics"
        Me.Name = "frmPerfStats"
        Me.ClientSize = PSTATS_FRM_SIZE
        Me.parent_frm = frm
        frm.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        Me.Show()
    End Sub

    Private Sub save_perf_click(o As Object, e As EventArgs)
        If Me.save_file.ShowDialog() = DialogResult.OK Then
            Saving.save_statistics(Me.save_file.FileName)
            MsgBox("Performance saved.",, "File Saved")
        Else
            MsgBox("Performance not saved.",, "Save Performance")
        End If
    End Sub

    'Function to generate the average trace
    Private Sub generate_trace(Optional bell_index As Integer = -1)
        Dim x_vals As New List(Of Double)
        Dim y_vals As New List(Of Double)
        Dim std_dev As Double
        Dim mean As Double
        Dim min_x As Double
        Dim max_x As Double
        Dim max_y As Double
        Dim x_inc As Double
        'expectation of the square mean
        Dim e_x2 As Double

        If bell_index = -1 Then
            ' I think this should work
            mean = Statistics.bell_stats.Average(Function(s) s.average)
            e_x2 = Statistics.bell_stats.Average(Function(s) (s.std_dev ^ 2) + (s.average ^ 2))
            std_dev = Math.Sqrt(e_x2 - mean ^ 2)
        Else
            mean = Statistics.bell_stats(bell_index).average
            std_dev = Statistics.bell_stats(bell_index).std_dev
        End If

        min_x = mean - 2 * std_dev
        max_x = mean + 2 * std_dev
        x_inc = (max_x - min_x) / 49
        For ii = 0 To 49
            x_vals.Add(min_x + ii * x_inc)
            y_vals.Add((1 / (std_dev * Math.Sqrt(2 * Math.PI))) * Math.E ^ (-((x_vals(ii) - mean) ^ 2) / (2 * std_dev ^ 2)))
            max_y = Math.Max(max_y, y_vals(ii))
        Next

        If bell_index = -1 Then
            Me.av_trace.Points.DataBindXY(x_vals, y_vals)
            Me.chart_area.AxisY.Maximum = Math.Min(max_y + 0.05, 1)
            Me.chart_area.AxisX.Minimum = Math.Floor(min_x)
            Me.chart_area.AxisX.Maximum = Math.Ceiling(max_x)
        Else
            Me.bell_trace.Points.DataBindXY(x_vals, y_vals)
        End If
    End Sub

    'Function to set the text value of the fields on this form
    Private Sub generate_delay_vals()
        Dim bell_index As Integer
        bell_index = Me.choose_bell.SelectedIndex
        Me.average_delay_val.Text = Statistics.bell_stats(bell_index).average.ToString("####0")
        Me.handstroke_delay_val.Text = Statistics.bell_stats(bell_index).h_average.ToString("####0")
        Me.backstroke_delay_val.Text = Statistics.bell_stats(bell_index).b_average.ToString("####0")
        Me.handstroke_lead_val.Text = Statistics.bell_lead_stats(bell_index).h_average.ToString("####0")
        Me.backstroke_lead_val.Text = Statistics.bell_lead_stats(bell_index).b_average.ToString("####0")
        Me.average_ddelay_val.Text = (Val(Me.average_delay_val.Text) - Val(Me.average_tdelay_val.Text)).ToString("####0")
        Me.handstroke_ddelay_val.Text = (Val(Me.handstroke_delay_val.Text) - Val(Me.handstroke_tdelay_val.Text)).ToString("####0")
        Me.backstroke_ddelay_val.Text = (Val(Me.backstroke_delay_val.Text) - Val(Me.backstroke_tdelay_val.Text)).ToString("####0")
        Me.handstroke_dlead_val.Text = (Val(Me.handstroke_lead_val.Text) - Val(Me.handstroke_tlead_val.Text)).ToString("####0")
        Me.backstroke_dlead_val.Text = (Val(Me.backstroke_lead_val.Text) - Val(Me.backstroke_tlead_val.Text)).ToString("####0")
        Me.generate_trace(bell_index)
    End Sub

    'Function to handle the bell to show stats for changing
    Private Sub bell_stats_changed(o As Object, e As EventArgs)
        Me.generate_delay_vals()
    End Sub

    'Function to calculate coordinates for this form
    Private Function coordinate(x As Integer, y As Integer) As Point
        Dim ii As Integer
        Dim jj As Integer

        jj = PSTATS_FIELD_GAP + y * (PSTATS_FIELD_HEIGHT + PSTATS_FIELD_GAP)

        If x = 0 Then
            If y = 1 Then
                ii = PSTATS_FIELD_GAP
                jj = PSTATS_TB_HEIGHT + PSTATS_FIELD_GAP
            Else
                ii = 0
                jj = 0
            End If
        ElseIf x = 1
            ii = PSTATS_TB_WIDTH + PSTATS_FIELD_GAP
        Else
            ii = PSTATS_TB_WIDTH + PSTATS_FIELD_GAP + PSTATS_FIELD_WIDTH + (x - 2) * PSTATS_VAL_WIDTH
        End If
        Return New Point(ii, jj)
    End Function

End Class
