Partial Class frmInputTracer
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

    Private Const MAX_DATA_POINTS As Integer = 50
    Private CHART_SIZE As New Size(800, 400)
    Private CHK_SIZE As New Size(100, 25)
    Private UPDOWN_SIZE As New Size(50, 25)
    Private MENU_HEIGHT As Integer = 25
    Private LABEL_SIZE As Size = New Size(100, 25)

    Friend trace As DataVisualization.Charting.Series
    Friend chart As DataVisualization.Charting.Chart
    Friend chart_area As DataVisualization.Charting.ChartArea
    Friend main_menu As ToolStrip
    Friend reset As ToolStripButton
    Friend debounce_enable As CheckBox
    Friend debounce_value As NumericUpDown
    Friend hold_enable As CheckBox
    Friend signal_width As Label

    Private port_pin As PortPin
    Private state As Integer = 0
    Private start_time As New DateTime
    Private timer As Timer
    Private debounce_triggered As Boolean
    Public ReadOnly Property status As Integer
        Get
            Return Me.state
        End Get
    End Property
    Public ReadOnly Property debounce_enabled As Boolean
        Get
            Return Me.debounce_enable.Checked
        End Get
    End Property
    Public ReadOnly Property debounce As Integer
        Get
            If Me.debounce_enabled = False Then Return 0
            Return Me.debounce_value.Value
        End Get
    End Property
    Public ReadOnly Property hold As Boolean
        Get
            Return Me.hold_enable.Checked
        End Get
    End Property

    ' These lists record the states and times to be plotted
    ' They are the x and y values of the graph, so must be kept in sync.
    ' The times are recorded in milliseconds
    Private states As List(Of Integer)
    Private times As List(Of Integer)

    Public Sub New(parent As Form)
        Me.states = New List(Of Integer)
        Me.times = New List(Of Integer)
        Me.start_time = DateTime.Now
        GlobalVariables.input_tracer = True
        Me.generate(parent)
    End Sub

    Private Sub generate(parent As Form)
        Me.chart = New DataVisualization.Charting.Chart
        Me.trace = New DataVisualization.Charting.Series
        Me.chart_area = New DataVisualization.Charting.ChartArea("ca")
        Me.main_menu = New ToolStrip
        Me.reset = New ToolStripButton
        Me.debounce_enable = New CheckBox
        Me.debounce_value = New NumericUpDown
        Me.hold_enable = New CheckBox
        Me.signal_width = New Label
        Me.timer = New Timer()

        'reset
        Me.reset.Name = "reset"
        Me.reset.Text = "Reset"
        AddHandler Me.reset.Click, AddressOf Me.reset_click

        'Menu
        Me.main_menu.Items.Add(Me.reset)
        Me.main_menu.Height = Me.MENU_HEIGHT

        'signal_width
        Me.signal_width.Visible = False
        Me.signal_width.Text = "Signal width: " & 0.ToString
        Me.signal_width.Name = "signal_width"
        Me.signal_width.Size = Me.LABEL_SIZE
        Me.signal_width.Location = New Point(150, 450)

        'hold_enable
        With Me.hold_enable
            .Checked = False
            .Name = "hold_enable"
            .Text = "Hold?"
            .CheckAlign = ContentAlignment.MiddleRight
            .Size = Me.CHK_SIZE
            .Location = New Point(0, 450)
            AddHandler .CheckedChanged, AddressOf Me.hold_checked
        End With

        'debounce_enable
        With Me.debounce_enable
            .Checked = False
            .Name = "debounce_enable"
            .Text = "Debounce?"
            .CheckAlign = ContentAlignment.MiddleRight
            .Size = Me.CHK_SIZE
            .Location = New Point(0, 425)
            AddHandler .CheckedChanged, AddressOf Me.debounce_checked
        End With

        'debounce_value
        With Me.debounce_value
            .Visible = False
            .Enabled = False
            .Maximum = 1000
            .Minimum = 0
            .Value = GlobalVariables.debounce_time
            .Size = Me.UPDOWN_SIZE
            .Location = New Point(150, 425)
        End With

        'timer
        Me.timer.Interval = 100
        AddHandler Me.timer.Tick, AddressOf Me.timer_tick
        Me.timer.Enabled = True

        'chart_area
        Me.chart_area.AxisX.Minimum = -MAX_DATA_POINTS * Me.timer.Interval
        Me.chart_area.AxisX.Maximum = 0
        Me.chart_area.AxisX.LabelStyle.Enabled = False
        Me.chart_area.AxisX.MajorTickMark.Enabled = False
        Me.chart_area.AxisX.MajorGrid.Enabled = False
        Me.chart_area.AxisY.Minimum = 0
        Me.chart_area.AxisY.Maximum = 1
        Me.chart_area.AxisY.MajorGrid.Enabled = False
        Me.chart_area.AxisY.Interval = 1

        Me.trace.ChartArea = "ca"
        Me.trace.ChartType = DataVisualization.Charting.SeriesChartType.Line
        Me.trace.Name = "Input Trace"
        Me.trace.Color = Color.Red

        Me.chart.Size = Me.CHART_SIZE
        Me.chart.Location = New Point(0, 25)
        Me.chart.ChartAreas.Clear()
        Me.chart.ChartAreas.Add(Me.chart_area)
        Me.chart.Series.Add(Me.trace)

        Me.Controls.Add(Me.chart)
        Me.Controls.Add(Me.debounce_enable)
        Me.Controls.Add(Me.main_menu)
        Me.Controls.Add(Me.debounce_value)
        Me.Controls.Add(Me.signal_width)
        Me.Controls.Add(Me.hold_enable)
        Me.ClientSize = New Size(800, 500)
        Me.Text = "Input Tracer"
        Me.Name = "frmInputTracer"

        parent.AddOwnedForm(Me)

        AddHandler Me.FormClosing, AddressOf Me.close_form

        Me.Show()

    End Sub

    Public Sub port_pin_triggered(pp As PortPin)
        ' If the debounce state is triggered then log and do nothing
        If Me.debounce_triggered Then
            Console.WriteLine("debounce drop")
            Exit Sub
        End If

        Console.WriteLine("input tracer pp trig")

        ' Only start debounce if it is enabled and non-zero
        If Me.debounce > 0 Then
            Me.debounce_triggered = True
            start_new_timer(Me.debounce, AddressOf Me.debounce_timer_pop)
        End If

        ' If we haven't got a port_pin yet then latch onto the first one
        ' we see.
        If Me.port_pin Is Nothing Then
            Me.port_pin = pp
            Me.start_time = DateTime.Now()
        End If

        ' If the triggered pin matches our pin then add this to our set of values
        If Me.port_pin.Equals(pp) Then
            Me.add_state_time(Me.get_time_diff.Subtract(New TimeSpan(0, 0, 0, 0, 1)))
            Me.state = (Me.state + 1) Mod 2
            Me.add_state_time(Me.get_time_diff)
            Me.update_graph()
        End If
    End Sub

    Private Sub debounce_timer_pop(t As Timer, e As EventArgs)
        t.Enabled = False
        Me.debounce_triggered = False
    End Sub

    Private Sub add_state_time(time As TimeSpan)
        Me.states.Add(Me.state)
        Me.times.Add(time.TotalMilliseconds)
    End Sub

    Private Sub update_graph()
        'We only want so many points on our graph, so remove items from the beginning of our list
        'Until we have the correct number
        While Me.states.Count > MAX_DATA_POINTS
            Me.states.Remove(Me.states.First)
            Me.times.Remove(Me.times.First)
        End While

        'Now update the graph
        Me.chart_area.AxisX.Minimum = Me.times.Last - MAX_DATA_POINTS * Me.timer.Interval
        Me.chart_area.AxisX.Maximum = Me.times.Last
        Me.trace.Points.DataBindXY(Me.times, Me.states)
        Me.maybe_update_signal_width()
    End Sub

    ' Function to get the width of the signal.
    ' We only do this in hold mode as then we expect only 1 signal on the graph.
    Private Sub maybe_update_signal_width()
        Dim ii As Integer = 0
        Dim signal_width As Integer

        If Not Me.hold Then Exit Sub

        ' Set ii to the index of the first point to be different to the starting value
        While Me.states(ii) = Me.states(0)
            ii += 1
        End While

        ' Then get the time difference
        signal_width = Me.times.Last - Me.times(ii)

        Me.signal_width.Text = "Signal width: " & signal_width.ToString

    End Sub

    Private Sub timer_tick(t As Timer, e As EventArgs)
        Me.add_state_time(Me.get_time_diff)
        If Not Me.hold Then
            Me.update_graph()
        End If
    End Sub

    Private Function get_time_diff() As TimeSpan
        Return DateTime.Now().Subtract(Me.start_time)
    End Function

    Private Sub close_form(o As Object, e As EventArgs)
        GlobalVariables.input_tracer = False
        dispose_of_form(Me, e)
        Me.timer.Enabled = False
    End Sub

    Private Sub reset_click(o As Object, e As EventArgs)
        Me.reset_form
    End Sub

    Private Sub reset_form()
        ' Reset the port pin
        Me.port_pin = Nothing
        Me.debounce_triggered = False
        Me.debounce_enable.Checked = False
        Me.hold_enable.Checked = False
        Me.state = 0
        ' empty the graph
        Me.states.Clear()
        Me.times.Clear()
        Me.chart_area.AxisX.Minimum = -MAX_DATA_POINTS * Me.timer.Interval
        Me.chart_area.AxisX.Maximum = 0
        Me.trace.Points.DataBindXY(Me.times, Me.states)
    End Sub

    Private Sub debounce_checked(o As Object, e As EventArgs)
        Me.debounce_value.Enabled = Me.debounce_enabled
        Me.debounce_value.Visible = Me.debounce_enabled
    End Sub

    Private Sub hold_checked(o As Object, e As EventArgs)
        Me.signal_width.Visible = Me.hold
    End Sub
End Class
