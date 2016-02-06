Public Class Switch
    Inherits Input
    Private switch_running As Boolean
    Public button As New Button

    ' Function to initialize an object of this class.
    ' This just makes sure that the switch is initialized not running
    Public Sub New(name As String)
        Me.name = name
        Me.switch_running = False

        Me.new_button()
    End Sub

    ' Function to reset the button as it will be used on a new form.
    Public Sub new_button()
        Me.button = New Button
        Me.button.Name = Me.name + "_btn"
        Me.button.Text = "Configure switch"
        AddHandler Me.button.Click, AddressOf configure_button_pressed
    End Sub

    Public ReadOnly Property is_running() As Boolean
        Get
            Return Me.switch_running
        End Get
    End Property

    ' Function for changing the state of the switch
    Public Overrides Sub trigger_input()
        If Me.debounce_state_value Then Exit Sub

        Me.debounce_state_value = True
        start_debounce_timer(AddressOf Me.debounce_timer_tick)
        Me.state = (Me.state + 1) Mod 2

        If Me.state = 0 Then Exit Sub

        If Not Me.is_running Then
            Me.start_running()
        Else
            Me.stop_running()
        End If
    End Sub

    ' Function to handle when the program starts recording
    Public Sub start_running()
        ' First check that the current form is either frmPerf or frmStats.
        ' These are the only forms we should start from
        If Not (Form.ActiveForm.GetType Is GetType(frmPerf) Or
                Form.ActiveForm.GetType Is GetType(frmStats)) Then
            ' We should not start running
            Exit Sub
        End If
        ' If there is no loaded config then do not start
        If Not GlobalVariables.config_loaded Then Exit Sub

        Dim frm As frmStats

        For Each bell In GlobalVariables.bells
            bell.reset()
        Next
        Statistics.reset_stats()
        Me.switch_running = True
        Console.WriteLine("Started running.")
        frm = New frmStats(frmPerf)
    End Sub

    ' Function to handle when the program stops recording
    Public Sub stop_running()
        Me.switch_running = False
        Console.WriteLine("Stopped running")

        Try
            ' If we are stopping and the row is full then stop here.
            If Statistics.rows.Last.row_is_full Then
                GlobalVariables.method_started = False
            End If
        Catch
        End Try
    End Sub

    ' Function to handle when the configure button is pressed.
    Public Sub configure_button_pressed(btn As Button, e As EventArgs)
        Dim frm As New frmConfigure(Me)
        Me.can_be_configured = True

        ' We only want 1 configuration happening at a time, so pop out another form to stop this happening
        frm.generate(btn.Parent)
    End Sub

    Public Sub reset()
        Me.stop_running()
        Me.state = 0
    End Sub

End Class
