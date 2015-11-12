Public Class Bell
    Private serial_port As String
    Private port_pin As Integer

    ' Function to set the serial_port and port_pin when the bell is being configured.
    Public Sub set_port(port As String, pin As Integer)
        Me.serial_port = port
        Me.port_pin = pin
    End Sub

    ' Function to check if the port and pin matches this bell
    Public Function check_port(port As String, pin As Integer) As Boolean
        Return (port.Equals(Me.serial_port) And pin = Me.port_pin)
    End Function

    ' state is one of 0, 1, 2 or 3. This represents handstroke, handstroke on, and backstroke, backstroke on.
    Private state As UShort

    ' The list index gives us the number of the change so no need to record this.
    ' Note the list index will start from 0, so deal with this when needed.
    Public change_times As List(Of Integer)

    ' The delay of the bell, measured in milliseconds gap between signal and bell sounding
    Public handstroke_delay As Integer
    Public backstroke_delay As Integer

    Public name As String

    ' Variable to show whether the bell is waiting to be configured with a new port and pin.
    Public can_be_configured As Boolean

    Public button As Button

    ' Handstroke and backstroke have different delays, so easiest to use different timers.
    ' It is possible that the bell changes state while the timer is still counting
    ' So need a different timer for each state change
    Private WithEvents state01_timer As New Timer
    Private WithEvents state12_timer As New Timer
    Private WithEvents state23_timer As New Timer
    Private WithEvents state30_timer As New Timer
    Private WithEvents debounce_timer As New Timer

    ' Function to initialise the bell
    Public Sub New(name As String)
        Me.name = name
        Me.can_be_configured = False
    End Sub

    ' Method to set the delays on the timers
    Public Sub set_delays(handstroke_delay As Integer,
                          backstroke_delay As Integer,
                          timer_offset As UInteger,
                          debounce_delay As UInteger)

        ' Store off the delays into the local variables, in case we need them later
        Me.handstroke_delay = handstroke_delay
        Me.backstroke_delay = backstroke_delay

        ' We need to set the timers to delay by the delay + offset + 1, so that each timer counts down
        ' from a positive number, and so that the biggest negative delay has overall delay 1. This ensures that
        ' all timers are the same delay relative to each other.
        Me.state01_timer.Interval = handstroke_delay + timer_offset + 1
        Me.state12_timer.Interval = handstroke_delay + timer_offset + 1
        Me.state23_timer.Interval = backstroke_delay + timer_offset + 1
        Me.state30_timer.Interval = backstroke_delay + timer_offset + 1

        Me.debounce_timer.Interval = debounce_delay
    End Sub

    'Method to handle debounce timer popping
    Private Sub debounce_tick0(timer As Timer, e As EventArgs) Handles debounce_timer.Tick
        timer.Enabled = False
    End Sub

    ' Function to handle the delay timers popping.
    Private Sub delay_timer_timeout(timer As Timer, e As EventArgs) _
            Handles state01_timer.Tick, state12_timer.Tick, state23_timer.Tick, state30_timer.Tick

        ' Increment the state, stop the timer
        timer.Enabled = False
        state += 1

        ' If we have reached state 4, set the state back to 0.
        If Me.state = 4 Then
            Me.state = 0
        End If

        ' If we have just reached state 1 or 3 then the bell has just rung.
        If (Me.state = 1 Or Me.state = 3) Then
            Me.add_change_to_list()
        End If

    End Sub

    ' Method to add the current time to the list of changes.
    ' Calculated from the time difference between the start of the ringing
    ' and the time now.
    Private Sub add_change_to_list()
        Dim current_time As DateTime
        Dim time_diff As UInteger
        Dim start_time As DateTime
        current_time = DateTime.Now()
        time_diff = (current_time.Ticks - start_time.Ticks) / TimeSpan.TicksPerMillisecond
        Me.change_times.Add(time_diff)
    End Sub

    ' Function to handle when the bell sensor changes state.
    ' We want to start a timer based on the delay of the bell, which will ensure that
    ' the bell changes state when it sounds.
    Public Sub trigger_bell()

        ' Only do something if the debounce timer isn't running
        If Me.debounce_timer.Enabled = False Then
            Me.debounce_timer.Enabled = True

            ' Depending on the state we need different times on the timer.
            Select Case Me.state
                Case 0
                    Me.state01_timer.Enabled = True
                Case 1
                    Me.state12_timer.Enabled = True
                Case 2
                    Me.state23_timer.Enabled = True
                Case 3
                    Me.state30_timer.Enabled = True
            End Select
        End If
    End Sub

    ' Function to reset the bell.
    Public Sub reset()
        ' deactivate all timers
        Me.state01_timer.Enabled = False
        Me.state12_timer.Enabled = False
        Me.state23_timer.Enabled = False
        Me.state30_timer.Enabled = False

        ' reset the state
        Me.state = 0
    End Sub

End Class
