Public Class Bell
    Private port_pin_value As PortPin
    Public Property port_pin As PortPin
        Get
            Return Me.port_pin_value
        End Get
        Set(value As PortPin)
            Me.port_pin_value = value
        End Set
    End Property

    ' state is one of 0, 1, 2 or 3. This represents handstroke, handstroke on, and backstroke, backstroke on.
    Private state As UShort
    Public ReadOnly Property get_state As UShort
        Get
            Return Me.state
        End Get
    End Property

    ' The fields that appear for this bell in a form
    Public fields As BellFields

    ' The list index gives us the number of the change so no need to record this.
    ' Note the list index will start from 0, so deal with this when needed.
    Public change_times As New List(Of ChangeTime)

    ' The delay of the bell, measured in milliseconds gap between signal and bell sounding
    Private handstroke_delay_value As Integer
    Private backstroke_delay_value As Integer

    Public Property handstroke_delay As Integer
        Get
            Return Me.handstroke_delay_value
        End Get
        Set(value As Integer)
            Me.handstroke_delay_value = value
        End Set
    End Property

    Public Property backstroke_delay As Integer
        Get
            Return Me.backstroke_delay_value
        End Get
        Set(value As Integer)
            Me.backstroke_delay_value = value
        End Set
    End Property

    Private name_value As String

    Public ReadOnly Property name As String
        Get
            Return Me.name_value
        End Get
    End Property

    Private bell_number_value As UShort

    Public ReadOnly Property bell_number As UShort
        Get
            Return Me.bell_number_value
        End Get
    End Property

    ' Variable to show whether the bell is waiting to be configured with a new port and pin.
    Public can_be_configured As Boolean

    Private debounce_timer_state As Boolean
    Public ReadOnly Property debounce_state As Boolean
        Get
            Return debounce_timer_state
        End Get
    End Property

    ' Function to initialise the bell
    Public Sub New(number As UShort)
        Me.bell_number_value = number
        Me.name_value = "bell" & bell_number.ToString()
        Me.can_be_configured = False
        Me.handstroke_delay_value = 10
        Me.backstroke_delay_value = 10
        Me.fields = New BellFields(Me)
    End Sub

    ' Method to handle debounce timer popping
    Private Sub debounce_timer_timeout(timer As Timer, e As EventArgs)
        Console.WriteLine("debounce timer popped")
        Me.debounce_timer_state = False
        timer.Enabled = False
    End Sub

    ' Function to see if we should begin recording
    Private Function should_we_start_recording() As Boolean
        If Me.bell_number = 2 Then
            Console.WriteLine("We are bell 1 or 2. Start recording.")
            GlobalVariables.recording = True
            Return True
        End If
        Return False
    End Function

    ' Function to handle the delay timers popping.
    Private Sub delay_timer_timeout(timer As Timer, e As EventArgs)
        Console.WriteLine("{0} delay timer timeout", Me.name)

        ' Whatever we do we stop the timer.
        timer.Enabled = False

        ' If the switch is running but we aren't recording then drop out.
        ' We do this to avoid recording the first change as e.g. 67821435
        ' We set the program to record when bell 2 is rung after the switch starts running
        If Not (GlobalVariables.switch.isRunning = GlobalVariables.recording) Then
            Console.WriteLine("Switch is on but not recording.")
            If Not Me.should_we_start_recording() Then
                GoTo EXIT_LABEL
            End If
        End If

        state += 1

        ' We arent running so change the colour of the light
        If (Me.state Mod 2 = 1) Then
            Me.fields.blob.BackColor = Color.Red

            ' If we are running then check to see if we add to the list
            If GlobalVariables.recording Then
                Me.add_change_to_list()
            End If
        Else
            Me.fields.blob.BackColor = Color.Gray
        End If

        ' If we have reached state 4, set the state back to 0.
        If Me.state = 4 Then
            Me.state = 0
        End If
EXIT_LABEL:
    End Sub

    ' Method to add the current time to the list of changes.
    ' Calculated from the time difference between the start of the ringing
    ' and the time now.
    Private Sub add_change_to_list()
        Dim current_time As DateTime
        Dim time_diff As UInteger

        current_time = DateTime.Now()
        time_diff = (current_time.Ticks - GlobalVariables.start_time.Ticks) / TimeSpan.TicksPerMillisecond
        Me.change_times.Add(New ChangeTime(Me.bell_number, time_diff))
    End Sub

    ' Function to handle when the bell sensor changes state.
    ' We want to start a timer based on the delay of the bell, which will ensure that
    ' the bell changes state when it sounds.
    Public Sub trigger_bell()
        Console.WriteLine("{0} triggered", Me.name)

        ' Only do something if the debounce timer isn't running
        If Me.debounce_timer_state = True Then
            Console.WriteLine("Debounce drop")
            GoTo EXIT_LABEL
        End If
        start_debounce_timer(AddressOf debounce_timer_timeout)

        ' Depending on the state we need different times on the timer.
        If Me.state / 2 = 0 Then
            start_new_timer(handstroke_delay, AddressOf delay_timer_timeout)
        Else
            start_new_timer(backstroke_delay, AddressOf delay_timer_timeout)
        End If
EXIT_LABEL:
    End Sub

    ' Function to reset the bell.
    Public Sub reset()

        ' reset the state
        Me.state = 0

        ' clear the timing list
        Me.change_times.Clear()
    End Sub

End Class
