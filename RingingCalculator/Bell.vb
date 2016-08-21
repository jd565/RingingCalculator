Public Class Bell
    Inherits Input

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

    Private bell_number_value As UShort

    Public ReadOnly Property bell_number As UShort
        Get
            Return Me.bell_number_value
        End Get
    End Property

    ' Function to initialise the bell
    Public Sub New(number As UShort)
        Me.bell_number_value = number
        Me.name = "bell" & bell_number.ToString()
        Me.can_be_configured = False
        Me.handstroke_delay_value = 10
        Me.backstroke_delay_value = 10
        Me.new_fields()
    End Sub

    ' Function to generate new fields for use on a new form.
    Public Sub new_fields()
        Me.fields = New BellFields(Me)
    End Sub

    ' Function to handle the delay timers popping.
    Private Sub delay_timer_timeout(timer As Timer, e As EventArgs)
        Dim ct As ChangeTime
        RcDebug.debug_entry("delay_timer_timeout")
        RcDebug.debug_print(Me.name)

        ' Whatever we do we stop the timer.
        timer.Enabled = False

        ct = Me.generate_changetime
        Me.state = (Me.state + 1) Mod 4

        ' If we are recording then check to see if we add to the list
        If GlobalVariables.switch.is_running Then
            If Me.state Mod 2 = 1 Then
                Me.change_times.Add(ct)
                bell_has_just_rung(Me)
            End If
        End If
        Me.update_blob()
        RcDebug.debug_exit()
    End Sub

    ' Function to update the colour of the blob
    Private Sub update_blob()
        If Me.fields.blob.IsDisposed Then Exit Sub
        If Me.state = 1 Then
            Me.fields.blob.BackColor = Color.Red
            start_new_timer(GlobalVariables.bell_light_time, AddressOf Me.gray_blob)
        ElseIf Me.state = 3 Then
            Me.fields.blob.BackColor = Color.Blue
            start_new_timer(GlobalVariables.bell_light_time, AddressOf Me.gray_blob)
        End If

        ' If focus on this is lost it may throw an exception. Catch it and ignore
        Try
            If state Mod 2 = 1 And Form.ActiveForm.GetType Is GetType(frmBells) Then
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
            End If
        Catch
        End Try
    End Sub

    ' Function to change the colour of the blob back to gray
    Private Sub gray_blob(t As Timer, e As EventArgs)
        Me.fields.blob.BackColor = Color.Gray
        t.Enabled = False
    End Sub

    ' Method to add the current time to the list of changes.
    ' Calculated from the time difference between the start of the ringing
    ' and the time now.
    Private Function generate_changetime() As ChangeTime
        Dim current_time As DateTime

        current_time = DateTime.Now()
        Return (New ChangeTime(Me.bell_number, Me.change_times.Count + 1, current_time))
    End Function

    ' Function to handle when the bell sensor changes state.
    ' We want to start a timer based on the delay of the bell, which will ensure that
    ' the bell changes state when it sounds.
    Public Overrides Sub trigger_input()
        RcDebug.debug_print(Me.name & " triggered")

        ' Only do something if the debounce timer isn't running
        If Me.debounce_state = True Then
            RcDebug.debug_print("Debounce drop")
            Exit Sub
        End If
        start_debounce_timer(AddressOf Me.debounce_timer_tick)
        Me.debounce_state_value = True

        RcDebug.debug_print("State " & Me.state)
        ' Depending on the state we need different times on the timer.
        If Me.state \ 2 = 0 Then
            start_new_timer(handstroke_delay, AddressOf delay_timer_timeout)
        Else
            start_new_timer(backstroke_delay, AddressOf delay_timer_timeout)
        End If
    End Sub

    ' Function to reset the bell.
    Public Sub reset()

        ' reset the state
        Me.state = 0
        Me.fields.blob.BackColor = Color.Gray

        ' clear the timing list
        Me.change_times.Clear()
    End Sub

End Class
