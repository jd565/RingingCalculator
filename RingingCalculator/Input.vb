Public MustInherit Class Input
    Protected port_pin_value As PortPin
    Public can_be_configured As Boolean
    Public name As String
    Protected debounce_state_value As Boolean
    Protected state As Integer = 0

    Public ReadOnly Property debounce_state As Boolean
        Get
            Return Me.debounce_state_value
        End Get
    End Property

    Public Property port_pin As PortPin
        Get
            Return Me.port_pin_value
        End Get
        Set(value As PortPin)
            Me.port_pin_value = value
        End Set
    End Property

    Protected Sub debounce_timer_tick(timer As Timer, e As EventArgs)
        RcDebug.debug_print("debounce timer popped")
        timer.Enabled = False
        Me.debounce_state_value = False
    End Sub

    Public Sub configure(port_pin As PortPin)
        Me.port_pin = port_pin

        ' The pyhsical device is now in state 1 as the 0->1 change configured the device. Reflect this here
        Me.state = 1
    End Sub

    Public Overridable Sub trigger_input()
    End Sub

End Class
