Public Class Config
    Public COM_ports As List(Of COMPortConfig)
    Public Bells As List(Of BellConfig)
    Public Switch As SwitchConfig
    Public debounce_time As Integer
    Public changes_per_lead As Integer
    Public leads_per_course As Integer

    Public Sub New()
    End Sub

    Public Sub New(ports As List(Of COMPort),
                   bells As List(Of Bell),
                   switch As Switch,
                   debounce_time As Integer,
                   changes_per_lead As Integer,
                   leads_per_course As Integer)
        Me.COM_ports = New List(Of COMPortConfig)
        For Each port In ports
            Me.COM_ports.Add(New COMPortConfig(port))
        Next
        Me.Bells = New List(Of BellConfig)
        For Each bell In bells
            Me.Bells.Add(New BellConfig(bell))
        Next
        Me.Switch = New SwitchConfig(switch)
        Me.debounce_time = debounce_time
        Me.changes_per_lead = changes_per_lead
        Me.leads_per_course = leads_per_course
    End Sub

    ' Function to initialize the global variables with the
    ' Values stored in this config
    Public Function initialize() As Boolean
        ' Start by checking we are in a decent state
        If GlobalVariables.switch.is_running Then
            RcDebug.debug_print("tried to initialize from config but we are running")
            Return False
        End If

        ' Now reset all the global variables
        GlobalVariables.reset()

        ' open the COM ports
        For Each port In Me.COM_ports
            GlobalVariables.COM_ports.Add(port.initilaize())
        Next
#If DEBUG Then
        If Not Testing.test_mode Then
            If GlobalVariables.connect_COM_ports() = False Then
                Return False
            End If
        End If
#Else
        If GlobalVariables.connect_COM_ports() = False Then
                Return False
            End If
#End If

        ' Wait for 50ms. Connecting the COM ports may cause the pins to swap states
        ' We wait so that the state change can't be linked to a bell or switch.
        GlobalVariables.wait(50)

        ' Apply this config
        GlobalVariables.switch = Switch.initialize()
        GlobalVariables.debounce_time = Me.debounce_time
        GlobalVariables.changes_per_lead = Me.changes_per_lead
        GlobalVariables.leads_per_course = Me.leads_per_course
        For Each bell In Me.Bells
            GlobalVariables.bells.Add(bell.initialize())
        Next

        GlobalVariables.update_changes_per_course()

        GlobalVariables.config_loaded = True
        Return True
    End Function

End Class

Public Class COMPortConfig
    Public PortName As String

    Public Sub New()
    End Sub

    Public Sub New(port As IO.Ports.SerialPort)
        Me.PortName = port.PortName
    End Sub

    Public Function initilaize() As IO.Ports.SerialPort
        Dim port As New COMPort()
        port.PortName = Me.PortName
        AddHandler port.PinChanged, AddressOf port_pin_changed_wrapper
        Return port
    End Function
End Class

Public Class BellConfig
    Public name As String
    Public bell_number As Integer
    Public port_pin As PortPinConfig
    Public handstroke_delay As Integer
    Public backstroke_delay As Integer

    Public Sub New()
    End Sub

    Public Sub New(bell As Bell)
        Me.name = bell.name
        Me.bell_number = bell.bell_number
        Me.port_pin = PortPinConfig.create(bell.port_pin)
        Me.handstroke_delay = bell.handstroke_delay
        Me.backstroke_delay = bell.backstroke_delay
    End Sub

    Public Function initialize() As Bell
        Dim bell As New Bell(Me.bell_number)
        If Me.port_pin IsNot Nothing Then
            bell.port_pin = Me.port_pin.initialize
        End If
        bell.handstroke_delay = Me.handstroke_delay
        bell.backstroke_delay = Me.backstroke_delay
        Return bell
    End Function
End Class

Public Class PortPinConfig
    Public port As String
    Public pin As Integer

    Public Sub New()
    End Sub

    Private Sub New(p As PortPin)
        Me.port = p.port
        Me.pin = p.pin
    End Sub

    Public Shared Function create(p As PortPin) As PortPinConfig
        If p Is Nothing Then Return Nothing
        Return New PortPinConfig(p)
    End Function

    Public Function initialize() As PortPin
        Dim p As New PortPin(Me.port, Me.pin)
        Return p
    End Function
End Class

Public Class SwitchConfig
    Public name As String
    Public port_pin As PortPinConfig

    Public Sub New()
    End Sub

    Public Sub New(switch As Switch)
        Me.name = switch.name
        Me.port_pin = PortPinConfig.create(switch.port_pin)
    End Sub

    Public Function initialize() As Switch
        Dim switch As New Switch(Me.name)
        If Me.port_pin IsNot Nothing Then
            switch.port_pin = Me.port_pin.initialize
        End If
        Return switch
    End Function
End Class