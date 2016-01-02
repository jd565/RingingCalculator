Public Class PortPin
    Private port_name_value As String
    Private pin_value As Integer

    Public ReadOnly Property port As String
        Get
            Return port_name_value
        End Get
    End Property

    Public ReadOnly Property pin As Integer
        Get
            Return pin_value
        End Get
    End Property

    Public Sub New(port_name As String, pin As Integer)
        Me.port_name_value = port_name
        Me.pin_value = pin
    End Sub

    ' Function to say whether the supplied object matches this one.
    ' Check that this object exists, and that it is a PortPin object,
    ' Return true if the port_name and pin values are the same.
    Public Overrides Function Equals(obj As Object) As Boolean

        If obj Is Nothing Then Return False
        If obj.GetType() IsNot GetType(PortPin) Then Return False

        Dim port_pin As PortPin = CType(obj, PortPin)
        If port_pin.port Is Nothing Then Return False
        Return (port_pin.port.Equals(Me.port) And port_pin.pin = Me.pin)

    End Function

End Class
