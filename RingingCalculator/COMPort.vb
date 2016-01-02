Public Class COMPort
    Inherits System.IO.Ports.SerialPort

    Public txt_field As TextBox
    Public label As Label

    Public Sub New()
        Me.new_fields()
    End Sub

    Public Sub new_fields()
        Me.txt_field = New TextBox
        Me.label = New Label
    End Sub

End Class
