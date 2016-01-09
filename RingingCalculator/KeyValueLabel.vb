Public Class KeyValueLabel
    Public key As Label
    Public value As Label
    Public enabled As Boolean
    Public name As String

    Public Sub New(enabled As Boolean, name As String)
        Me.key = New Label
        Me.value = New Label
        Me.enabled = enabled
        Me.name = name
    End Sub

    Public Sub new_fields()
        Me.key = New Label
        Me.value = New Label
    End Sub
End Class
