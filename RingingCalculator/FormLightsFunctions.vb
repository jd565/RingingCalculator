Partial Class frmLights
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

    Const LIGHT_FIELD_GAP As Integer = 10
    Const LIGHT_LABEL_HEIGHT As Integer = 20
    Const LIGHT_FIELD_HEIGHT As Integer = 50
    Const LIGHT_FIELD_WIDTH As Integer = 70
    Const LIGHT_LIGHT_DIAMETER As Integer = 50
    Private LIGHT_LABEL_SIZE As New Size(LIGHT_FIELD_WIDTH, LIGHT_LABEL_HEIGHT)
    Private LIGHT_GENERAL_SIZE As New Size(LIGHT_FIELD_WIDTH, LIGHT_FIELD_HEIGHT)
    Private LIGHT_LIGHT_SIZE As New Size(Math.Min(Math.Min(LIGHT_LIGHT_DIAMETER, LIGHT_FIELD_HEIGHT), LIGHT_FIELD_WIDTH),
                                        Math.Min(Math.Min(LIGHT_LIGHT_DIAMETER, LIGHT_FIELD_HEIGHT), LIGHT_FIELD_WIDTH))
    Private LIGHT_LIGHT_OFFSET As New Point((LIGHT_FIELD_WIDTH - LIGHT_LIGHT_SIZE.Width) / 2,
                                           (LIGHT_FIELD_HEIGHT - LIGHT_LIGHT_SIZE.Height) / 2)

    Public Sub generate(parent As Form)
        Dim bell_label As Label
        Dim ii As Integer = 0

        For Each bell In GlobalVariables.bells
            bell.new_fields()

            bell_label = New Label()
            bell_label.Text = "Bell " + bell.bell_number.ToString
            bell_label.Size = LIGHT_LABEL_SIZE
            bell_label.Location = coordinate(ii, 0)
            bell_label.Name = "lbl" + bell.name

            bell.fields.blob.Size = LIGHT_LIGHT_SIZE
            bell.fields.blob.Location = coordinate(ii, 1) + LIGHT_LIGHT_OFFSET

            Me.Controls.Add(bell.fields.canvas)
            Me.Controls.Add(bell_label)

            ii += 1
        Next

        Me.Text = "Lights"
        Me.Name = "frmLights"
        Me.Font = DEFAULT_FONT
        Me.ClientSize = New Size(coordinate(ii, 2))
        parent.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form
        Me.Show()
    End Sub

    Private Function coordinate(x As Integer, y As Integer) As Point
        Dim ii As Integer = LIGHT_FIELD_GAP + x * (LIGHT_FIELD_GAP + LIGHT_FIELD_WIDTH)
        Dim jj As Integer

        If y <= 1 Then jj = LIGHT_FIELD_GAP + y * (LIGHT_FIELD_GAP + LIGHT_LABEL_HEIGHT)
        If y > 1 Then jj = LIGHT_LABEL_HEIGHT + 2 * LIGHT_FIELD_GAP + (y - 1) * (LIGHT_FIELD_GAP + LIGHT_FIELD_HEIGHT)
        Return New Point(ii, jj)
    End Function

End Class
