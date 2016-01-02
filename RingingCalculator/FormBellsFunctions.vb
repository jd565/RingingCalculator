Partial Class frmBells
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

    Const BELL_FIELD_GAP As Integer = 40
    Const BELL_LABEL_HEIGHT As Integer = 25
    Const BELL_UPDOWN_HEIGHT As Integer = 25
    Const BELL_FIELD_HEIGHT As Integer = BELL_LABEL_HEIGHT + BELL_UPDOWN_HEIGHT
    Const BELL_FIELD_WIDTH As Integer = 100
    Const BELL_LIGHT_DIAMETER As Integer = 50
    Private BELL_LABEL_SIZE As New Size(BELL_FIELD_WIDTH, BELL_LABEL_HEIGHT)
    Private BELL_UPDOWN_SIZE As New Size(BELL_FIELD_WIDTH, BELL_UPDOWN_HEIGHT)
    Private BELL_UPDOWN_OFFSET As New Point(0, BELL_LABEL_HEIGHT)
    Private BELL_GENERAL_SIZE As New Size(BELL_FIELD_WIDTH, BELL_FIELD_HEIGHT)
    Private BELL_LIGHT_SIZE As New Size(Math.Min(Math.Min(BELL_LIGHT_DIAMETER, BELL_FIELD_HEIGHT), BELL_FIELD_WIDTH),
                                        Math.Min(Math.Min(BELL_LIGHT_DIAMETER, BELL_FIELD_HEIGHT), BELL_FIELD_WIDTH))
    Private BELL_LIGHT_OFFSET As New Point((BELL_FIELD_WIDTH - BELL_LIGHT_SIZE.Width) / 2,
                                           (BELL_FIELD_HEIGHT - BELL_LIGHT_SIZE.Height) / 2)
    Friend debounce_time As NumericUpDown
    Friend debounce_label As Label
    Friend btn_ok As Button
    Private parent_frm As Form

    ' Function to return the coordinate of a point on the grid
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(BELL_FIELD_GAP + x * (BELL_FIELD_GAP + BELL_FIELD_WIDTH),
                         BELL_FIELD_GAP + y * (BELL_FIELD_GAP + BELL_FIELD_HEIGHT))
    End Function

    ' Function to generate the bells form with the required number of bells
    ' We decide a required width per bell, and set the overall form width.
    ' We then generate the required boxes for each bell.
    Public Sub generate(parent As Form)
        Dim ii As Integer = 0

        Me.new_frm_bells_fields()

        ' Switch
        GlobalVariables.switch.new_button()
        GlobalVariables.switch.button.Size = BELL_GENERAL_SIZE
        GlobalVariables.switch.button.Location = coordinate(0, 0)

        ' Debounce time
        Me.debounce_label.Name = "debounce_time_label"
        Me.debounce_label.Text = "Debounce delay:"
        Me.debounce_label.Size = BELL_LABEL_SIZE
        Me.debounce_label.Location = coordinate(1, 0)
        Me.debounce_time.Name = "debounce_time_updown"
        Me.debounce_time.Value = GlobalVariables.debounce_time
        Me.debounce_time.Size = BELL_UPDOWN_SIZE
        Me.debounce_time.Location = coordinate(1, 0) + BELL_UPDOWN_OFFSET
        AddHandler Me.debounce_time.ValueChanged, AddressOf debounce_time_changed

        ' Bells
        ' Reset the column width as the bells go on the row below these configs
        ii = 0
        For Each bell In GlobalVariables.bells
            add_bell_field(ii, bell)
            ii += 1
        Next

        ' btn_ok
        Me.btn_ok.Name = "btn_ok"
        Me.btn_ok.Size = BELL_GENERAL_SIZE
        Me.btn_ok.Location = coordinate(0, 5)
        Me.btn_ok.Text = "OK"
        AddHandler Me.btn_ok.Click, AddressOf Me.btn_ok_click

        Me.Controls.Add(GlobalVariables.switch.button)
        Me.Controls.Add(debounce_time)
        Me.Controls.Add(debounce_label)
        Me.Controls.Add(btn_ok)

        Me.Text = "Ringing Simulator"
        Me.Name = "frmBells"
        Me.Font = DEFAULT_FONT
        Me.ClientSize = New Size(coordinate(ii, 6))
        parent.AddOwnedForm(Me)
        Me.Parent_frm = parent
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        Me.Show()
        parent.Hide()

    End Sub

    ' Function to add one bell. The bell fields will be placed at the coordinate specified,
    ' and named using bell_name
    Private Sub add_bell_field(x As Integer, bell As Bell)
        bell.new_fields()

        ' Hnadstroke delay
        bell.fields.handstroke_delay.Size = BELL_UPDOWN_SIZE
        bell.fields.handstroke_delay.Location = coordinate(x, 1) + BELL_UPDOWN_OFFSET
        bell.fields.handstroke_label.Size = BELL_LABEL_SIZE
        bell.fields.handstroke_label.Location = coordinate(x, 1)

        ' Backstroke delay
        bell.fields.backstroke_delay.Size = BELL_UPDOWN_SIZE
        bell.fields.backstroke_delay.Location = coordinate(x, 2) + BELL_UPDOWN_OFFSET
        bell.fields.backstroke_label.Size = BELL_LABEL_SIZE
        bell.fields.backstroke_label.Location = coordinate(x, 2)

        ' configure port pin button
        bell.fields.button.Size = BELL_GENERAL_SIZE
        bell.fields.button.Location = coordinate(x, 3)

        ' Light
        bell.fields.blob.Size = BELL_LIGHT_SIZE
        bell.fields.blob.Location = coordinate(x, 4) + BELL_LIGHT_OFFSET

        Me.Controls.Add(bell.fields.button)
        Me.Controls.Add(bell.fields.handstroke_delay)
        Me.Controls.Add(bell.fields.handstroke_label)
        Me.Controls.Add(bell.fields.backstroke_delay)
        Me.Controls.Add(bell.fields.backstroke_label)
        Me.Controls.Add(bell.fields.canvas)

    End Sub

    Private Sub debounce_time_changed(debounce_time As NumericUpDown, e As EventArgs)
        GlobalVariables.debounce_time_changed(debounce_time.Value)
    End Sub

    Private Sub new_frm_bells_fields()
        Me.debounce_time = New NumericUpDown
        Me.debounce_label = New Label
        Me.btn_ok = New Button
    End Sub

    Private Sub btn_ok_click(b As Button, e As EventArgs)
        Dim frm As frmPerf
        frm = find_form(GetType(frmPerf))
        frm.save_config()
    End Sub

End Class
