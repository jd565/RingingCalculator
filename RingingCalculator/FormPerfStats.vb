Partial Class frmPerfStats
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

    Private PSTATS_TB_WIDTH As Integer = 450
    Private PSTATS_TB_HEIGHT As Integer = 500
    Private PSTATS_TB_SIZE As Size = New Size(PSTATS_TB_WIDTH, PSTATS_TB_HEIGHT)
    Private PSTATS_BUTTON_SIZE As Size = New Size(150, 100)
    Private PSTATS_FRM_SIZE As Size = New Size(PSTATS_TB_WIDTH + 200, PSTATS_TB_HEIGHT + 25)

    Private parent_frm

    Friend perf_details As TextBox
    Friend save_perf As Button
    Friend save_file As SaveFileDialog

    Public Sub New(frm As Form)
        Me.generate(frm)
    End Sub

    Private Sub generate(frm As Form)
        Me.perf_details = New TextBox
        Me.save_perf = New Button
        Me.save_file = New SaveFileDialog

        'perf_details
        Me.perf_details.Size = PSTATS_TB_SIZE
        Me.perf_details.Font = New Font("Courier New", DefaultFont.Size)
        Me.perf_details.Location = New Point(0, 0)
        Me.perf_details.Multiline = True
        Me.perf_details.Text = statistics_string()
        Me.perf_details.Name = "perf_details"

        'save_perf
        Me.save_perf.Size = PSTATS_BUTTON_SIZE
        Me.save_perf.Location = New Point(PSTATS_TB_WIDTH + 25, 25)
        Me.save_perf.Text = "Save Performance"
        Me.save_perf.Name = "save_perf"
        AddHandler Me.save_perf.Click, AddressOf Me.save_perf_click

        'save_dialog
        Me.save_file.FileName = "ringingcalculator.txt"
        Me.save_file.Filter = "txt files|*.txt|All files|*.*"
        Me.save_file.Title = "Save Performance"
        Me.save_file.DefaultExt = "txt"

        'frm
        Me.Controls.Add(Me.perf_details)
        Me.Controls.Add(Me.save_perf)

        Me.Text = "Performance Statistics"
        Me.Name = "frmPerfStats"
        Me.ClientSize = PSTATS_FRM_SIZE
        Me.parent_frm = frm
        frm.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        Me.Show()
    End Sub

    Private Sub save_perf_click(o As Object, e As EventArgs)
        If Me.save_file.ShowDialog() = DialogResult.OK Then
            Saving.save_statistics(Me.save_file.FileName)
            MsgBox("Performance saved.",, "File Saved")
        Else
            MsgBox("Performance not saved.",, "Save Performance")
        End If
    End Sub

End Class
