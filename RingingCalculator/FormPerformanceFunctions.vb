Partial Class frmPerf
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

    Const PERF_FIELD_GAP As Integer = 15
    Const PERF_FIELD_HEIGHT As Integer = 25
    Const PERF_FIELD_WIDTH As Integer = 120
    Const PERF_MENU_HEIGHT As Integer = 25
    Const PERF_BUTTON_HEIGHT As Integer = PERF_FIELD_GAP + 2 * PERF_FIELD_HEIGHT
    Private PERF_LABEL_SIZE As New Size(PERF_FIELD_WIDTH, PERF_FIELD_HEIGHT)
    Private PERF_UPDOWN_SIZE As New Size(PERF_FIELD_WIDTH, PERF_FIELD_HEIGHT)
    Private PERF_GENERAL_SIZE As New Size(PERF_FIELD_WIDTH, PERF_FIELD_HEIGHT)
    Private PERF_BUTTON_SIZE As New Size(PERF_FIELD_WIDTH, PERF_BUTTON_HEIGHT)

    Private config_filename As String

    Friend changes_per_lead As TextBox
    Friend lbl_changes_per_lead As Label
    Friend leads_per_course As TextBox
    Friend lbl_leads_per_course As Label
    Friend changes_per_Peal As TextBox
    Friend lbl_changes_per_Peal As Label
    Friend btn_start As Button
    Friend load_dialog As OpenFileDialog
    Friend save_file As SaveFileDialog
    Friend main_menu As MenuStrip
    Friend config_button As ToolStripDropDownButton
    Friend config_options As ToolStripDropDown
    Friend load_config As ToolStripButton
    Friend new_config As ToolStripButton
    Friend edit_config As ToolStripButton
#If DEBUG Then
    Friend btn_test As ToolStripButton
#End If

    Public Sub New()
        If Not GlobalVariables.statistics_init Then
            Statistics.init()
        End If
        Me.InitializeComponent()
    End Sub

    ' Function to return the coordinate of a point on the grid
    Private Function coordinate(x As Integer, y As Integer) As Point
        Return New Point(PERF_FIELD_GAP + x * (PERF_FIELD_GAP + PERF_FIELD_WIDTH),
                         PERF_FIELD_GAP + PERF_MENU_HEIGHT + y * (PERF_FIELD_GAP + PERF_FIELD_HEIGHT))
    End Function

    ' Function to generate the performance details (this is the main form)
    Private Sub InitializeComponent()
        Me.changes_per_lead = New TextBox
        Me.lbl_changes_per_lead = New Label
        Me.leads_per_course = New TextBox
        Me.lbl_leads_per_course = New Label
        Me.changes_per_Peal = New TextBox
        Me.lbl_changes_per_Peal = New Label
        Me.btn_start = New Button
        Me.load_dialog = New OpenFileDialog
        Me.save_file = New SaveFileDialog
        Me.main_menu = New MenuStrip
        Me.config_button = New ToolStripDropDownButton
        Me.config_options = New ToolStripDropDown
        Me.load_config = New ToolStripButton
        Me.new_config = New ToolStripButton
        Me.edit_config = New ToolStripButton

        ' Changes per lead
        lbl_changes_per_lead.Text = "Changes per lead:"
        lbl_changes_per_lead.Size = PERF_LABEL_SIZE
        lbl_changes_per_lead.Location = coordinate(0, 0)
        changes_per_lead.Text = GlobalVariables.changes_per_lead.ToString
        changes_per_lead.Size = PERF_UPDOWN_SIZE
        changes_per_lead.Location = coordinate(1, 0)
        AddHandler changes_per_lead.TextChanged, AddressOf changes_per_lead_changed

        ' leads per course
        lbl_leads_per_course.Text = "Leads per course:"
        lbl_leads_per_course.Size = PERF_LABEL_SIZE
        lbl_leads_per_course.Location = coordinate(0, 1)
        leads_per_course.Text = GlobalVariables.leads_per_course.ToString
        leads_per_course.Size = PERF_UPDOWN_SIZE
        leads_per_course.Location = coordinate(1, 1)
        AddHandler leads_per_course.TextChanged, AddressOf leads_per_course_changed

        ' changes per peal
        lbl_changes_per_Peal.Text = "Peal length:"
        lbl_changes_per_Peal.Size = PERF_LABEL_SIZE
        lbl_changes_per_Peal.Location = coordinate(0, 2)
        changes_per_Peal.Text = GlobalVariables.changes_per_peal.ToString
        changes_per_Peal.Size = PERF_UPDOWN_SIZE
        changes_per_Peal.Location = coordinate(1, 2)
        AddHandler changes_per_Peal.TextChanged, AddressOf changes_per_peal_changed

        'btn_start
        Me.btn_start.Name = "btn_start"
        Me.btn_start.Text = "Config present" & vbCrLf & "Begin performance"
        Me.btn_start.Size = PERF_BUTTON_SIZE
        Me.btn_start.Location = coordinate(1, 3)
        Me.btn_start.Enabled = False
        Me.btn_start.Visible = False
        AddHandler Me.btn_start.Click, AddressOf Me.start_perf

        'load_dialog
        load_dialog.DefaultExt = "conf"
        load_dialog.FileName = "ringingcalculator.conf"
        load_dialog.Filter = "Conf files|*.conf|All files|*.*"
        load_dialog.Title = "Configuration File"

        'save_dialog
        Me.save_file.FileName = "ringingcalculator.conf"
        Me.save_file.Filter = "Conf files|*.conf|All files|*.*"
        Me.save_file.Title = "Configuration File"
        Me.save_file.DefaultExt = "conf"

        'load_config
        load_config.Name = "load_config"
        load_config.Text = "Load Config"
        AddHandler load_config.Click, AddressOf load_config_click

        'new_config
        Me.new_config.Name = "new_config"
        Me.new_config.Text = "New Config"
        AddHandler new_config.Click, AddressOf new_config_click

        'edit_config
        Me.edit_config.Name = "edit_config"
        Me.edit_config.Text = "Edit Config"
        AddHandler Me.edit_config.Click, AddressOf Me.edit_config_click

        'Menu
        Me.config_button.Text = "Configuration"
        Me.config_button.DropDown = Me.config_options
        Me.config_button.DropDownDirection = ToolStripDropDownDirection.BelowRight
        Me.config_button.ShowDropDownArrow = False
        Me.config_options.Items.AddRange(New ToolStripItem() {Me.new_config, Me.load_config, Me.edit_config})
        Me.main_menu.Items.Add(Me.config_button)
        Me.main_menu.Height = PERF_MENU_HEIGHT

#If DEBUG Then
        Me.add_testing_button()
#End If

        Me.Controls.Add(lbl_changes_per_lead)
        Me.Controls.Add(changes_per_lead)
        Me.Controls.Add(lbl_changes_per_Peal)
        Me.Controls.Add(changes_per_Peal)
        Me.Controls.Add(lbl_leads_per_course)
        Me.Controls.Add(leads_per_course)
        Me.Controls.Add(Me.btn_start)
        Me.Controls.Add(main_menu)

        Me.Text = "New Performance"
        Me.Name = "frmPerf"
        Me.Font = DEFAULT_FONT

        Me.Show()
    End Sub

    Private Sub changes_per_lead_changed(txt As TextBox, e As EventArgs)
        If GlobalVariables.switch.is_running Then
            Console.WriteLine("Tried to change changes per lead, but we are running")
            txt.Text = GlobalVariables.changes_per_lead
        Else
            GlobalVariables.changes_per_lead = Val(txt.Text)
            GlobalVariables.update_changes_per_course()
        End If
    End Sub

    Private Sub leads_per_course_changed(txt As TextBox, e As EventArgs)
        If GlobalVariables.switch.is_running Then
            Console.WriteLine("Tried to change leads per course, but we are running")
            txt.Text = GlobalVariables.leads_per_course
        Else
            GlobalVariables.leads_per_course = Val(txt.Text)
            GlobalVariables.update_changes_per_course()
        End If
    End Sub

    Private Sub changes_per_peal_changed(txt As TextBox, e As EventArgs)
        If GlobalVariables.switch.is_running Then
            Console.WriteLine("Tried to change peal length, but we are running")
            txt.Text = GlobalVariables.changes_per_peal
        Else
            GlobalVariables.changes_per_peal = Val(txt.Text)
        End If
    End Sub

    Private Sub load_config_click(sender As Object, e As EventArgs)
        If Me.load_dialog.ShowDialog() = DialogResult.OK Then
            Me.config_filename = Me.load_dialog.FileName
            Saving.load_config(Me.config_filename)
        End If
        Me.maybe_show_btn_start()
    End Sub

    Private Sub new_config_click(sender As Object, e As EventArgs)
        Dim frm As New frmNewConf
        GlobalVariables.reset()
        Me.maybe_show_btn_start()
        frm.generate(Me)
    End Sub

    Private Sub edit_config_click(sender As Object, e As EventArgs)
        If GlobalVariables.config_loaded = False Then Exit Sub
        Dim frm As New frmBells
        frm.generate(Me)
    End Sub

    Public Sub save_config()
        If GlobalVariables.config_loaded Then
            ' This has come because we have edited the config.
            Saving.save_config(Me.config_filename)
            MsgBox("File successfully edited.",, "File saved")
        Else
            If Me.save_file.ShowDialog() = DialogResult.OK Then
                Me.config_filename = Me.save_file.FileName
                Saving.save_config(Me.config_filename)
                MsgBox("Configuration saved to " & Me.config_filename & ".",, "File Saved")
            Else
                MsgBox("Config not saved.",, "Configuration")
            End If
            GlobalVariables.config_loaded = True
        End If
        Me.close_all_children()
        Me.maybe_show_btn_start()
    End Sub

    Private Sub close_all_children()
        For Each form In Me.OwnedForms
            form.Close()
        Next
        Me.Show()
    End Sub

    Private Sub start_perf(b As Button, e As EventArgs)
        GlobalVariables.switch.start_running()
    End Sub

    Private Sub maybe_show_btn_start()
        Console.WriteLine("maybe_show_btn_start, config_loaded: {0}", GlobalVariables.config_loaded)
        Me.btn_start.Enabled = GlobalVariables.config_loaded
        Me.btn_start.Visible = GlobalVariables.config_loaded
    End Sub

    Private Sub add_testing_button()
        Me.btn_test = New ToolStripButton
        Me.btn_test.Name = "testing"
        Me.btn_test.Text = "Run Tests"
        AddHandler Me.btn_test.Click, AddressOf Me.run_tests

        Me.main_menu.Items.Add(Me.btn_test)
    End Sub

    Private Sub run_tests(o As Object, e As EventArgs)
        Testing.run_tests(Me)
    End Sub

End Class
