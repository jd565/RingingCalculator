Partial Class frmNewConf
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

    Friend txt_bells As TextBox
    Friend txt_coms As TextBox
    Friend lbl_bells As Label
    Friend lbl_coms As Label
    Friend btn_ok As Button

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub generate(parent As Form)
        Me.txt_bells = New System.Windows.Forms.TextBox()
        Me.txt_coms = New System.Windows.Forms.TextBox()
        Me.lbl_bells = New System.Windows.Forms.Label()
        Me.lbl_coms = New System.Windows.Forms.Label()
        Me.btn_ok = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtBells
        '
        Me.txt_bells.Location = New System.Drawing.Point(180, 31)
        Me.txt_bells.Name = "txt_bells"
        Me.txt_bells.Size = New System.Drawing.Size(100, 26)
        Me.txt_bells.TabIndex = 0
        '
        'txtCOMs
        '
        Me.txt_coms.Location = New System.Drawing.Point(180, 95)
        Me.txt_coms.Name = "txt_coms"
        Me.txt_coms.Size = New System.Drawing.Size(100, 26)
        Me.txt_coms.TabIndex = 1
        '
        'Label1
        '
        Me.lbl_bells.AutoSize = True
        Me.lbl_bells.Location = New System.Drawing.Point(12, 31)
        Me.lbl_bells.Name = "lbl_bells"
        Me.lbl_bells.Size = New System.Drawing.Size(110, 24)
        Me.lbl_bells.TabIndex = 2
        Me.lbl_bells.Text = "No. of Bells:"
        '
        'Label2
        '
        Me.lbl_coms.AutoSize = True
        Me.lbl_coms.Location = New System.Drawing.Point(12, 73)
        Me.lbl_coms.Name = "lbl_coms"
        Me.lbl_coms.Size = New System.Drawing.Size(150, 48)
        Me.lbl_coms.TabIndex = 3
        Me.lbl_coms.Text = "No. of COM" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ports to be used:"
        '
        'btnGenerate
        '
        Me.btn_ok.Location = New System.Drawing.Point(180, 159)
        Me.btn_ok.Name = "btn_ok"
        Me.btn_ok.Size = New System.Drawing.Size(100, 47)
        Me.btn_ok.TabIndex = 4
        Me.btn_ok.Text = "OK"
        Me.btn_ok.UseVisualStyleBackColor = True
        AddHandler Me.btn_ok.Click, AddressOf Me.btn_ok_click
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(439, 329)
        Me.Controls.Add(Me.btn_ok)
        Me.Controls.Add(Me.lbl_coms)
        Me.Controls.Add(Me.lbl_bells)
        Me.Controls.Add(Me.txt_coms)
        Me.Controls.Add(Me.txt_bells)
        Me.Font = DEFAULT_FONT
        Me.Name = "frmNewConf"
        Me.Text = "New Configuration"
        parent.AddOwnedForm(Me)
        Me.ResumeLayout(False)
        Me.PerformLayout()
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        parent.Hide()
        Me.Show()

    End Sub

    Private Sub btn_ok_click(sender As Button, e As EventArgs)
        Dim frm As New frmCom
        GlobalVariables.generate_COM_ports(Val(Me.txt_coms.Text))
        frm.generate(Me)
    End Sub

    Public Sub generate_frm_bells()
        Dim frm As New frmBells
        GlobalVariables.generate_bells(Val(Me.txt_bells.Text))
        frm.generate(Me)
    End Sub

End Class
