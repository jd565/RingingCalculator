<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txtBells = New System.Windows.Forms.TextBox()
        Me.txtCOMs = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.btnTest = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtBells
        '
        Me.txtBells.Location = New System.Drawing.Point(180, 31)
        Me.txtBells.Name = "txtBells"
        Me.txtBells.Size = New System.Drawing.Size(100, 26)
        Me.txtBells.TabIndex = 0
        '
        'txtCOMs
        '
        Me.txtCOMs.Location = New System.Drawing.Point(180, 95)
        Me.txtCOMs.Name = "txtCOMs"
        Me.txtCOMs.Size = New System.Drawing.Size(100, 26)
        Me.txtCOMs.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!)
        Me.Label1.Location = New System.Drawing.Point(12, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(110, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "No. of Bells:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!)
        Me.Label2.Location = New System.Drawing.Point(12, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(150, 48)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "No. of COM" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ports to be used:"
        '
        'btnGenerate
        '
        Me.btnGenerate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!)
        Me.btnGenerate.Location = New System.Drawing.Point(180, 159)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(100, 47)
        Me.btnGenerate.TabIndex = 4
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'btnTest
        '
        Me.btnTest.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!)
        Me.btnTest.Location = New System.Drawing.Point(331, 232)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(75, 72)
        Me.btnTest.TabIndex = 5
        Me.btnTest.Text = "Run tests"
        Me.btnTest.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(439, 329)
        Me.Controls.Add(Me.btnTest)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCOMs)
        Me.Controls.Add(Me.txtBells)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.Name = "frmMain"
        Me.Text = "Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtBells As TextBox
    Friend WithEvents txtCOMs As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnGenerate As Button
    Friend WithEvents btnTest As Button
End Class
