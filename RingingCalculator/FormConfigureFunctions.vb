Partial Class frmConfigure
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

    Friend label As Label
    Friend cancel As Button
    Public input As Input

    Public Sub New(ByRef i As Input)
        Me.input = i
    End Sub

    'Function to create a form to show while waiting for the config to be set
    Public Sub generate(parent As Form)
        Me.label = New Label
        Me.cancel = New Button

        Me.label.Text = "Please ring bell or press switch"
        Me.label.Size = New Size(70, 40)
        Me.label.Location = New Point(35, 30)

        Me.cancel.Text = "Cancel"
        Me.cancel.Size = New Size(70, 40)
        Me.cancel.Location = New Point(35, 70)
        AddHandler cancel.Click, AddressOf close_parent_form

        Me.Text = "Configure"
        Me.Name = "frmConfigure"
        Me.Font = DEFAULT_FONT
        Me.ClientSize = New Size(140, 140)
        Me.Controls.Add(cancel)
        Me.Controls.Add(label)
        parent.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf Me.cancel_self

        parent.Hide()
        Me.Show()
    End Sub

    ' Function to cancel the config waiting form and show the previous form
    ' We first set all bells to be not configurable,
    ' then delete this form.
    Private Sub cancel_self(frm As Form, e As EventArgs)
        Me.input.can_be_configured = False
        dispose_of_form(frm, e)
    End Sub

    ' Function to be called when a configure event is hit on a serial port.
    ' We expect the waiting form to be open, but check this.
    ' If the waiting form is open, then close it, as the bell has been configured.
    Public Sub item_has_been_configured()
        Me.Close()
    End Sub

End Class
