Public Class frmMain
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If Not GlobalVariables.COM_ports_configured Then
            GlobalVariables.generate_COM_ports(Val(Me.txtCOMs.Text))
            generate_frmCOMs()
        Else
            GlobalVariables.generate_bells(Val(Me.txtBells.Text))
            generate_frmbells()
        End If
    End Sub
End Class
