Public Class frmMain
    Private Sub btnGenerate_Click(sender As Button, e As EventArgs) Handles btnGenerate.Click
        If Not GlobalVariables.COM_ports_configured Then
            GlobalVariables.generate_COM_ports(Val(Me.txtCOMs.Text))
            generate_frmCOMs(sender.Parent)
        Else
            GlobalVariables.generate_bells(Val(Me.txtBells.Text))
            generate_frmBells(sender.Parent)
        End If
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        run_tests(sender.parent)
    End Sub

    Private Sub txtCOMs_TextChanged(sender As Object, e As EventArgs) Handles txtCOMs.TextChanged


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnLoadConfig.Click
        load_config()
    End Sub
End Class
