Module GeneralFormFunctions

    ' Function to dispose of a Form that is no longer being used, e.g. if it has been closed.
    ' Make sure that if we are being disposed of, the owner of this form is being shown,
    ' and that it is removed from the owned list.
    Public Sub dispose_of_form(frm As Form, e As EventArgs)
        frm.Owner.Show()
        frm.Owner.RemoveOwnedForm(frm)
        frm.Dispose()
    End Sub

    ' Function to close the form that the calling object is a child of.
    Public Sub close_parent_form(ctl As Control, e As EventArgs)
        Dim frm As Form

        If (ctl.Parent.GetType() Is GetType(Form)) Then
            frm = ctl.Parent
            frm.Close()
        End If
    End Sub

End Module
