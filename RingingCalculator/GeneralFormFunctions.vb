Module GeneralFormFunctions
    Const DEFAULT_FONT_SIZE As Integer = 12
    Const DEFAULT_FONT_STYLE As FontStyle = FontStyle.Regular
    Const DEFAULT_FONT_FONT As String = "Microsft Sans Serif"
    Public DEFAULT_FONT As New Font(DEFAULT_FONT_FONT, DEFAULT_FONT_SIZE, DEFAULT_FONT_STYLE)

    ' Function to dispose of a Form that is no longer being used, e.g. if it has been closed.
    ' Make sure that if we are being disposed of, the owner of this form is being shown,
    ' and that it is removed from the owned list.
    Public Sub dispose_of_form(frm As Form, e As EventArgs)
        frm.Owner.Show()
        frm.Owner.RemoveOwnedForm(frm)
        For Each form In frm.OwnedForms
            form.Close()
        Next
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

    Public Function find_form(name As String, Optional parent As Form = Nothing) As Form
        Dim ret As Form
        If parent Is Nothing Then
            parent = frmMain
        End If
        If name.Equals(parent.Name) Then
            Return parent
        End If
        For Each frm In frmMain.OwnedForms
            ret = find_form(name, frm)
            If ret IsNot Nothing Then
                Return ret
            End If
        Next
        Return Nothing
    End Function

End Module
