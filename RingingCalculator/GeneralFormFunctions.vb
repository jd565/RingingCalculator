﻿Module GeneralFormFunctions
    Const DEFAULT_FONT_SIZE As Integer = 8
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

        If TypeOf ctl.Parent Is Form Then
            frm = ctl.Parent
            frm.Close()
        End If
    End Sub

    ' Looks through the tree of forms to find one that is the same type as the one we are searching for.
    ' Returns the form, or nothing.
    Public Function find_form(t As Type, Optional parent As Form = Nothing) As Form
        Dim ret As Form
        If parent Is Nothing Then
            parent = frmPerf
        End If
        If parent.GetType = t Then
            Return parent
        End If
        For Each frm In parent.OwnedForms
            ret = find_form(t, frm)
            If ret IsNot Nothing Then
                Return ret
            End If
        Next
        Return Nothing
    End Function

End Module
