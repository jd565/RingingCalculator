Partial Class frmMethod
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

    Private pn_text As String = "Place Notation"
    Private comp_text As String = "Composition." & vbCrLf & "Please enter composition in aligned rows, e.g." & vbCrLf & "W M  H" & vbCrLf & "- 2- 3" & vbCrLf & "  s2 2"

    Friend place_notation As TextBox
    Friend bells_text As TextBox
    Friend composition As TextBox
    Friend gen_method As Button

    Public Sub New(parent As Form)
        Me.generate(parent)
    End Sub

    Private Sub generate(parent As Form)
        Me.place_notation = New TextBox
        Me.bells_text = New TextBox
        Me.composition = New TextBox
        Me.gen_method = New Button

        'place_notation
        Me.place_notation.Name = "place_notation"
        Me.place_notation.Text = Me.pn_text
        Me.place_notation.Size = New Size(200, 30)
        Me.place_notation.Location = New Point(20, 20)

        'bells_text
        Me.bells_text.Name = "bells_text"
        Me.bells_text.Text = "Number of bells"
        Me.bells_text.Size = New Size(200, 30)
        Me.bells_text.Location = New Point(20, 70)

        'composition
        Me.composition.Name = "composition"
        Me.composition.Text = Me.comp_text
        Me.composition.Size = New Size(200, 150)
        Me.composition.Location = New Point(20, 120)
        Me.composition.Multiline = True

        'gen_method
        Me.gen_method.Name = "gen_method"
        Me.gen_method.Text = "Generate"
        Me.gen_method.Size = New Size(200, 100)
        Me.gen_method.Location = New Point(20, 290)
        AddHandler Me.gen_method.Click, AddressOf Me.gen_method_click

        'form
        Me.Controls.Add(Me.place_notation)
        Me.Controls.Add(Me.bells_text)
        Me.Controls.Add(Me.composition)
        Me.Controls.Add(Me.gen_method)
        Me.Text = "Method"
        Me.Name = "frmPerf"
        Me.ClientSize = New Size(250, 400)
        Me.Font = New Font("Courier New", DEFAULT_FONT.Size, DEFAULT_FONT.Style)
        parent.AddOwnedForm(Me)
        AddHandler Me.FormClosing, AddressOf dispose_of_form

        parent.Hide()
        Me.Show()
    End Sub

    Private Sub gen_method_click(o As Object, e As EventArgs)
        Dim method As Method
        Dim frm As frmPerfStats
        method = New Method(Me.place_notation.Text, Val(Me.bells_text.Text))
        If Not Me.composition.Text.Equals(Me.comp_text) Then
            method.add_composition(Me.composition.Text)
        End If
        method.generate()
        frm = New frmPerfStats(Me, method)
    End Sub

End Class
