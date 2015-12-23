Public Class Method
    Public place_notation As PlaceNotation
    Public bells As Integer
    Public rows As List(Of Row)

    Public Sub New(pn As String, bells As Integer)
        Me.place_notation = New PlaceNotation(pn)
        Me.bells = bells
    End Sub

    ' Function to generate the rows of this method, based on the notation
    Public Sub generate()
        Me.rows = New List(Of Row)

        ' Add rounds as the first row
        Me.create_first_row()

        ' We use a change time of 1s between bells, and use this to indicate
        ' their position within a row
        ' Now generate rows until we get back to rounds
        Me.generate_next_row()

        While Not Me.rows.Last.is_this_rounds()
            Me.generate_next_row()
        End While
    End Sub

    ' Function to generate the next row of this method
    Private Sub generate_next_row()
        Dim hash As List(Of Integer)
        Dim index As Integer
        Dim next_row As New Row
        Dim change_time As ChangeTime

        ' Get the index of the notation to use now.
        ' The notation to go from the first row to the second is at index 0
        index = (Me.rows.Count - 1) Mod Me.place_notation.notation.Count
        hash = Me.place_notation.notation(index).change_hash(Me.bells)

        For ii As Integer = 0 To hash.Count - 1
            change_time = New ChangeTime(Me.rows.Last.bells(ii).bell,
                                         Me.rows.Count + 1,
                                         hash(ii))
            next_row.add(change_time)
        Next
        next_row.sort()
        Me.rows.Add(next_row)
    End Sub

    ' Function to put rounds as the first row
    Private Sub create_first_row()
        Dim row As New Row
        Dim ct As ChangeTime
        For ii As Integer = 1 To bells
            ct = New ChangeTime(ii, 1, ii)
            row.add(ct)
        Next
        Me.rows.Add(row)
    End Sub
End Class
