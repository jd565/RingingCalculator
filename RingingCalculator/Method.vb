Public Class Method
    Public place_notation As PlaceNotation
    Public bells As Integer
    Public rows As List(Of Row)
    Public changes_per_lead As Integer

    Public Sub New(pn As String, bells As Integer)
        Me.place_notation = New PlaceNotation(pn, bells)
        Me.bells = bells
        Me.changes_per_lead = Me.place_notation.main_block.Count + Me.place_notation.lead_end.Count
    End Sub

    ' Function to generate the rows of this method, based on the notation
    Public Sub generate()
        Me.rows = New List(Of Row)
        Me.create_first_row()

        ' We use a change time of 1s between bells, and use this to indicate
        ' their position within a row
        ' Now generate rows until we get back to rounds
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
        ' The notation to go from the first row to the second is at index 1
        index = (Me.rows.Count) Mod Me.changes_per_lead
        If index >= Me.place_notation.main_block.Count Then
            index -= Me.place_notation.main_block.Count
            hash = Me.place_notation.lead_end(index).change_hash(Me.bells)
        Else
            hash = Me.place_notation.main_block(index).change_hash(Me.bells)
        End If

        For ii As Integer = 0 To hash.Count - 1
            change_time = New ChangeTime(Me.rows.Last.bells(ii).bell,
                                         Me.rows.Count + 1,
                                         New DateTime(hash(ii)))
            next_row.add(change_time)
        Next
        next_row.sort()
        Me.rows.Add(next_row)
    End Sub

    ' Function to put rounds as the first row
    Private Sub create_first_row()
        Dim hash As List(Of Integer)
        Dim index As Integer
        Dim next_row As New Row
        Dim change_time As ChangeTime

        index = 0
        hash = Me.place_notation.main_block(index).change_hash(Me.bells)

        For ii As Integer = 0 To hash.Count - 1
            change_time = New ChangeTime(ii + 1,
                                         1,
                                         New DateTime(hash(ii)))
            next_row.add(change_time)
        Next
        next_row.sort()
        Me.rows.Add(next_row)
    End Sub
End Class
