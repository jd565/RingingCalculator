Module ReportingFunctions

    ' Function to print a certain row of the ringing.
    ' It does this by creating a list of the times of each change,
    ' and then sorting the list before returning the row
    Public Function print_change(change_id As Integer) As String
        Dim change As New List(Of ChangeTime)
        Dim row_to_print As String = ""

        For Each bell In GlobalVariables.bells
            change.Add(bell.change_times(change_id - 1))
        Next

        ' we now have a list of the time of every bell for this change.
        change.Sort()

        ' Now go through this list and create a string.
        ' We do have to convert each bell number to the string representation
        For Each change_time In change
            row_to_print += bell_number_to_string(change_time.bell)
        Next

        Return row_to_print
    End Function

    ' Function to return the string representation of the bell number
    Private Function bell_number_to_string(bell_number As Integer) As String
        If bell_number < 10 Then Return bell_number.ToString()
        If bell_number = 10 Then Return "0"
        If bell_number = 11 Then Return "E"
        If bell_number = 12 Then Return "T"
        Return "A" - 13 + bell_number
    End Function

End Module
