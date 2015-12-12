Module ReportingFunctions

    ' Function to create a single row, as a list of change times.
    Public Function get_row(change_id As Integer) As Row
        Dim change As New row

        ' We may be part way through a change when this function is called.
        ' If we are we will try and find items in a list above the max of the list,
        ' so handle exceptions here
        For Each bell In GlobalVariables.bells
            Try
                change.Add(bell.change_times(change_id))
            Catch
                Exit For
            End Try
        Next

        ' we now have a list of the time of every bell for this change.
        change.Sort()

        Return change
    End Function

    ' Function to convert a time in seconds to a string of hours and minutes and seconds
    Public Function time_to_string(time As Integer) As String
        Dim hours As String
        Dim minutes As String
        Dim seconds As String

        hours = (time \ 3600000).ToString
        minutes = ((time - (time \ 3600000) * 3600000) \ 60000).ToString
        seconds = ((time Mod 60000) \ 1000).ToString
        Return hours + ":" + minutes + ":" + seconds

    End Function

    ' Function to find the location of a bell in a certain row
    ' If the bell is not found in the list (unlikely) the function returns 0
    Public Function get_bell_location_in_row(change_id As Integer, bell_number As Integer) As Integer
        Dim change As List(Of ChangeTime) = get_row(change_id)

        For ii As Integer = 1 To change.Count
            If change(ii - 1).bell = bell_number Then
                Return ii
            End If
        Next
        Return 0
    End Function

    ' Function to print a certain row of the ringing.
    ' It does this by creating a list of the times of each change,
    ' and then sorting the list before returning the row
    Public Function print_change(change_id As Integer) As String
        Dim row_to_print As String = ""
        Dim change As List(Of ChangeTime) = get_row(change_id)

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
        Return Convert.ToChar(Convert.ToInt32(CChar("A")) - 13 + bell_number).ToString()
    End Function

End Module
