Module ReportingFunctions

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

    ' Function to be called once all bells have rung a certain row.
    ' This happens for every row, and the sorted row is put into the rows
    ' list.
    Public Sub row_is_full(change_id As Integer)
        Console.WriteLine("Row is full")
        Statistics.rows(change_id).sort()
        ' If the switch has been pressed to stop running then stop recording here
        If Not GlobalVariables.switch.isRunning Then
            GlobalVariables.recording = False
            update_statistics(change_id)
        End If
        If is_this_lead_end(change_id) Then
            update_statistics(change_id)
        End If
    End Sub

    ' Function for checking if this is a lead end
    Public Function is_this_lead_end(change_id As Integer) As Boolean
        Return ((change_id + 1) Mod GlobalVariables.changes_per_lead = 0)
    End Function

    Private Sub update_statistics(change_id As Integer)
        Console.WriteLine("Update lead end stats")

        Statistics.time = Statistics.rows(change_id).time
        Statistics.peal_speed = Statistics.rows(change_id).time * GlobalVariables.changes_per_peal / (change_id + 1)
        Statistics.lead_end_row_value.Text = Statistics.rows(change_id).print
        Statistics.leads_value.Text = Statistics.leads.ToString
        Statistics.changes_value.Text = Statistics.changes.ToString
        Statistics.time_value.Text = time_to_string(Statistics.time)
        Statistics.peal_speed_value.Text = time_to_string(Statistics.peal_speed)
        If change_id - GlobalVariables.changes_per_course >= 0 Then
            Statistics.last_course_time_value.Text = time_to_string(Statistics.rows(change_id).time -
                                                                Statistics.rows(change_id - GlobalVariables.changes_per_course).time)
            Statistics.last_course_peal_speed_value.Text = time_to_string(Val(Statistics.last_course_time_value.Text) *
                                                                              GlobalVariables.changes_per_peal /
                                                                              GlobalVariables.changes_per_course)
        End If
    End Sub

    ' Function to return the string representation of the bell number
    Public Function bell_number_to_string(bell_number As Integer) As String
        If bell_number < 10 Then Return bell_number.ToString()
        If bell_number = 10 Then Return "0"
        If bell_number = 11 Then Return "E"
        If bell_number = 12 Then Return "T"
        If bell_number <= 16 Then Return Convert.ToChar(Convert.ToInt32(CChar("A")) - 13 + bell_number).ToString()
        Return Convert.ToChar(Convert.ToInt32(CChar("F")) - 17 + bell_number).ToString()
    End Function

    ' Function to return the integer representation of the bell based on the string val
    Public Function bell_string_to_number(str As String) As Integer
        If (str(0) >= CChar("1") And str(0) <= CChar("9")) Then Return Val(str(0))
        If str(0) = CChar("0") Then Return 10
        If str(0) = CChar("E") Then Return 11
        If str(0) = CChar("T") Then Return 12
        If (str(0) >= CChar("A") And str(0) <= CChar("D")) Then Return (Convert.ToInt32(str(0)) - Convert.ToInt32(CChar("A")) + 13)
        Return (Convert.ToInt32(str(0)) - Convert.ToInt32(CChar("F")) + 17)
    End Function

End Module
