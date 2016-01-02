Module ReportingFunctions

    ' Function to be called once all bells have rung a certain row.
    ' This happens for every row, and the sorted row is put into the rows
    ' list.
    Public Sub row_is_full(change_id As Integer)
        Console.WriteLine("Row is full")
        Statistics.rows(change_id).sort()
        update_changes(change_id)
        ' If the switch has been pressed to stop running then stop recording here
        If Not GlobalVariables.switch.is_running Then
            GlobalVariables.recording = False
            update_statistics(change_id)
        End If
        If is_this_lead_end(change_id) Then
            update_statistics(change_id)
        End If
    End Sub

    ' Function for checking if this is a lead end
    Public Function is_this_lead_end(change_id As Integer) As Boolean
        If (change_id + 1) Mod GlobalVariables.changes_per_lead = 0 Then Return True
        Return False
    End Function

    ' Function to find how many changes have been rung in the last minute
    Private Function get_changes_last_minute(change_id As Integer) As Integer
        Dim curr_time As TimeSpan
        Dim check_time As TimeSpan
        Dim change_to_check As Integer = 30

        curr_time = Statistics.rows(change_id).time
        check_time = curr_time - TimeSpan.FromMinutes(1)

        ' Check that we have been running more than a minute
        If check_time.CompareTo(TimeSpan.Zero) < 0 Then Return 0

        ' The general case is to start searching 30 changes before this one.
        ' Check that we can look 30 back
        If change_id < 30 Then
            change_to_check = change_id
        End If

        ' First move earlier in the list until we have found a change that happened
        ' More than a minute ago.
        While Statistics.rows(change_id - change_to_check).time.CompareTo(check_time) > 0
            change_to_check += 1
            If change_to_check > change_id Then Return 0
        End While

        ' Now look later until we have found a change that happened less than a minute ago
        ' We do this this way round as this means that the change we are looking at is the last
        ' change to happen inside a minute.
        While Statistics.rows(change_id - change_to_check).time.CompareTo(check_time) < 0
            change_to_check -= 1
            If change_to_check < 0 Then Return 0
        End While

        ' Return + 1 to create the right number of changes (think of the case where the current change
        ' Was the only one to occur in the last minute)
        Return (change_to_check + 1)

    End Function

    Private Sub update_changes(change_id As Integer)
        Dim clm As Integer

        Statistics.changes += 1
        Statistics.changes_value.Text = Statistics.changes.ToString
        Statistics.time = Statistics.rows(change_id).time
        Statistics.time_value.Text = Statistics.time.ToString(GlobalVariables.time_string_format)
        Statistics.changes_per_minute_value.Text =
            ((change_id + 1) / (Statistics.rows(change_id).time.TotalMinutes)).ToString(GlobalVariables.cpm_string_format)
        clm = get_changes_last_minute(change_id)
        ' get_changes_last_minute returns 0 if a minute has not elapsed, or the number of changes can't be found
        If clm > 0 Then
            Statistics.last_minute_changes_value.Text = get_changes_last_minute(change_id).ToString()
        End If
    End Sub

    Private Sub update_statistics(change_id As Integer)
        Console.WriteLine("Update lead end stats")

        Statistics.peal_speed = New TimeSpan(Statistics.rows(change_id).time.Ticks * GlobalVariables.changes_per_peal / (change_id + 1))
        Statistics.lead_end_row_value.Text = Statistics.rows(change_id).print
        Statistics.leads_value.Text = Statistics.leads.ToString
        Statistics.peal_speed_value.Text = Statistics.peal_speed.ToString(GlobalVariables.time_string_format)
        maybe_update_course_statistics(change_id)
    End Sub

    ' Function to update the course en statistics if this is a course end
    Private Sub maybe_update_course_statistics(change_id As Integer)

        ' If this is not a course end then drop out here
        If (change_id + 1) Mod GlobalVariables.changes_per_course <> 0 Then Exit Sub

        Dim course_speed As TimeSpan
        Dim course_peal_speed As TimeSpan
        course_speed = Statistics.rows(change_id).time -
                       Statistics.rows(change_id + 1 - GlobalVariables.changes_per_course).time
        course_peal_speed = New TimeSpan(course_speed.Ticks * GlobalVariables.changes_per_peal / GlobalVariables.changes_per_course)
        Statistics.last_course_peal_speed_value.Text = course_peal_speed.ToString(GlobalVariables.time_string_format)
        Statistics.last_course_time_value.Text = course_speed.ToString(GlobalVariables.time_string_format)
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
