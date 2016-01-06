Module ReportingFunctions

    ' Function to be called once all bells have rung a certain row.
    ' This happens for every row, and the sorted row is put into the rows
    ' list.
    Public Sub row_is_full(change_id As Integer)
        Console.WriteLine("Row is full")
        Statistics.rows(change_id).sort()

        ' Check if this row is the start of the method
        If Not GlobalVariables.method_started Then
            If has_method_started(change_id) Then
                ' We want to reset the number of changes at the start of the method
                Statistics.changes = 0
            End If
        End If

        update_changes(change_id)
        ' If the switch has been pressed to stop running then stop recording here
        If Not GlobalVariables.switch.is_running Then
            GlobalVariables.method_started = False
            maybe_update_lead_statistics(change_id, True)
        End If
        If GlobalVariables.method_started Then
            maybe_update_lead_statistics(change_id)
        End If

    End Sub

    ' Function to find how many changes have been rung in the last minute
    Private Function get_changes_last_minute(change_id As Integer) As Integer
        Dim curr_time As TimeSpan
        Dim check_time As TimeSpan
        Dim change_to_check As Integer = 30

        curr_time = Statistics.rows(change_id).time.Subtract(Statistics.rows(0).time)
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

    Private Sub update_changes(change_id As Integer, Optional force As Boolean = False)
        Dim time_diff As TimeSpan
        Dim cpm_val As Double
        Dim clm_diff As TimeSpan
        Dim clm_val As Double

        ' If we haven't started the method then the start time may not be properly set.
        ' Use the time of the first change for this instead.
        If GlobalVariables.method_started Then
            time_diff = Statistics.rows(change_id).time.Subtract(GlobalVariables.start_time)
            cpm_val = (change_id + 1) / (time_diff.TotalMinutes)
        Else
            time_diff = Statistics.rows(change_id).time.Subtract(Statistics.rows(0).time)
            If change_id <> 0 Then
                cpm_val = (change_id) / (time_diff.TotalMinutes)
            Else
                cpm_val = 0
            End If
        End If

        If change_id > 10 Then
            clm_diff = Statistics.rows(change_id).time.Subtract(Statistics.rows(change_id - 10).time)
            clm_val = 10 / time_diff.TotalMinutes
        Else
            clm_val = 0
        End If

        Statistics.changes += 1
        Statistics.changes_value.Text = Statistics.changes.ToString
        Statistics.time = time_diff
        Statistics.time_value.Text = Statistics.time.ToString(GlobalVariables.hours_and_mins)
        Statistics.changes_per_minute_value.Text = cpm_val.ToString(GlobalVariables.cpm_string_format)
        Statistics.last_minute_changes_value.Text = clm_val.ToString(GlobalVariables.cpm_string_format)
    End Sub

    ' Function to check whether this is a lead end and update the statistics.
    ' The force parameter provides a way to update stats even if it is not a lead end.
    Private Sub maybe_update_lead_statistics(change_id As Integer, Optional force As Boolean = False)
        Dim row As Row = Statistics.rows(change_id)
        Dim changes_per_peal As Integer = GlobalVariables.changes_per_peal
        Dim time_diff As TimeSpan
        Dim start_time As DateTime = GlobalVariables.start_time

        If (Statistics.changes Mod GlobalVariables.changes_per_lead <> 0) And Not force Then Exit Sub

        If start_time.Ticks = 0 Then
            start_time = Statistics.rows(0).time
        End If

        time_diff = Statistics.rows(change_id).time.Subtract(start_time)
        Console.WriteLine("Update lead end stats")

        Statistics.peal_speed = New TimeSpan(time_diff.Ticks * changes_per_peal / Statistics.changes)
        Statistics.lead_end_row_value.Text = Statistics.rows(change_id).print
        Statistics.leads_value.Text = Statistics.leads.ToString
        Statistics.peal_speed_value.Text = Statistics.peal_speed.ToString(GlobalVariables.hours_and_mins)
        maybe_update_course_statistics(change_id)
    End Sub

    ' Function to update the course en statistics if this is a course end
    Private Sub maybe_update_course_statistics(change_id As Integer)
        ' If this is not a course end then drop out here
        If (Statistics.changes) Mod GlobalVariables.changes_per_course <> 0 Then Exit Sub

        Dim course_speed As TimeSpan
        Dim start_time_of_course As DateTime
        Dim course_peal_speed As TimeSpan

        If change_id < GlobalVariables.changes_per_course Then
            start_time_of_course = Statistics.rows(0).time
        Else
            start_time_of_course = Statistics.rows(change_id - GlobalVariables.changes_per_course).time
        End If
        course_speed = Statistics.rows(change_id).time.Subtract(start_time_of_course)
        course_peal_speed = New TimeSpan(course_speed.Ticks * GlobalVariables.changes_per_peal / GlobalVariables.changes_per_course)
        Statistics.last_course_peal_speed_value.Text = course_peal_speed.ToString(GlobalVariables.hours_and_mins)
        Statistics.last_course_time_value.Text = course_speed.ToString(GlobalVariables.hours_and_mins)
        Statistics.courses_value.Text = Statistics.courses.ToString
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
