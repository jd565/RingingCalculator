Module MethodEstimatingFunctions

    ' Function for checking if this is a lead end
    Public Function is_this_lead_end(change_id As Integer) As Boolean
        return (GlobalVariables.changes_per_lead Mod (change_id + 1) = 0)
    End Function

    Private Sub update_lead_end_statistics(change_id As Integer)
        Console.WriteLine("Update lead end stats")
        Statistics.lead_end_times.Add(GlobalVariables.bells(0).change_times(change_id))

        Statistics.changes = change_id + 1
        Statistics.time = GlobalVariables.bells(0).change_times(change_id).time
        Statistics.peal_speed = GlobalVariables.bells(0).change_times(change_id).time *
                                                          GlobalVariables.changes_per_peal / (change_id + 1)
        Statistics.lead_end_row_value.Text = print_change(Statistics.changes - 1)
        Statistics.leads_value.Text = Statistics.number_of_leads.ToString
        Statistics.changes_value.Text = Statistics.changes.ToString
        Statistics.time_value.Text = time_to_string(Statistics.time)
        Statistics.peal_speed_value.Text = time_to_string(Statistics.peal_speed)
    End Sub

    ' Function to call when the row is full,
    ' This is called to print the row.
    Public Sub row_is_full(row as row)
        Console.WriteLine("Row is full")
	statistics.rows.add(row)
	statistics.changes++
	if is_this_lead_end(statistics.changes - 1) then
	    update_lead_end_statistics(statistics.changes - 1)
        end if
    End Sub

End Module
