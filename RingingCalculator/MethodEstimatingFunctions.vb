Module MethodEstimatingFunctions

    ' Function for checking if this is a lead end
    Public Function is_this_lead_end(change_id As Integer) As Boolean
        Return (get_bell_location_in_row(change_id, 1) = 1 And
                get_bell_location_in_row(change_id - 1, 1) = 1)
    End Function

    ' Function to call when the treble has just rung 
    Public Sub treble_has_just_rung(change_id As Integer)
        ' Let us check if it is the lead end
        If is_this_lead_end(change_id) Then
            update_lead_end_statistics(change_id)
        End If
    End Sub

    Private Sub update_lead_end_statistics(change_id As Integer)
        Statistics.lead_end_times.Add(GlobalVariables.bells(0).change_times(change_id - 1))

        ' If this is the first lead end try to work out information about the method
        If Statistics.estimate_method_data Then
            estimate_method_data()
        End If

    End Sub

    ' Function for estimating data about the method that is being rung right now
    ' This is called once, unless the function detects that the result doesn't make sense,
    ' In which case it will try again at the next lead end
    Private Sub estimate_method_data()

        ' Work out the number of changes per lead, which will help us with estimating
        ' the number of leads per course
        If Statistics.number_of_leads = 1 Then
            Statistics.changes_per_lead = Statistics.lead_end_times(Statistics.number_of_leads - 1).change
        Else
            Statistics.changes_per_lead = (Statistics.lead_end_times(Statistics.number_of_leads - 1).change -
                                          Statistics.lead_end_times(Statistics.number_of_leads - 2).change)
        End If

        ' Sanatize the changes per lead. It cannot be odd, and we expect it to be greater than 5
        If Statistics.changes_per_lead Mod 2 = 1 Or Statistics.changes_per_lead < 5 Then
            Console.WriteLine("attempted to estimate method data but got confused")
            Exit Sub
        End If

        ' Figure out which bell is the last working bell (lwb).
        ' Do this by looking at the first 3 rows of the method and seeing if the bells move place.
        For ii As Integer = GlobalVariables.bells.Count To 1 Step -1
            If is_this_bell_lwb(GlobalVariables.bells(ii - 1)) Then
                Console.Write("found lwb as {0}", GlobalVariables.bells(ii - 1).name)
                Statistics.lwb = GlobalVariables.bells(ii - 1)
                Exit For
            End If
        Next

        ' The changes per lead should match a multiple of the lwb
        ' e.g. surprise major has lwb as 8, 32 changes per lead.
        ' obviously this doesn't work for all methods (little bob), but the
        ' majority will pass this
        If Statistics.changes_per_lead Mod Statistics.lwb.bell_number <> 0 Then
            Exit Sub
        End If

        Statistics.leads_per_course = Statistics.lwb.bell_number - 1

        ' If we have got here, then we have been successful and should turn off the checker.
        Statistics.estimate_method_data = False

    End Sub

    ' Function to check if this bell is the last working bell.
    ' check the first 3 changes and see if it moves place.
    Private Function is_this_bell_lwb(bell As Bell) As Boolean
        For ii As Integer = 1 To 3
            If (get_bell_location_in_row(ii, bell.bell_number) <> bell.bell_number) Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
