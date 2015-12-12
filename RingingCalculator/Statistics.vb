Public Class Statistics
    Public Shared rows as new List(of Row)
    Public Shared estimate_method_data As Boolean = True
    Public Shared changes_per_lead As Integer
    Public Shared leads_per_course As Integer

    Public Shared changes As Integer=0
    Public Shared leads As Integer=0
    Public Shared time As Integer
    Public Shared peal_speed As Integer

    Public Shared changes_key As Label
    Public Shared changes_value As Label
    Public Shared leads_key As Label
    Public Shared leads_value As Label
    Public Shared time_key As Label
    Public Shared time_value As Label
    Public Shared peal_speed_key As Label
    Public Shared peal_speed_value As Label
    Public Shared lead_end_row_key As Label
    Public Shared lead_end_row_value As Label

    Public Shared Sub reset_stats()
        Statistics.changes_key = New Label
        Statistics.changes_value = New Label
        Statistics.leads_key = New Label
        Statistics.leads_value = New Label
        Statistics.time_key = New Label
        Statistics.time_value = New Label
        Statistics.peal_speed_key = New Label
        Statistics.peal_speed_value = New Label
        Statistics.lead_end_row_key = New Label
        Statistics.lead_end_row_value = New Label
    End Sub


    Public Shared ReadOnly Property number_of_leads As Integer
        Get
            Return Statistics.lead_end_times.Count
        End Get
    End Property
End Class
