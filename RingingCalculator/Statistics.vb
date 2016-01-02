Public Class Statistics
    Public Shared rows As New List(Of Row)

    Public Shared changes As Integer = 0

    Public Shared ReadOnly Property leads As Integer
        Get
            Return Statistics.changes \ GlobalVariables.changes_per_lead
        End Get
    End Property
    Public Shared time As TimeSpan
    Public Shared peal_speed As TimeSpan

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
    Public Shared last_course_time_key As Label
    Public Shared last_course_time_value As Label
    Public Shared last_course_peal_speed_key As Label
    Public Shared last_course_peal_speed_value As Label
    Public Shared changes_per_minute_key As Label
    Public Shared changes_per_minute_value As Label
    Public Shared last_minute_changes_key As Label
    Public Shared last_minute_changes_value As Label

    Public Shared Sub reset_stats()
        Statistics.rows.Clear()
        Statistics.time = TimeSpan.Zero
        Statistics.peal_speed = TimeSpan.Zero
        Statistics.changes = 0
    End Sub

    Public Shared Sub reset_stats_fields()
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
        Statistics.last_course_time_key = New Label
        Statistics.last_course_time_value = New Label
        Statistics.last_course_peal_speed_key = New Label
        Statistics.last_course_peal_speed_value = New Label
        Statistics.changes_per_minute_key = New Label
        Statistics.changes_per_minute_value = New Label
        Statistics.last_minute_changes_key = New Label
        Statistics.last_minute_changes_value = New Label
    End Sub

End Class
