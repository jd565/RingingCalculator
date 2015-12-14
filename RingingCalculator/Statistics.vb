Public Class Statistics
    Public Shared rows As New List(Of Row)

    Public Shared ReadOnly Property changes As Integer
        Get
            Return Statistics.rows.Count
        End Get
    End Property
    Public Shared ReadOnly Property leads As Integer
        Get
            Return Statistics.changes Mod GlobalVariables.changes_per_lead
        End Get
    End Property
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
    Public Shared last_course_time_key As Label
    Public Shared last_course_time_value As Label
    Public Shared last_course_peal_speed_key As Label
    Public Shared last_course_peal_speed_value As Label

    Public Shared Sub reset_stats()
        Statistics.rows.Clear()
        Statistics.time = 0
        Statistics.peal_speed = 0
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
    End Sub

End Class
