Public Class Statistics
    Public Shared rows As New List(Of Row)

    Public Shared changes As Integer = 0

    Public Shared ReadOnly Property courses As Integer
        Get
            Return Statistics.changes \ GlobalVariables.changes_per_course
        End Get
    End Property

    Public Shared ReadOnly Property leads As Integer
        Get
            Return Statistics.changes \ GlobalVariables.changes_per_lead
        End Get
    End Property
    Public Shared time As TimeSpan
    Public Shared peal_speed As TimeSpan
    Public Shared changes_per_minute As Double

    Public Shared key_vals As New Dictionary(Of String, KeyValueLabel)

    Public Shared Sub init()
        Statistics.key_vals("Changes") = (New KeyValueLabel(True, "Changes"))
        Statistics.key_vals("Leads") = (New KeyValueLabel(True, "Leads"))
        Statistics.key_vals("Time") = (New KeyValueLabel(True, "Time"))
        Statistics.key_vals("Peal Speed") = (New KeyValueLabel(True, "Peal Speed"))
        Statistics.key_vals("Lead End Row") = (New KeyValueLabel(True, "Lead End Row"))
        Statistics.key_vals("Last Course Time") = (New KeyValueLabel(True, "Last Course Time"))
        Statistics.key_vals("Last Course Peal Speed") = (New KeyValueLabel(True, "Last Course Peal Speed"))
        Statistics.key_vals("Changes Per Minute") = (New KeyValueLabel(True, "Changes Per Minute"))
        Statistics.key_vals("Last Minute Changes") = (New KeyValueLabel(True, "Last Minute Changes"))
        Statistics.key_vals("Courses") = (New KeyValueLabel(True, "Courses"))
        GlobalVariables.statistics_init = True
    End Sub

    Public Shared Sub reset_stats()
        Statistics.rows.Clear()
        Statistics.time = TimeSpan.Zero
        Statistics.peal_speed = TimeSpan.Zero
        Statistics.changes = 0
    End Sub

    Public Shared Sub reset_stats_fields()
        For Each kvp As KeyValuePair(Of String, KeyValueLabel) In Statistics.key_vals
            kvp.Value.new_fields()
        Next
    End Sub

End Class
