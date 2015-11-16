Public Class Statistics
    Public Shared lead_end_times As New List(Of ChangeTime)
    Public Shared estimate_method_data As Boolean = True
    Public Shared changes_per_lead As Integer
    Public Shared lwb As Bell
    Public Shared leads_per_course As Integer

    Public Shared ReadOnly Property number_of_leads As Integer
        Get
            Return Statistics.lead_end_times.Count
        End Get
    End Property
End Class
