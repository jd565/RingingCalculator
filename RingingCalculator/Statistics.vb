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
    Public Shared place_stats As New List(Of PlaceStats)
    Public Shared bell_stats As New List(Of PlaceStats)
    Public Shared bell_lead_stats As New List(Of PlaceStats)

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
        Statistics.place_stats.Clear()
        Statistics.bell_stats.Clear()
        Statistics.bell_lead_stats.Clear()
    End Sub

    Public Shared Sub reset_stats_fields()
        For Each kvp As KeyValuePair(Of String, KeyValueLabel) In Statistics.key_vals
            kvp.Value.new_fields()
        Next
    End Sub

    ' Function to create a list of stats big enough for the data we have
    Public Shared Sub create_place_stats()
        Statistics.place_stats.Clear()
        Statistics.bell_lead_stats.Clear()
        Statistics.bell_stats.Clear()
        For ii = 1 To Statistics.rows(0).size
            Statistics.place_stats.Add(New PlaceStats)
            Statistics.bell_stats.Add(New PlaceStats)
            Statistics.bell_lead_stats.Add(New PlaceStats)
        Next
    End Sub

    ' Function to fill out the delay for every row.
    Public Shared Sub generate_place_delays()
        Dim num_bells As Integer = Statistics.rows(0).size
        Dim handstroke As Boolean
        Dim delay As Integer

        Statistics.create_place_stats()

        For jj = GlobalVariables.start_index To Statistics.rows.Count - 1
            handstroke = False
            If Statistics.rows(jj).bells(0).change Mod 2 = 1 Then
                handstroke = True
            End If
            For ii = 1 To num_bells
                delay = calc_row_place_delay(ii, jj)
                Statistics.place_stats(ii - 1).add(delay, handstroke)

                If ii > 1 Then
                    ' Add statistics to the bell stats if this isn't the lead.
                    ' If it is the lead then add them in a different stats group
                    Statistics.bell_stats(Statistics.rows(jj).bells(ii - 1).bell - 1).add(delay, handstroke)
                Else
                    Statistics.bell_lead_stats(Statistics.rows(jj).bells(ii - 1).bell - 1).add(delay, handstroke)
                End If
            Next
        Next
    End Sub

    ' Function to calculate the delay for a certain place in a certain row.
    ' This is called a lot by the relevant functions.
    Private Shared Function calc_row_place_delay(place As Integer, row_idx As Integer) As Integer
        Dim delay As Integer
        Dim row As Row = Statistics.rows(row_idx)
        Dim num_bells As Integer = Statistics.rows(0).size
        If place = 1 Then
            delay = row.bells(0).time.Subtract(Statistics.rows(row_idx - 1).bells(num_bells - 1).time).TotalMilliseconds
        Else
            delay = row.bells(place - 1).time.Subtract(row.bells(place - 2).time).TotalMilliseconds
        End If

        Return delay
    End Function

End Class

Public Class PlaceStats
    ' This class calculates stats for a single bell or place.
    ' This holds the average handstroke, backstroke and total delay in a list.
    ' The supplied properties provide means of adding, and retrieving stats from the class.
    Private handstroke_delay As List(Of UInteger)
    Private backstroke_delay As List(Of UInteger)
    Private total_delay As List(Of UInteger)

    Public Sub New()
        Me.handstroke_delay = New List(Of UInteger)
        Me.backstroke_delay = New List(Of UInteger)
        Me.total_delay = New List(Of UInteger)
    End Sub

    Public Sub add(delay As UInteger, handstroke As Boolean)
        Me.total_delay.Add(delay)
        If handstroke Then
            Me.handstroke_delay.Add(delay)
        Else
            Me.backstroke_delay.Add(delay)
        End If
    End Sub

    Public ReadOnly Property h_average As Double
        Get
            Return Me.handstroke_delay.Average(Function(input) input)
        End Get
    End Property
    Public ReadOnly Property b_average As Double
        Get
            Return Me.backstroke_delay.Average(Function(input) input)
        End Get
    End Property
    Public ReadOnly Property average As Double
        Get
            Return Me.total_delay.Average(Function(input) input)
        End Get
    End Property

    Public ReadOnly Property h_std_dev As Double
        Get
            Dim mean As Double = Me.h_average
            Dim variance As Double
            variance = Me.handstroke_delay.Average(Function(input) ((input - mean) ^ 2))
            Return Math.Sqrt(variance)
        End Get
    End Property
    Public ReadOnly Property b_std_dev As Double
        Get
            Dim mean As Double = Me.b_average
            Dim variance As Double
            variance = Me.backstroke_delay.Average(Function(input) ((input - mean) ^ 2))
            Return Math.Sqrt(variance)
        End Get
    End Property
    Public ReadOnly Property std_dev As Double
        Get
            Dim mean As Double = Me.average
            Dim variance As Double
            variance = Me.total_delay.Average(Function(input) ((input - mean) ^ 2))
            Return Math.Sqrt(variance)
        End Get
    End Property

End Class
