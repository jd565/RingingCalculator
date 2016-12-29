Module Saving

    Private time_format As String = GlobalVariables.full_time

    'Function to save a method you have just rung
    ' Takes an input of how often you want to print a row
    Public Sub save_statistics(Optional name As String = "ringingcalculator.txt",
                               Optional frequency As Integer = 1,
                               Optional method As Method = Nothing)
        Dim file = My.Computer.FileSystem.OpenTextFileWriter(
            name, False)
        Dim out_string As String

        out_string = statistics_string(frequency, method)

        If out_string IsNot Nothing Then
            file.WriteLine(out_string)
        End If

EXIT_LABEL:
        file.Close()
    End Sub

    Public Function statistics_string(Optional frequency As Integer = 1,
                                      Optional method As Method = Nothing) As String
        Dim current_index As Integer = 0
        Dim ii As Integer = 1
        Dim changes As Integer
        Dim leads As Integer
        Dim courses As Integer
        Dim time As TimeSpan
        Dim cpm As Double
        Dim cpl As Integer
        Dim start_idx As Integer
        Dim rows As List(Of Row)
        Dim start_time As Date
        Dim out_string As String = ""
        Dim false_row_ids As New List(Of Integer)
        Dim peal_speed As TimeSpan
        Dim lpc As Integer

        If frequency < 1 Then
            RcDebug.debug_print("Frequency is negative")
            Return Nothing
        End If

        If method IsNot Nothing Then
            changes = method.rows.Count
            leads = changes \ method.changes_per_lead
            cpl = method.changes_per_lead
            start_idx = 0
            rows = method.rows
            start_time = New DateTime(0)
        Else
            cpm = Statistics.changes_per_minute
            cpl = GlobalVariables.changes_per_lead
            lpc = GlobalVariables.leads_per_course
            start_idx = GlobalVariables.start_index
            changes = Statistics.changes
            rows = Statistics.rows.GetRange(start_idx, changes)
            leads = Statistics.leads
            start_time = GlobalVariables.start_time
            time = Statistics.time
            cpm = Statistics.changes_per_minute
            courses = Statistics.courses
            peal_speed = Statistics.peal_speed
        End If

        out_string += ("Changes: " & changes & vbCrLf)
        out_string += ("Leads: " & leads & vbCrLf)
        out_string += ("Changes per lead: " & cpl & vbCrLf)
        If method Is Nothing Then
            out_string += ("Courses: " & courses & vbCrLf)
            out_string += ("Time: " & time.ToString(time_format) & vbCrLf)
            out_string += ("Changes per minute: " & cpm.ToString(GlobalVariables.cpm_string_format) & vbCrLf)
            out_string += ("Changes per course: " & GlobalVariables.changes_per_course & vbCrLf)
            out_string += ("Leads per course: " & lpc & vbCrLf)
            out_string += ("Peal speed: " & peal_speed.tostring(GlobalVariables.hours_and_mins) & vbCrLf)
        End If
        out_string += vbCrLf

        ' Print out the numbers of the lead ends and the time taken for them.
        ' If this is a method then do not print times.
        If method Is Nothing Then
            out_string += generate_lead_ends_output(cpl, rows, start_time, lpc, True)
        Else
            out_string += generate_lead_ends_output(0, rows, start_time, 0, False)
        End If

        If method Is Nothing Then
#If 0 Then
            ' Add in the bell delay stats.
            out_string += "Bell delay statistics:".PadRight(22) & vbCrLf
            out_string += " ".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += ("Bell " & (ii + 1).ToString).PadLeft(9)
            Next
            out_string += vbCrLf
            out_string += "Average delay (ms):".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_stats(ii).average.ToString("####0").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += "Average deviation:".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_stats(ii).std_dev.ToString("##0.##").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += "Handstroke delay (ms):".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_stats(ii).h_average.ToString("####0").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += "Backstroke delay (ms):".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_stats(ii).b_average.ToString("####0").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += "Handstroke lead (ms):".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_lead_stats(ii).h_average.ToString("####0").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += "Backstroke lead (ms):".PadRight(22)
            For ii = 0 To rows(0).size - 1
                out_string += Statistics.bell_lead_stats(ii).b_average.ToString("####0").PadLeft(9, "-")
            Next
            out_string += vbCrLf
            out_string += vbCrLf
#End If
        End If

        If RingingCalculator.Row.list_is_true(rows, false_row_ids) Then
            out_string += ("This performance is true" & vbCrLf)
        Else
            out_string += ("This performance is not true" & vbCrLf)
            out_string += "False rows: " & vbCrLf
            For Each index In false_row_ids
                out_string += ((index + 1).ToString.PadRight(6, " ") &
                           rows(index).print & vbCrLf)
            Next
        End If

        out_string += vbCrLf

        out_string += ("Printing every " & frequency & " rows." & vbCrLf)

        ' Start printing at the first frequency.
        ' We require the -1 to move it to a 0 based index.
        current_index = frequency - 1
        While current_index < rows.Count
            out_string += ((current_index + 1).ToString.PadRight(6, " ") &
                           rows(current_index).print)
            Try
                If false_row_ids.Contains(current_index) Then
                    out_string += " <- False row"
                End If
            Catch ex As Exception
            End Try
            out_string += vbCrLf
            If (current_index + 1) Mod cpl = 0 And current_index > 0 Then
                out_string += ("-".PadLeft(6 + rows(0).size, "-") & vbCrLf)
            End If
            current_index += frequency
        End While

        Return out_string

    End Function

    ' Function to save all your configuration to a file that can be read
    ' This is essentially a list of all the globalvariables
    Public Sub save_config(Optional filename As String = "ringingcalculator.conf")
        Dim file = My.Computer.FileSystem.OpenTextFileWriter(
            filename, False)
        Dim writer As New Newtonsoft.Json.JsonTextWriter(file)
        writer.Formatting = Newtonsoft.Json.Formatting.Indented
        Dim serializer As New Newtonsoft.Json.JsonSerializer

        Dim config As New Config(GlobalVariables.COM_ports,
                                 GlobalVariables.bells,
                                 GlobalVariables.switch,
                                 GlobalVariables.debounce_time,
                                 GlobalVariables.changes_per_lead,
                                 GlobalVariables.leads_per_course)

        serializer.Serialize(writer, config)

        file.Close()
    End Sub

    ' Function to read the config file.
    ' This is saved in a JSON format, so read it in here
    ' Returns a true/false value depending on whether it succeeded
    Public Function load_config(Optional filename As String = "ringingcalculator.conf") As Boolean
        Dim file = My.Computer.FileSystem.OpenTextFileReader(
            filename)
        Dim reader As New Newtonsoft.Json.JsonTextReader(file)
        Dim serializer As New Newtonsoft.Json.JsonSerializer
        Dim ret As Boolean

        Dim config As Config
        config = serializer.Deserialize(Of Config)(reader)
        ret = config.initialize()
        If ret = False Then
            MsgBox("Failed to open " & filename,, "Error")
        End If
        file.Close()
        Return ret
    End Function

    ' Function to generate string output for lead ands of stats form
    Private Function generate_lead_ends_output(cpl As Integer,
                                               rows As List(Of Row),
                                               start_time As Date,
                                               lpc As Integer,
                                               print_time As Boolean)
        Dim out_string As String
        Dim current_index As Integer
        Dim row As String
        Dim total_time As TimeSpan
        Dim lead_time As Double
        Dim course_time As TimeSpan
        Dim ii As Integer = 0
        out_string = ("Lead end".PadRight(12) &
                      "Row".PadRight(18))
        If print_time Then
            out_string += ("Time at lead end".PadRight(20) &
                           "Time of lead (s)".PadRight(20) &
                           "Time of course".PadRight(12))
        End If
        out_string += vbCrLf
        current_index = cpl - 1
        ii = 1
        While current_index < rows.Count
            row = rows(current_index).print
            out_string += (ii.ToString.PadRight(12, "-") &
                           row.PadRight(18, "-"))
            If print_time Then
                total_time = rows(current_index).time.Subtract(start_time)
                If ii = 1 Then
                    lead_time = total_time.TotalSeconds
                Else
                    lead_time = rows(current_index).time.Subtract(rows(current_index - cpl).time).TotalSeconds
                End If
                out_string += (total_time.ToString(time_format).PadRight(20, "-") &
                           lead_time.ToString("##0.00").PadRight(20, "-"))
                If ii Mod lpc = 0 Then
                    ' Add in the time for this course
                    If ii / lpc = 1 Then
                        course_time = total_time
                    Else
                        course_time = rows(current_index).time.Subtract(rows(current_index - cpl * lpc).time)
                    End If
                    out_string += course_time.ToString("mm\:ss").PadRight(12, "-")
                End If
            End If
            out_string += vbCrLf
            ii += 1
            current_index += cpl
        End While

        out_string += vbCrLf
        Return out_string
    End Function

End Module
