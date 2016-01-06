Module Saving

    'Function to save a method you have just rung
    ' Takes an input of how often you want to print a row
    Public Sub save_statistics(Optional name As String = "ringingcalculator.txt", Optional frequency As Integer = 1)
        Dim current_index As Integer = 0
        Dim time_format As String = GlobalVariables.full_time
        Dim ii As Integer = 1
        Dim row As String
        Dim total_time As TimeSpan
        Dim lead_time As TimeSpan
        Dim file = My.Computer.FileSystem.OpenTextFileWriter(
            name, False)
        If frequency < 1 Then
            Console.WriteLine("Frequency is negative")
            GoTo EXIT_LABEL
        End If
        file.WriteLine("{0} changes rung in {1}.",
                       Statistics.changes.ToString,
                       Statistics.time.ToString(time_format))
        file.WriteLine("Printing every {0} rows.", frequency)
        file.WriteLine()

        ' Print out the numbers of the lead ends and the time taken for them
        file.WriteLine("Lead end".PadRight(20) &
                       "Row".PadRight(20) &
                       "Time at lead end".PadRight(20) &
                       "Time of lead".PadRight(20))
        current_index = GlobalVariables.changes_per_lead + GlobalVariables.start_row - 1
        row = Statistics.rows(current_index).print
        total_time = Statistics.rows(current_index).time.Subtract(GlobalVariables.start_time)
        lead_time = Statistics.rows(current_index).time.Subtract(GlobalVariables.start_time)
        file.WriteLine(ii.ToString.PadRight(20, "-") &
                       row.PadRight(20, "-") &
                       total_time.ToString(time_format).PadRight(20, "-") &
                       lead_time.ToString(time_format).PadRight(20, "-"))
        current_index += GlobalVariables.changes_per_lead
        ii += 1
        While current_index < Statistics.rows.Count
            row = Statistics.rows(current_index).print
            total_time = Statistics.rows(current_index).time.Subtract(GlobalVariables.start_time)
            lead_time = Statistics.rows(current_index).time.Subtract(
                               Statistics.rows(current_index - GlobalVariables.changes_per_lead).time)
            file.WriteLine(ii.ToString.PadRight(20, "-") &
                           row.PadRight(20, "-") &
                           total_time.ToString(time_format).PadRight(20, "-") &
                           lead_time.ToString(time_format).PadRight(20, "-"))
            ii += 1
            current_index += GlobalVariables.changes_per_lead
        End While

        file.WriteLine()

        ' Start printing at the first frequency.
        ' We require the -1 to move it to a 0 based index.
        current_index = frequency + GlobalVariables.start_row - 1
        While current_index < Statistics.rows.Count
            file.WriteLine(Statistics.rows(current_index).print)
            current_index += frequency
        End While

EXIT_LABEL:
        file.Close()
    End Sub

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

End Module
