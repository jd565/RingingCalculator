Module Saving

    'Function to save a method you have just rung
    ' Takes an input of how often you want to print a row
    Public Sub save_statistics(Optional frequency As Integer = 1)
        Dim current_index As Integer = 0
        Dim file = My.Computer.FileSystem.OpenTextFileWriter(
            "ringing_stats.txt", False)
        If frequency < 1 Then
            Console.WriteLine("frequency is negative")
            GoTo EXIT_LABEL
        End If
        file.WriteLine("{0} changes rung in {1}.",
                       Statistics.changes.ToString,
                       time_to_string(Statistics.time))
        file.WriteLine("Printing every {0} rows.", frequency)

        ' Start printing at the first frequency.
        ' We require the -1 to move it to a 0 based index.
        current_index = frequency - 1
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
    Public Sub load_config(Optional filename As String = "ringingcalculator.conf")
        Dim file = My.Computer.FileSystem.OpenTextFileReader(
            filename)
        Dim reader As New Newtonsoft.Json.JsonTextReader(file)
        Dim serializer As New Newtonsoft.Json.JsonSerializer

        Dim config As Config
        config = serializer.Deserialize(Of Config)(reader)
        config.initialize()

    End Sub

End Module
