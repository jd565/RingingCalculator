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
    Public Sub save_config()
        Dim file = My.Computer.FileSystem.OpenTextFileWriter(
            "ringingcalculator.conf", False)
        file.WriteLine("{")
        file.WriteLine("  ""COM_ports"":")
        file.WriteLine("  [")
        file.WriteLine("    {")
        For Each port In GlobalVariables.COM_ports
            file.WriteLine("      ""PortName"":""{0}""", port.PortName)
            file.WriteLine("    }")
            If Not port.Equals(GlobalVariables.COM_ports.Last) Then
                file.Write(",")
            End If
        Next
        file.WriteLine("  ],")
        file.WriteLine("  ""Bells:"":")
        file.WriteLine("  [")
        file.WriteLine("    {")
        For Each bell In GlobalVariables.bells
            file.WriteLine("      ""name"":""{0}"",",
                           bell.name)
            file.WriteLine("      ""bell_number"":{0},",
                           bell.bell_number.ToString)
            file.WriteLine("      ""port_pin"":")
            file.WriteLine("      {")
            file.WriteLine("        ""port"":""{0}"",", bell.port_pin.port)
            file.WriteLine("        ""pin"":{0}", bell.port_pin.pin.ToString)
            file.WriteLine("      },")
            file.WriteLine("      ""handstroke_delay"":{0},",
                           bell.handstroke_delay.ToString)
            file.WriteLine("      ""backstroke_delay"":{0}",
                           bell.backstroke_delay.ToString)
            file.WriteLine("    }")
            If Not bell.Equals(GlobalVariables.bells.Last) Then
                file.Write(",")
            End If
        Next
        file.WriteLine("  ],")
        file.WriteLine("  ""Switch:"":")
        file.WriteLine("  {")
        file.WriteLine("    ""name"":""{0}"",",
                       GlobalVariables.switch.name)
        file.WriteLine("      ""port_pin"":")
        file.WriteLine("      {")
        file.WriteLine("        ""port"":""{0}"",", GlobalVariables.switch.port_pin.port)
        file.WriteLine("        ""pin"":{0}", GlobalVariables.switch.port_pin.pin.ToString)
        file.WriteLine("      },")
        file.WriteLine("  },")
        file.WriteLine("  ""debounce_time"":{0},",
                       GlobalVariables.debounce_time.ToString)
        file.WriteLine("  ""changes_per_lead"":{0},",
                       GlobalVariables.changes_per_lead.ToString)
        file.WriteLine("  ""leads_per_course"":{0}",
                       GlobalVariables.leads_per_course.ToString)
        file.WriteLine("}")

        file.Close()
    End Sub

End Module
