Module RcDebug

#If DEBUG Then
    Private func_list As New List(Of String)
    Public debug_log As String = "debug_output.log"

    Public Sub debug_entry(func_name As String)
        func_list.Add(func_name)
        debug_print("Entry")
    End Sub

    Public Sub trace_input(bell_number As Integer, debounce As Boolean)
        Dim print_string As String
        Dim d_str As String = ""

        If debounce = True Then
            d_str = "Ignored"
        End If
        print_string = "Bell: " & Str(bell_number).PadRight(2, " ") & d_str & vbCrLf
        My.Computer.FileSystem.WriteAllText("input_tracing.log", print_string, True)
    End Sub

    Public Sub debug_print(message As String)
        Dim print_string As String

        If func_list.Count = 0 Then
            print_string = "NULL".PadRight(40, " ") & message & vbCrLf
        Else
            print_string = func_list.LastOrDefault().PadRight(40, " ") & message & vbCrLf
        End If
        My.Computer.FileSystem.WriteAllText(debug_log, print_string, True)
    End Sub

    Public Sub debug_exit()
        debug_print("Exit")
        func_list.RemoveAt(func_list.LastIndexOf(func_list.Last))
    End Sub

    Public Sub debug_init()
        My.Computer.FileSystem.WriteAllText(debug_log, "", False)
        My.Computer.FileSystem.WriteAllText("input_tracing.log", "", False)
    End Sub
#Else
    Public Sub debug_init()
    End Sub
    Public Sub debug_exit()
    End Sub
    Public Sub debug_print(message as String)
    End Sub
    Public Sub debug_entry(entry as String)
    End Sub
    Public Sub trace_input(bell as Integer, d as Boolean)
    End Sub
#End If

End Module
