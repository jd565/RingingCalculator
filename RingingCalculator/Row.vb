Public Class Row
    Public bells as new List(of ChangeTime)

    ' Function to return the number of bells in this row
    public readonly property size as integer
	get
	    return bells.count
	end get
    end Property

    ' Function to return the time for this row
    ' taken as the time in ms of the first bell
    ' in this row.
    Public ReadOnly Property time As Integer
        Get
            Return bells(0).time
        End Get
    End Property

    Public ReadOnly Property row_is_full As Boolean
        Get
            Return (Me.size = GlobalVariables.bells.Count)
        End Get
    End Property

    ' Function to add a bell to this row
    Public Sub add(change_time As changetime)
        bells.Add(change_time)
    End Sub

    ' Function to sort this row
    Public Sub sort()
        Me.bells.Sort()
    End Sub

    ' Function to locate a bell in a row
    Public Function locate(bell As Integer)
        For ii As Integer = 0 To Me.size - 1
            If Me.bells(ii).bell = bell Then
                Return ii + 1
            End If
        Next
        Return 0
    End Function

    ' Function to print this row
    Public Function print()
        Dim str As String = ""
        For Each change_time In Me.bells
            str += bell_number_to_string(change_time.bell)
        Next
        Return str
    End Function

    ' Function to return the string representation of the bell number
    Private Function bell_number_to_string(bell_number As Integer) As String
        If bell_number < 10 Then Return bell_number.ToString()
        If bell_number = 10 Then Return "0"
        If bell_number = 11 Then Return "E"
        If bell_number = 12 Then Return "T"
        Return Convert.ToChar(Convert.ToInt32(CChar("A")) - 13 + bell_number).ToString()
    End Function
End class