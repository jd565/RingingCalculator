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
    Public ReadOnly Property time As DateTime
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
    Public Function print() As String
        Dim str As String = ""
        For Each change_time In Me.bells
            str += bell_number_to_string(change_time.bell)
        Next
        Return str
    End Function

    Public Function is_this_rounds() As Boolean
        For ii As Integer = 0 To Me.bells.Count - 1
            If Me.bells(ii).bell <> ii + 1 Then Return False
        Next
        Return True
    End Function
End class