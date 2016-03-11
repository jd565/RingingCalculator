Imports RingingCalculator

Public Class Row
    Implements IComparable(Of Row)
    Public bells As New List(Of ChangeTime)

    ' Function to return the number of bells in this row
    Public ReadOnly Property size As Integer
        Get
            Return bells.Count
        End Get
    End Property

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
    Public Sub add(change_time As ChangeTime)
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

    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
            Return False
        End If

        Dim other_row As Row = CType(obj, Row)
        If Me.size <> other_row.size Then
            Return False
        End If
        For ii = 0 To Me.size - 1
            If Me.bells(ii).bell <> other_row.bells(ii).bell Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Shared Function list_is_true(rows As List(Of Row)) As Boolean
        For ii = 0 To rows.Count - 2
            For jj = ii + 1 To rows.Count - 1
                If rows(ii).Equals(rows(jj)) Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Public Function CompareTo(other As Row) As Integer Implements IComparable(Of Row).CompareTo
        Throw New NotImplementedException()
    End Function
End Class