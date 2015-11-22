Public Class ChangeTime
    Implements IComparable(Of ChangeTime)

    Private change_time As Integer
    Public ReadOnly Property time() As Integer
        Get
            Return change_time
        End Get
    End Property

    Private bell_number As Integer
    Public ReadOnly Property bell() As Integer
        Get
            Return Me.bell_number
        End Get
    End Property

    Private change_number As Integer
    Public ReadOnly Property change As Integer
        Get
            Return Me.change_number
        End Get
    End Property

    Public ReadOnly Property string_time As String
        Get
            Return time_to_string(Me.time)
        End Get
    End Property

    Public Sub New(bell_number As Integer, change_number As Integer, change_time As Integer)
        Me.bell_number = bell_number
        Me.change_number = change_number
        Me.change_time = change_time
    End Sub

    ' Function to allow this class to be compared and sorted.
    Public Function CompareTo(other_changetime As ChangeTime) As Integer _
        Implements IComparable(Of ChangeTime).CompareTo

        ' If the object to compare to doesn't exist, return 1
        If other_changetime Is Nothing Then Return 1

        ' Otherwise just return the value of comparing the 2 times
        Return Me.time.CompareTo(other_changetime.time)
    End Function

End Class
