Public Class Row
    Public bells as new List(of ChangeTime)

    ' Function to return the number of bells in this row
    public readonly property size as integer
	get
	    return bells.count
	end get
    end property

    ' Function to return the time for this row
    ' taken as the time in ms of the first bell
    ' in this row.
    public readonly property time as integer
        get
	    return bells(0).time
        end get
    end property

    ' Function to add a bell to this row
    Public Sub add(change_time as changetime)
	bells.add(change_time)
    end sub
end class