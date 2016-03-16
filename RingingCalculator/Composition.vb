Public Class RingingCall
    ' This class holds information about each call
    Public location As String
    Public call_to_make As String
End Class


Public Class Composition
    ' This class holds composition data and parses it into a usable format
    ' The expected format for the composition is:
    ' M W B H
    ' - 2
    '   s   3
    ' ...

    ' The full composition string
    Public full_composition As String

    ' The parsed composition string
    Public composition As List(Of RingingCall)

    Public Sub New(full_comp As String)
        Me.full_composition = full_comp
        Me.composition = New List(Of RingingCall)
        Me.parse()
    End Sub

    ' Function to parse the full notation into a string
    Public Sub parse()
        Dim ii As Integer = 0
        Dim call_indicies As New List(Of Integer)
        Dim calls As New List(Of String)
        Dim temp As String = ""
        Dim sep_list As Char() = {" ", vbCr, vbLf}
        Dim line_index As Integer = 0
        Dim call_index As Integer
        Dim c As RingingCall

        ' Parse the top line to populate the calls and call_indicies
        While Me.full_composition(ii) <> vbLf
            If (sep_list.Contains(Me.full_composition(ii))) Then
                If temp <> "" Then
                    calls.Add(temp)
                    call_indicies.Add(ii - temp.Length)
                End If
                temp = ""
            Else
                temp += Me.full_composition(ii)
            End If
            ii += 1
        End While



        ' Now look through the next lines for the rest of the composition.
        ' if we find a -, number or s that lines up with a call then add it to our list
        While (ii + line_index < Me.full_composition.Length)
            temp = ""
            c = New RingingCall
            If Me.full_composition(ii + line_index) = vbLf Then
                ' We want to look through line by line so set the line index
                ' to the start of the line you're looking through and the index as the
                ' distance into this line.
                line_index = ii + line_index + 1
                ii = 0
            End If
            If (ii + line_index < Me.full_composition.Length AndAlso
                call_indicies.Contains(ii) And
                Not sep_list.Contains(Me.full_composition(ii + line_index))) Then
                ' There is a call at this location. Add the call to the list
                While (ii + line_index < Me.full_composition.Length AndAlso
                       Not sep_list.Contains(Me.full_composition(ii + line_index)))
                    temp += Me.full_composition(ii + line_index)
                    ii += 1
                End While
                call_index = call_indicies.IndexOf(ii - temp.Length)
                c.location = calls(call_index)

                ' Now handle the composition having numbers of multiple calls e.g. 3 or 2s
                ' The temp string contains the whole call, so look through this for a number
                ' We do not handle 2-digit numbers as this will never happen.
                ' We do handle writing 2s or s2 however. A number with no call is assumed bob.
                If temp.Length = 1 Then
                    If Not (temp(0) >= "0" And temp(0) <= "9") Then
                        ' We just have a standard call. Add this.
                        c.call_to_make = temp
                        Me.composition.Add(c)
                    Else
                        ' Length 1 but a number. This is multiple bobs at this location
                        c.call_to_make = "-"
                        For jj As Integer = 1 To Val(temp(0))
                            Me.composition.Add(c)
                        Next
                    End If
                Else
                    ' This is not a standard call. Check for the first char being a number.
                    If (temp(0) >= "0" And temp(0) <= "9") Then
                        ' The number has been set. The rest of temp contains the call to be made.
                        c.call_to_make = temp.Substring(1)
                        For jj As Integer = 1 To Val(temp(0))
                            Me.composition.Add(c)
                        Next
                    Else
                        ' The number is the last one in the string
                        If Not (temp.Last >= "0" And temp.Last <= "9") Then
                            ' If the last char is not a number then throw an execption and exit.
                            Throw New System.Exception("Bad composition")
                            Exit Sub
                        End If
                        c.call_to_make = temp.Substring(0, temp.Length - 1)
                        For jj As Integer = 1 To Val(temp.Last)
                            Me.composition.Add(c)
                        Next
                    End If
                End If
            Else
                ii += 1
            End If
        End While

    End Sub

End Class
