Public Class PlaceNotation
    Public full_notation As String
    Public notation As New List(Of Notation)

    Public Sub New(notation As String)
        Me.full_notation = notation
        Me.parse()
    End Sub

    ' Function to parse the full notation and split it into smaller parts
    Public Sub parse()
        Dim c As String = Me.full_notation
        Dim temp As String
        Dim ii As Integer = 0
        Dim block_done As Boolean
        Dim block_idx As Integer = 0
        Dim rev_block As Integer = -1
        Me.notation.Clear()

        ' Now move through the notation and extract each block
        ' Blocks are separated by commas, or if the next input is smaller.
        While (ii < c.Count)
            block_done = False
            temp = ""
            If (c(ii) = CChar("x") Or c(ii) = CChar("X")) Then
                block_done = True
                temp = "X"
                ii += 1
            End If
            If (c(ii) = CChar(",")) Then
                ii += 1
            End If
            While Not block_done
                If c(ii) = CChar("&") Then
                    ii += 1
                    rev_block = block_idx
                End If

                ' Loop through adding the string to temp until:
                ' - we hit an x, a comma, or a number that is smaller.
                temp += c(ii)
                If ii = c.Count - 1 Then
                    ' We have reached the end of the input so exit
                    ' Increment ii to escape the first loop
                    ii += 1
                    Exit While
                End If

                If (c(ii + 1) = CChar("x") Or c(ii + 1) = CChar("X") Or c(ii + 1) = CChar(",") Or c(ii + 1) = CChar("&")) Then
                    block_done = True
                ElseIf (bell_string_to_number(c(ii + 1)) < bell_string_to_number(c(ii))) Then
                    block_done = True
                End If
                ii += 1
            End While
            Me.notation.Add(New Notation(temp))
            block_idx += 1
        End While

        ' Now deal with there being a reversed part of the notation
        If rev_block <> -1 Then
            For ii = Me.notation.Count - 2 To rev_block Step -1
                Me.notation.Add(Me.notation(ii))
            Next
        End If
    End Sub
End Class

Public Class Notation
    Public notation As String

    Public Sub New(notation As String)
        Me.notation = notation
    End Sub

    ' Function to take notation and fill it out
    ' e.g. 4 becomes 14 and 45 becomes 145x, where x is the number of bells
    Public Function fill_notation(bells As Integer) As String
        Dim c As String = Me.notation
        Dim temp As String
        Dim string_len As Integer
        Dim ii As Integer = 0
        If c(0) = CChar("x") Or c(0) = CChar("X") Then
            Return CChar("X").ToString
        End If
        string_len = c.Count
        While (ii < string_len)
            Console.WriteLine(ii.ToString)
            If (Val(c(ii).ToString) Mod 2 = ii Mod 2) Then
                If ii = 0 Then
                    temp = "1"
                Else
                    temp = bell_number_to_string(bell_string_to_number(c(ii - 1)) + 1)
                    temp = (Val(c(ii - 1).ToString) + 1).ToString
                End If
                c = c.Insert(ii, temp)
                string_len += 1
            End If
            ii += 1
        End While

        ' Check that the last character in the notation is even. If it is not then add a character to the end.
        ' This is the same as checking the length of the list is even.
        If c.Count Mod 2 = 1 Then
            c = c.Insert(c.Count, bell_number_to_string(bells))
        End If
        Return c
    End Function

    ' Function to provide the hash of how bells move.
    ' The hash index is the position you are in the current row,
    ' the value at that location is your next position.
    ' e.g. change 1234->2143 has hash {2, 1, 3, 4}.
    ' the hash is the same for 4321->3412.
    ' It is important to remember that arrays are 0-indexed.
    Public Function change_hash(bells As Integer) As List(Of Integer)
        Dim hash As Integer()
        Dim c As String = Me.fill_notation(bells)
        Dim ii As Integer = 0
        Dim n_index As Integer = 0

        ReDim hash(bells - 1)
        ' Handle the simple x case
        If c(0) = "x" Or c(0) = "X" Then
            For ii = 1 To bells
                If ii Mod 2 = 0 Then
                    hash(ii - 1) = ii - 1
                Else
                    hash(ii - 1) = ii + 1
                End If
            Next
            GoTo EXIT_LABEL
        End If

        ' Now handle other cases. We do this by:
        ' 1 - check if this position is staying the same, if we are, then stay here and move to next bell
        ' 2 - if we are moving, then look at the last number. If both numbers are the same oddness then we move down, otherwise move up.
        ' NOTE: the function fill_notation should mean that all the implied numbers are in the notation so there is nothing to worry about there.
        For ii = 1 To bells
            If n_index < c.Count Then
                If ii = bell_string_to_number(c(n_index)) Then
                    hash(ii - 1) = ii
                    n_index += 1
                    Continue For
                End If
            End If
            If n_index = 0 Then
                If (0 = ii Mod 2) Then
                    hash(ii - 1) = ii - 1
                Else
                    hash(ii - 1) = ii + 1
                End If
                Continue For
            End If

            If (bell_string_to_number(c(n_index - 1)) Mod 2 = ii Mod 2) Then
                hash(ii - 1) = ii - 1
            Else
                hash(ii - 1) = ii + 1
            End If

        Next
EXIT_LABEL:
        Return hash.ToList
    End Function
End Class
