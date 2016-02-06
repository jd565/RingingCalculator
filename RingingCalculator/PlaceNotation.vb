Public Class PlaceNotation
    Public full_notation As String
    Public main_block As New List(Of Notation)
    Public lead_end As New List(Of Notation)
    Public bob_end As New List(Of Notation)
    Public single_end As New List(Of Notation)
    Public bells As Integer

    Public Sub New(notation As String, Optional bells As Integer = 0)
        Me.full_notation = notation
        Me.bells = bells
        Me.parse()
    End Sub

    'Function to find whether this is a standard notation or a microSiril one
    Public Sub parse()
        'we assume that if the place notation starts with a small letter that is not x then we are
        'parsing something from the microsiril library.
        If Me.full_notation(0) >= CChar("a") And Me.full_notation(0) <= CChar("z") And Me.full_notation(0) <> CChar("x") Then
            Console.WriteLine("microSiril notation")
            Me.parse_microsiril()
            Exit Sub
        End If

        'However this may still be microSiril but start with e.g. 3z. Check if the string contains a z.
        'There is not a reason to include a z in the place notation other than this.
        If Me.full_notation.Contains("z") Then
            Console.WriteLine("microSiril notation")
            Me.parse_microsiril()
            Exit Sub
        End If
        Console.WriteLine("standard notation")
        Me.parse_standard()
    End Sub

    'Function to parse a microsiril notation
    Public Sub parse_microsiril()
        Dim temp_pn As PlaceNotation
        Dim c As String = Me.full_notation
        Dim temp As String
        Dim ii As Integer = 0
        Dim block_done As Boolean
        Dim block_idx As Integer = 0
        Dim rev_block As Integer = -1
        Dim separating_chars As Char() = {",", ".", " ", "+"}
        Me.main_block.Clear()
        Dim notation_len As Integer
        Dim nop_block As Boolean = False

        If Me.bells = 0 Then
            Console.WriteLine("no bells, cannot parse")
            Exit Sub
        End If

        ' Get the length of the starting notation.
        notation_len = c.Count

        ' microSiril doesn't start with the place notation so move ii to the start.
        ' the notations has the lead head block then a space then the notation.
        ii = c.IndexOf(" ")
        ii += 1

        ' Now move through the notation and extract each block
        ' Blocks are separated by commas, or if the next input is smaller.
        While (ii < notation_len)
            block_done = False
            temp = ""
            If c(ii) = "-" Then
                block_done = True
                temp = "X"
                ii += 1
            End If
            If (separating_chars.Contains(c(ii))) Then
                ii += 1
            End If
            While Not block_done
                If c(ii) = "&" Then
                    ii += 1
                    rev_block = block_idx
                    ' This is a block in itself, but we don't want to update anything based on it
                    nop_block = True
                    Exit While
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

                If separating_chars.Contains(c(ii + 1)) Or c(ii + 1) = "-" Then
                    Console.WriteLine("found separator")
                    block_done = True
                ElseIf (bell_string_to_number(c(ii + 1)) < bell_string_to_number(c(ii))) Then
                    block_done = True
                End If
                ii += 1
            End While
            If Not nop_block Then
                Me.main_block.Add(New Notation(temp))
                block_idx += 1
            Else
                nop_block = False
            End If
        End While

        ' The lead end and bob/single data is stored in the format specifier at the start
        If c.Contains("z") Then
            ii = 0
            temp = ""
            While c(ii) <> "z"
                temp += c(ii)
                ii += 1
            End While
            If ii = 0 Then
                'This is for an odd method that has le 3.
                temp += "3"
            End If
            temp_pn = New PlaceNotation(temp)
            Me.lead_end = temp_pn.main_block
        ElseIf (c(0) >= "a" And c(0) <= "f") Or (c(0) >= "p" And c(0) <= "q") Then
            temp = "2"
            temp_pn = New PlaceNotation(temp)
            Me.lead_end = temp_pn.main_block
            temp = "4"
            temp_pn = New PlaceNotation(temp)
            Me.bob_end = temp_pn.main_block
            temp = "1234"
            temp_pn = New PlaceNotation(temp)
            Me.single_end = temp_pn.main_block
        ElseIf (c(0) >= "g" And c(0) <= "m") Or c(0) = "r" Or c(0) = "s" Then
            temp = bell_number_to_string(Me.bells)
            temp_pn = New PlaceNotation(temp)
            Me.lead_end = temp_pn.main_block
            temp = bell_number_to_string(Me.bells - 2)
            temp_pn = New PlaceNotation(temp)
            Me.bob_end = temp_pn.main_block
            temp = bell_number_to_string(Me.bells - 2) + bell_number_to_string(Me.bells - 1) + bell_number_to_string(Me.bells)
            temp_pn = New PlaceNotation(temp)
            Me.single_end = temp_pn.main_block
        End If

        If rev_block <> -1 Then
            For ii = Me.main_block.Count - 2 To rev_block Step -1
                Me.main_block.Add(Me.main_block(ii))
            Next
        End If

    End Sub

    ' Function to parse the full notation and split it into smaller parts
    Public Sub parse_standard()
        Dim temp_pn As PlaceNotation
        Dim c As String = Me.full_notation
        Dim temp As String
        Dim ii As Integer = 0
        Dim block_done As Boolean
        Dim block_idx As Integer = 0
        Dim rev_block As Integer = -1
        Dim separating_chars As Char() = {",", ".", " "}
        Dim x As Char() = {"x", "X", "-"}
        Me.main_block.Clear()
        Dim notation_len As Integer
        Dim nop_block As Boolean = False

        ' Get the length of the starting notation. This is maybe ended with a space
        notation_len = c.IndexOf(" ")
        If notation_len = -1 Then notation_len = c.Count
        ' Now move through the notation and extract each block
        ' Blocks are separated by commas, or if the next input is smaller.
        While (ii < notation_len)
            block_done = False
            temp = ""
            If (x.Contains(c(ii))) Then
                block_done = True
                temp = "X"
                ii += 1
            End If
            If (separating_chars.Contains(c(ii))) Then
                ii += 1
            End If
            While Not block_done
                If c(ii) = "&" Then
                    ii += 1
                    rev_block = block_idx
                    ' This is a block in itself, but we don't want to update anything based on it
                    nop_block = True
                    Exit While
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

                If (separating_chars.Contains(c(ii + 1)) Or x.Contains(c(ii + 1))) Then
                    Console.WriteLine("found separator")
                    block_done = True
                ElseIf (bell_string_to_number(c(ii + 1)) < bell_string_to_number(c(ii))) Then
                    block_done = True
                End If
                ii += 1
            End While
            If Not nop_block Then
                Me.main_block.Add(New Notation(temp))
                block_idx += 1
            Else
                nop_block = False
            End If
        End While

        ' If we have reached the end of the count then drop out here
        If ii = c.Count Then
            Exit Sub
        End If

        ' Otherwise, the method is reversed and has a specified lead end, or there are bobs and singles defined
        ' Check for a specified lead end
        ii = c.IndexOf("le")
        If ii <> -1 Then
            ' We have found the lead end specifier. Extract it here
            ii += "le ".Count
            temp = ""
            While (ii < c.Count)
                If (c(ii) = " ") Then
                    Exit While
                End If
                temp += c(ii)
                ii += 1
            End While
            temp_pn = New PlaceNotation(temp)
            Me.lead_end = temp_pn.main_block
            If rev_block = -1 Then
                ' We only expet the lead end block if we are reversing the notation
                rev_block = 0
            End If
        End If

        ' Check for a specified bob
        ii = c.IndexOf("b")
        If ii <> -1 Then
            ' We have found the lead end specifier. Extract it here
            ii += "b ".Count
            temp = ""
            While (ii < c.Count)
                If (c(ii) = " ") Then
                    Exit While
                End If
                temp += c(ii)
                ii += 1
            End While
            temp_pn = New PlaceNotation(temp)
            Me.bob_end = temp_pn.main_block
        End If

        ' Check for a specified single
        ii = c.IndexOf("s")
        If ii <> -1 Then
            ' We have found the lead end specifier. Extract it here
            ii += "s ".Count
            temp = ""
            While (ii < c.Count)
                If (c(ii) = " ") Then
                    Exit While
                End If
                temp += c(ii)
                ii += 1
            End While
            temp_pn = New PlaceNotation(temp)
            Me.single_end = temp_pn.main_block
        End If

        If rev_block <> -1 Then
            For ii = Me.main_block.Count - 2 To rev_block Step -1
                Me.main_block.Add(Me.main_block(ii))
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
        If c(0) = "x" Or c(0) = "X" Or c(0) = "-" Then
            Return "X".ToString
        End If
        string_len = c.Count
        While (ii < string_len)
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

        ' Check that the last character in the notation is same oddness as the number of bells.
        ' If it Is Not Then add a character to the end.
        ' This is the same as checking the length of the list is even.
        If c.Count Mod 2 <> bells Mod 2 Then
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
