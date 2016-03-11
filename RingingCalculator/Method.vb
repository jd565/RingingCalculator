Public Class Method
    Public place_notation As PlaceNotation
    Public bells As Integer
    Public rows As List(Of Row)
    Public changes_per_lead As Integer
    Public comp As Composition
    Private comp_loc As Integer
    Private do_call As Boolean = False

    Public Sub New(pn As String, bells As Integer)
        Me.place_notation = New PlaceNotation(pn, bells)
        Me.bells = bells
        Me.changes_per_lead = Me.place_notation.main_block.Count + Me.place_notation.lead_end.Count
    End Sub

    Public Sub add_composition(full_comp As String)
        Me.comp = New Composition(full_comp)
    End Sub

    ' Function to generate the rows of this method, based on the notation
    Public Sub generate()
        Dim ii As Integer = 0
        Me.rows = New List(Of Row)
        Me.create_first_row()

        ' We use a change time of 1s between bells, and use this to indicate
        ' their position within a row
        ' Now generate rows until we get back to rounds
        While Not Me.rows.Last.is_this_rounds() And ii < 10000
            Me.generate_next_row()
            ii += 1
        End While
    End Sub

    ''' <summary>
    ''' Check whether to add a call at this point.
    ''' This function must be called with the rows set up to
    ''' just before the start of the lead end block.
    ''' </summary>
    Private Sub maybe_do_a_call()
        Dim le_change As Integer()
        Dim hash As List(Of Integer)
        Dim ii As Integer = 0
        Dim position As Integer
        Dim call_found As Boolean = False
        Dim next_call As RingingCall
        ReDim le_change(bells - 1)
        If Me.comp Is Nothing Then
            Exit Sub
        End If
        If Me.comp_loc = Me.comp.composition.Count Then
            Exit Sub
        End If
        next_call = Me.comp.composition(Me.comp_loc)
        For Each call_type In Me.place_notation.call_ends
            If call_type.call_string.Equals(next_call.call_to_make) Then
                hash = Notation.full_list_hash(Me.bells, call_type.notation)
                For Each bell In hash
                    le_change(bell - 1) = Me.rows.Last.bells(ii).bell
                    ii += 1
                Next
                call_found = True
                Exit For
            End If
        Next
        If Not call_found Then
            Throw New System.Exception("No call matching next desired.")
        End If
        ' Check that the tenor bell is in the right place
        Select Case next_call.location
            Case "W"
                position = Me.bells - 1
            Case "H"
                position = Me.bells
            Case "M"
                position = Me.bells - 2
            Case Else
                position = Val(next_call.location)
        End Select
        If le_change(position - 1) = Me.bells Then
            Me.do_call = True
        End If
    End Sub

    ''' <summary>
    ''' Generate the next row of the method.
    ''' </summary>
    Private Sub generate_next_row()
        Dim hash As New List(Of Integer)
        Dim index As Integer
        Dim next_row As New Row
        Dim change_time As ChangeTime

        ' Get the index of the notation to use now.
        ' The notation to go from the first row to the second is at index 1
        index = (Me.rows.Count) Mod Me.changes_per_lead
        If index = Me.place_notation.main_block.Count Then
            Me.maybe_do_a_call()
        End If
        If index >= Me.place_notation.main_block.Count Then
            index -= Me.place_notation.main_block.Count
            If Me.do_call Then
                For Each call_type In Me.place_notation.call_ends
                    If call_type.call_string.Equals(Me.comp.composition(Me.comp_loc).call_to_make) Then
                        hash = call_type.notation(index).change_hash(Me.bells)
                        Exit For
                    End If
                Next
                If index = Me.place_notation.lead_end.Count - 1 Then
                    Me.do_call = False
                    Me.comp_loc += 1
                End If
            Else
                hash = Me.place_notation.lead_end(index).change_hash(Me.bells)
            End If
        Else
                hash = Me.place_notation.main_block(index).change_hash(Me.bells)
        End If

        For ii As Integer = 0 To hash.Count - 1
            change_time = New ChangeTime(Me.rows.Last.bells(ii).bell,
                                         Me.rows.Count + 1,
                                         New DateTime(hash(ii)))
            next_row.add(change_time)
        Next
        next_row.sort()
        Me.rows.Add(next_row)
    End Sub

    ' Function to put rounds as the first row
    Private Sub create_first_row()
        Dim hash As List(Of Integer)
        Dim index As Integer
        Dim next_row As New Row
        Dim change_time As ChangeTime

        index = 0
        hash = Me.place_notation.main_block(index).change_hash(Me.bells)

        For ii As Integer = 0 To hash.Count - 1
            change_time = New ChangeTime(ii + 1,
                                         1,
                                         New DateTime(hash(ii)))
            next_row.add(change_time)
        Next
        next_row.sort()
        Me.rows.Add(next_row)
    End Sub
End Class
