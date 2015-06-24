Sub New_game()
Dim n As Single, answers(1 To 4) As Single, j As Single, k As Single

For n = 1 To 4
    answers(n) = Int((6 - 1 + 1) * Rnd + 1)
Next n
'generating the random string which represent the answers


For n = 1 To 4
    Cells(100, n).Value = answers(n)
Next n
'storing the random string into cells out of view

For n = 4 To 14 Step 2
    For j = 4 To 7
        Cells(n, j).Clear
        Cells(n, j).Interior.ColorIndex = 16
    Next j
Next n

For n = 5 To 15 Step 2
    For j = 4 To 7
        Cells(n, j).Clear
        Cells(n, j).Interior.ColorIndex = 15
    Next j
Next n

For n = 4 To 15
    For j = 8 To 9
        Cells(n, j).Clear
        Cells(n, j).Interior.ColorIndex = 16
    Next j
Next n

For n = 4 To 7
    Cells(2, n).Clear
    Cells(2, n).Interior.ColorIndex = 16
Next n

        



'clearing the previous game


Cells(101, 1).Value = 0
End Sub
Sub guess()
Dim attempt(1 To 4) As Single, answers(1 To 4) As Single, n As Single, j As Single, k As Single, color As Single, location As Single
Dim blue_count_a As Single, blue_count_g As Single, red_count_a As Single, red_count_g As Single, green_count_a As Single, green_count_g As Single
Dim yellow_count_a As Single, yellow_count_g As Single, orange_count_a As Single, orange_count_g As Single, purple_count_a As Single, purple_count_g As Single
    
For n = 1 To 4
    answers(n) = Cells(100, n).Value
Next n
'storing the answers into an array


k = Cells(101, 1).Value

For n = 1 To 4
    If Cells(15 - k, n + 3).Value = "b" Then
        attempt(n) = 1
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 5
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ElseIf Cells(15 - k, n + 3).Value = "r" Then
        attempt(n) = 2
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 3
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ElseIf Cells(15 - k, n + 3).Value = "g" Then
        attempt(n) = 3
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 4
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ElseIf Cells(15 - k, n + 3).Value = "y" Then
        attempt(n) = 4
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 6
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ElseIf Cells(15 - k, n + 3).Value = "o" Then
        attempt(n) = 5
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 46
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ElseIf Cells(15 - k, n + 3).Value = "p" Then
        attempt(n) = 6
        Cells(15 - k, n + 3).Clear
        Cells(15 - k, n + 3).Interior.ColorIndex = 21
        Cells(15 - k, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(15 - k, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
Next n
'saving numbers into attempt()
'clearing the letter out of the cell
'changing the color of the cell to the color typed in


For n = 1 To 4
    If attempt(n) = answers(n) Then location = location + 1
Next n
'finding the amount of correct locations


For n = 1 To 4
    If attempt(n) = 1 Then blue_count_g = blue_count_g + 1
    If attempt(n) = 2 Then red_count_g = red_count_g + 1
    If attempt(n) = 3 Then green_count_g = green_count_g + 1
    If attempt(n) = 4 Then yellow_count_g = yellow_count_g + 1
    If attempt(n) = 5 Then orange_count_g = orange_count_g + 1
    If attempt(n) = 6 Then purple_count_g = purple_count_g + 1
    
    If answers(n) = 1 Then blue_count_a = blue_count_a + 1
    If answers(n) = 2 Then red_count_a = red_count_a + 1
    If answers(n) = 3 Then green_count_a = green_count_a + 1
    If answers(n) = 4 Then yellow_count_a = yellow_count_a + 1
    If answers(n) = 5 Then orange_count_a = orange_count_a + 1
    If answers(n) = 6 Then purple_count_a = purple_count_a + 1
Next n
 'counting the amout of colors in each array
 
  
If blue_count_g >= 1 And blue_count_a >= 1 And blue_count_g <= blue_count_a Then
    color = color + blue_count_g
ElseIf blue_count_g > blue_count_a Then
    color = color + blue_count_a
End If

If red_count_g >= 1 And red_count_a >= 1 And red_count_g <= red_count_a Then
    color = color + red_count_g
ElseIf red_count_g > red_count_a Then
    color = color + red_count_a
End If

If green_count_g >= 1 And green_count_a >= 1 And green_count_g <= green_count_a Then
    color = color + green_count_g
ElseIf green_count_g > green_count_a Then
    color = color + green_count_a
End If

If yellow_count_g >= 1 And yellow_count_a >= 1 And yellow_count_g <= yellow_count_a Then
    color = color + yellow_count_g
ElseIf yellow_count_g > yellow_count_a Then
    color = color + yellow_count_a
End If

If orange_count_g >= 1 And orange_count_a >= 1 And orange_count_g <= orange_count_a Then
    color = color + orange_count_g
ElseIf orange_count_g > orange_count_a Then
    color = color + orange_count_a
End If

If purple_count_g >= 1 And purple_count_a >= 1 And purple_count_g <= purple_count_a Then
    color = color + purple_count_g
ElseIf purple_count_g > purple_count_a Then
    color = color + purple_count_a
End If
'adding up the common colors


Cells(15 - k, 8).Value = location
Cells(15 - k, 9).Value = color - location

Cells(15 - k, 8).Font.Size = 14
Cells(15 - k, 8).Font.Bold = True
Cells(15 - k, 8).Font.color = &HFF&

Cells(15 - k, 9).Font.Size = 14
Cells(15 - k, 9).Font.Bold = True
Cells(15 - k, 9).Font.ThemeColor = 3
'putting the hints in the cells
    



If location = 4 Then
    For n = 1 To 4
        If Cells(100, n).Value = 1 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 5
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 2 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 3
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 3 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 4
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 4 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 6
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 5 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 46
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 6 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 21
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End If
    Next n
    MsgBox ("Winner!")
End If

k = k + 1

If k = 12 Then
    For n = 1 To 4
        If Cells(100, n).Value = 1 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 5
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 2 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 3
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 3 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 4
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 4 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 6
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 5 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 46
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ElseIf Cells(100, n).Value = 6 Then
            Cells(2, n + 3).Clear
            Cells(2, n + 3).Interior.ColorIndex = 21
            Cells(2, n + 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(2, n + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End If
    Next n
    MsgBox ("Sorry" & vbNewLine & vbNewLine & "Try again")
End If

    
Cells(101, 1).Value = k
'counting variable for the guess button


End Sub
