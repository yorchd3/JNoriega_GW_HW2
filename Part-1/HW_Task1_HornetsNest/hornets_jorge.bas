Attribute VB_Name = "Module1"
Sub Infestation()
'Part I: count number of hornets and add message box
Dim num_hornets As Integer



For i = 1 To 6
  For j = 1 To 7
    
    If Cells(i, j).Value = "Hornets" Then
     num_hornets = num_hornets + 1
     End If
     
    Next j
Next i

MsgBox ("num_hornets is:" & num_hornets)

End Sub

Sub infestation_partII()

'Part II: Modify the script such that it changes the word Hornets to "Bugs"

For i = 1 To 6
  For j = 1 To 7
    
    If Cells(i, j).Value = "Hornets" Then
     Cells(i, j).Value = "Bugs"
     End If
     
    Next j
Next i

End Sub

Sub infestation_partIII()

'Part III: replace hornets

Dim num_hornets As Integer
Dim num_bugs As Integer
Dim num_bees As Integer

num_bugs = Range("L2").Value
num_bees = Range("R2").Value

HornetsCount = 0

For i = 1 To 6
  For j = 1 To 7

    If Cells(i, j).Value = "Hornets" Then
      num_hornets = num_hornets + 1
      If (num_bugs > 0) Then
        Cells(i, j).Value = "Bugs"
        num_bugs = num_bugs - 1
      ElseIf (num_bees > 0) Then
        Cells(i, j).Value = "Bees"
        num_bees = num_bees - 1

      End If
    End If
  Next j
Next i

    If (Range("L2").Value + Range("R2").Value < num_hornets) Then
        MsgBox ("Oh no! We still have hornets... ")
    End If

End Sub
Sub makeallhornets()
For i = 1 To 6
  For j = 1 To 7
      Cells(i, j).Value = "Hornets"
    Next j
Next i
End Sub
