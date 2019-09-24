Attribute VB_Name = "Module1"
Sub lotto1()

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (lastrow)

For i = 2 To lastrow

    If Cells(i, 3).Value = 3957481 Then
    Cells(2, 6).Value = Cells(i, 1)
    Cells(2, 7).Value = Cells(i, 2)
    Cells(2, 8).Value = Cells(i, 3)
      
    ElseIf Cells(i, 3).Value = 5865187 Then
    Cells(3, 6).Value = Cells(i, 1)
    Cells(3, 7).Value = Cells(i, 2)
    Cells(3, 8).Value = Cells(i, 3)
    
    ElseIf Cells(i, 3).Value = 2817729 Then
    Cells(4, 6).Value = Cells(i, 1)
    Cells(4, 7).Value = Cells(i, 2)
    Cells(4, 8).Value = Cells(i, 3)
    
    End If

Next i

MsgBox ("Congratulations to first place winner: " & Cells(2, 6).Value & " " & Cells(2, 7).Value)


End Sub
