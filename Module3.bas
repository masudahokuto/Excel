Attribute VB_Name = "Module3"
Sub Sample2()
    
    
  Dim ctrRow
  
  Dim i
  
  ctrRow = 8
  
  For i = 8 To 12
  
  
    If (Range("E" & i).Value = Range("D2").Value) Then
        
        Range("L" & ctrRow).Value = Range("C" & i).Value
        Range("M" & ctrRow).Value = Range("D" & i).Value
        Range("N" & ctrRow).Value = Range("E" & i).Value
        Range("O" & ctrRow).Value = Range("F" & i).Value
        Range("P" & ctrRow).Value = Range("G" & i).Value
        Range("Q" & ctrRow).Value = Range("H" & i).Value
        
        ctrRow = ctrRow + 1
    End If
  
  Next i
End Sub
