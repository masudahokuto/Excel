Attribute VB_Name = "Module3"
Sub Sample2()
    
    
  Dim ctrRow
  
  Dim i
  
  Dim endRow
  
  ctrRow = 8
  
  endRow = Range("C8").End(xlDown).Row
  
  For i = 8 To endRow
  
  
    If (Range("E" & i).Value = Range("D2").Value) Then
        
        Worksheets("Sample_plactice").Range("C" & ctrRow).Value = Range("C" & i).Value
        Worksheets("Sample_plactice").Range("D" & ctrRow).Value = Range("D" & i).Value
        Worksheets("Sample_plactice").Range("E" & ctrRow).Value = Range("E" & i).Value
        Worksheets("Sample_plactice").Range("F" & ctrRow).Value = Range("F" & i).Value
        Worksheets("Sample_plactice").Range("G" & ctrRow).Value = Range("G" & i).Value
        Worksheets("Sample_plactice").Range("H" & ctrRow).Value = Range("H" & i).Value
        
        ctrRow = ctrRow + 1
    End If
  
  Next i
End Sub
