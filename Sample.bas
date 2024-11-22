Attribute VB_Name = "Module2"
Sub Sample1()

    Dim ctrRow
    
    ctrRow = 8
    
    If (Range("D2").Value = Range("E8").Value) Then
        
        Range("L" & ctrRow).Value = Range("C8").Value
        Range("M" & ctrRow).Value = Range("D8").Value
        Range("N" & ctrRow).Value = Range("E8").Value
        Range("O" & ctrRow).Value = Range("F8").Value
        Range("P" & ctrRow).Value = Range("G8").Value
        Range("Q" & ctrRow).Value = Range("H8").Value
        
        ctrRow = ctrRow + 1
    End If
    
    If (Range("D2").Value = Range("E9").Value) Then
        
        Range("L" & ctrRow).Value = Range("C9").Value
        Range("M" & ctrRow).Value = Range("D9").Value
        Range("N" & ctrRow).Value = Range("E9").Value
        Range("O" & ctrRow).Value = Range("F9").Value
        Range("P" & ctrRow).Value = Range("G9").Value
        Range("Q" & ctrRow).Value = Range("H9").Value
        
        ctrRow = ctrRow + 1
    End If

    If (Range("D2").Value = Range("E10").Value) Then
        
        Range("L" & ctrRow).Value = Range("C10").Value
        Range("M" & ctrRow).Value = Range("D10").Value
        Range("N" & ctrRow).Value = Range("E10").Value
        Range("O" & ctrRow).Value = Range("F10").Value
        Range("P" & ctrRow).Value = Range("G10").Value
        Range("Q" & ctrRow).Value = Range("H10").Value
        
        ctrRow = ctrRow + 1
    End If
    
    If (Range("D2").Value = Range("E11").Value) Then
        
        Range("L" & ctrRow).Value = Range("C11").Value
        Range("M" & ctrRow).Value = Range("D11").Value
        Range("N" & ctrRow).Value = Range("E11").Value
        Range("O" & ctrRow).Value = Range("F11").Value
        Range("P" & ctrRow).Value = Range("G11").Value
        Range("Q" & ctrRow).Value = Range("H11").Value
        
        ctrRow = ctrRow + 1
    End If
    
    If (Range("D2").Value = Range("E12").Value) Then
        
        Range("L" & ctrRow).Value = Range("C12").Value
        Range("M" & ctrRow).Value = Range("D12").Value
        Range("N" & ctrRow).Value = Range("E12").Value
        Range("O" & ctrRow).Value = Range("F12").Value
        Range("P" & ctrRow).Value = Range("G12").Value
        Range("Q" & ctrRow).Value = Range("H12").Value
        
        ctrRow = ctrRow + 1
    End If
End Sub
