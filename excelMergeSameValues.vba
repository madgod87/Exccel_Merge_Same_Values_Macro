Sub MergeSameValues()
    Dim myRange As Range
    Dim r As Long, c As Long
    Dim i As Long
    Dim temp As Variant
    
    ' Set the range of cells to the current selection
    Set myRange = Selection
    
    ' --- Failsafe: Create a backup of the sheet ---
    Dim originalSheet As Worksheet
    Set originalSheet = ActiveSheet
    
    originalSheet.Copy After:=originalSheet
    ' The copy becomes ActiveSheet. Rename it safely.
    On Error Resume Next
    ActiveSheet.Name = Left(originalSheet.Name, 15) & "_Bak_" & Format(Now, "hhmmss")
    On Error GoTo 0
    
    ' Activate the original sheet so the macro runs on the intended data
    originalSheet.Activate
    ' ----------------------------------------------
    
    ' Disable alerts and screen updating for performance and to suppress merge warnings
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Heuristic to determine merge direction:
    ' If selection has more rows than columns (or equal), assume Vertical merge per column.
    ' If selection has more columns than rows, assume Horizontal merge per row.
    Dim isVertical As Boolean
    If myRange.Rows.Count >= myRange.Columns.Count Then
        isVertical = True
    Else
        isVertical = False
    End If
    
    If isVertical Then
        ' Vertical Mode: Process each column independently
        Dim colRange As Range
        For c = 1 To myRange.Columns.Count
            Set colRange = myRange.Columns(c)
            i = 1 ' Start row index for the current merge block
            temp = colRange.Cells(i, 1).Value
            
            For r = 2 To colRange.Rows.Count
                If colRange.Cells(r, 1).Value = temp Then
                    ' If the value matches the tracked value, merge from start (i) to current (r)
                    Range(colRange.Cells(i, 1), colRange.Cells(r, 1)).Merge
                Else
                    ' If value differs, reset start index and update temp value
                    i = r
                    temp = colRange.Cells(r, 1).Value
                End If
            Next r
        Next c
    Else
        ' Horizontal Mode: Process each row independently
        Dim rowRange As Range
        For r = 1 To myRange.Rows.Count
            Set rowRange = myRange.Rows(r)
            i = 1 ' Start column index for the current merge block
            temp = rowRange.Cells(1, i).Value
            
            For c = 2 To rowRange.Columns.Count
                If rowRange.Cells(1, c).Value = temp Then
                    ' If the value matches the tracked value, merge from start (i) to current (c)
                    Range(rowRange.Cells(1, i), rowRange.Cells(1, c)).Merge
                Else
                    ' If value differs, reset start index and update temp value
                    i = c
                    temp = rowRange.Cells(1, c).Value
                End If
            Next c
        Next r
    End If
    
    ' Re-enable alerts and screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


