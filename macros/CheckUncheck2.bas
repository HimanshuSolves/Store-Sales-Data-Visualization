Attribute VB_Name = "CheckUncheck2"
Sub SingleSelectCheckbox()
    Dim ws As Worksheet
    Dim checkCells As Variant
    Dim clickedCheckbox As String
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("DATA VISUALIZATION")
    
    ' List of cells linked to checkboxes
    checkCells = Array("L65", "L66", "L67", "L68", "L69")
    
    ' Determine which checkbox called this macro
    ' Application.Caller returns the name of the shape (checkbox)
    clickedCheckbox = Application.Caller
    
    ' Find which cell is linked to the clicked checkbox
    Dim cb As Shape
    Dim linkedCell As String
    
    For Each cb In ws.Shapes
        If cb.Name = clickedCheckbox Then
            linkedCell = cb.ControlFormat.Value ' 1=Checked, -4146=Unchecked, 0=Mixed
            ' Unfortunately ControlFormat.Value gives state but not linked cell directly
            ' So we use the linked cell address stored in ControlFormat.LinkedCell
            linkedCell = cb.ControlFormat.linkedCell
            Exit For
        End If
    Next cb
    
    If linkedCell = "" Then Exit Sub ' No linked cell found
    
    Application.EnableEvents = False
    
    ' Uncheck all linked cells except the one clicked
    For i = LBound(checkCells) To UBound(checkCells)
        If ws.Range(checkCells(i)).Address(False, False) <> Replace(linkedCell, "$", "") Then
            ws.Range(checkCells(i)).Value = False
        Else
            ws.Range(checkCells(i)).Value = True
        End If
    Next i
    
    Application.EnableEvents = True
End Sub

