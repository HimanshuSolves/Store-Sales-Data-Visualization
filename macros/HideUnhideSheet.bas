Attribute VB_Name = "HideUnhideSheet"

Sub ToggleSheetsVisibility()
    Dim wsAccess As Worksheet
    Dim i As Long
    Dim sheetName As String
    Dim isChecked As Variant
    Dim targetSheet As Worksheet
    Dim lastRow As Long

    Set wsAccess = ThisWorkbook.Sheets("ACCESS SHEETS")
    lastRow = wsAccess.Cells(wsAccess.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    For i = 1 To lastRow
        sheetName = wsAccess.Cells(i, 1).Value
        isChecked = wsAccess.Cells(i, 2).Value

        If Len(sheetName) > 0 Then
            Set targetSheet = ThisWorkbook.Sheets(sheetName)
            If Not targetSheet Is Nothing Then
                If isChecked = True Then
                    targetSheet.Visible = xlSheetVisible
                Else
                    targetSheet.Visible = xlSheetHidden
                End If
            End If
        End If
        Set targetSheet = Nothing
    Next i
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
