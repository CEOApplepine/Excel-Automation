Sub GenerateReport()
    ' Refresh all PivotTables
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    MsgBox "Report generated successfully!", vbInformation
End Sub
