Option Explicit

Dim lastSheetName As String
Dim currentSheetName As String

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    lastSheetName = currentSheetName
    currentSheetName = Sh.Name
End Sub


Sub SwitchToLastActivedSheet()
    If lastSheetName <> Empty Then

    Sheets(lastSheetName).Activate

    End If
End Sub
