Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub Application_SheetSelectionChange(Sh As Object, Target As Range) Handles Application.SheetSelectionChange
        'utorok, 08 júna 2021, 22:11:26
        Dim rng As Excel.Range = TryCast(Globals.ThisAddIn.Application.Cells, Excel.Range)
        Dim activeRng As Excel.Range = TryCast(Globals.ThisAddIn.Application.ActiveCell, Excel.Range)

        'vypnutie označenia riadku bez nutnosti oddinštalovania add-inu
        If My.Settings.turnOffHighlight = 1 Then
            Exit Sub
        End If

        'cell value is copied to clipboard
        If My.Settings.copyCell = 1 Then
            If activeRng IsNot Nothing Then activeRng.Copy()
        End If

        rng.Cells.Interior.ColorIndex = XlColorIndex.xlColorIndexNone

        If My.Settings.highlightColumn = 1 Then
            Target.EntireColumn.Interior.Color = My.Settings.highlightColor '37
        End If

        'row is always highlighted
        Target.EntireRow.Interior.Color = My.Settings.highlightColor '37
        Target.Interior.ColorIndex = XlColorIndex.xlColorIndexNone

    End Sub

End Class

