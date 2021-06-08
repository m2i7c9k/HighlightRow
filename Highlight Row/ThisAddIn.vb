Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_SheetSelectionChange(Sh As Object, Target As Range) Handles Application.SheetSelectionChange
        '
        Dim rng As Excel.Range = TryCast(Globals.ThisAddIn.Application.Cells, Excel.Range)
        Dim activeRng As Excel.Range = TryCast(Globals.ThisAddIn.Application.ActiveCell, Excel.Range)

        'cell value is copied to clipboard
        If My.Settings.copyCell = 1 Then
            If activeRng IsNot Nothing Then activeRng.Copy()
        End If

        rng.Cells.Interior.ColorIndex = XlColorIndex.xlColorIndexNone

        If My.Settings.highlightColumn = 1 Then
            Target.EntireColumn.Interior.ColorIndex = 37
        End If

        'row is always highlighted
        Target.EntireRow.Interior.ColorIndex = 37
        Target.Interior.ColorIndex = XlColorIndex.xlColorIndexNone

    End Sub

    Private Sub Application_WorkbookOpen(Wb As Workbook) Handles Application.WorkbookOpen
        'utorok, 08 júna 2021, 21:32:14
        If My.Settings.copyCell = 1 Then
            Globals.Ribbons.Ribbon.chbCopyCell.Checked = True
        Else
            Globals.Ribbons.Ribbon.chbCopyCell.Checked = False
        End If
    End Sub
End Class

