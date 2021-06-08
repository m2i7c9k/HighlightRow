Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    End Sub

    Private Sub chbHighlightColumn_Click(sender As Object, e As RibbonControlEventArgs) Handles chbHighlightColumn.Click
        'utorok, 08 júna 2021, 21:05:47
        If chbHighlightColumn.Checked Then
            My.Settings.highlightColumn = 1
        Else
            My.Settings.highlightColumn = 0
        End If
        My.Settings.Save()
    End Sub

    Private Sub chbCopyCell_Click(sender As Object, e As RibbonControlEventArgs) Handles chbCopyCell.Click
        'utorok, 08 júna 2021, 21:08:13
        If chbCopyCell.Checked Then
            My.Settings.copyCell = 1
        Else
            My.Settings.copyCell = 0
        End If
        My.Settings.Save()
    End Sub

End Class
