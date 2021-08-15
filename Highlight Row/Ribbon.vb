Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Drawing

Public Class Ribbon

    Public Property Color As Color
    Private ReadOnly Application As Object

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        'utorok, 08 júna 2021, 21:32:14
        If My.Settings.copyCell = 1 Then
            Globals.Ribbons.Ribbon.chbCopyCell.Checked = True
        Else
            Globals.Ribbons.Ribbon.chbCopyCell.Checked = False
        End If

        If My.Settings.highlightColumn = 1 Then
            Globals.Ribbons.Ribbon.chbHighlightColumn.Checked = True
        Else
            Globals.Ribbons.Ribbon.chbHighlightColumn.Checked = False
        End If

        If My.Settings.turnOffHighlight = 0 Then
            'Add-inis On
            Globals.Ribbons.Ribbon.tgbTurnOffAddin.Checked = True
            With tgbTurnOffAddin
                .Image = My.Resources.TurnOn
                .Label = "Turned On"
            End With
        Else
            'Add-in is Off
            Globals.Ribbons.Ribbon.tgbTurnOffAddin.Checked = False
            With tgbTurnOffAddin
                .Image = My.Resources.TurnOff
                .Label = "Turned Off"
            End With
        End If

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
            Dim xlApp As Excel.Application = Globals.ThisAddIn.Application
            xlApp.CutCopyMode = False
        End If
        My.Settings.Save()
    End Sub
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles cmdColorDialog.Click
        'utorok, 08 júna 2021, 23:01:07
        Dim ColorDialog As New ColorDialog With {
            .Color = My.Settings.highlightColor
        }

        ColorDialog.FullOpen = True

        If (ColorDialog.ShowDialog() = DialogResult.OK) Then
            My.Settings.highlightColor = ColorDialog.Color
            My.Settings.Save()
        End If
    End Sub

    Private Sub tgbTurnOffAddin_Click(sender As Object, e As RibbonControlEventArgs) Handles tgbTurnOffAddin.Click
        'streda, 9 júna 2021, 7: 59:51
        If tgbTurnOffAddin.Checked Then
            'Add-in Off
            My.Settings.turnOffHighlight = 0
            With tgbTurnOffAddin
                .Image = My.Resources.TurnOn
                .Label = "Turned On"
            End With
        Else
            'Add-in On
            My.Settings.turnOffHighlight = 1

            Dim xlApp As Excel.Application = Globals.ThisAddIn.Application
            Dim rng As Excel.Range = TryCast(Globals.ThisAddIn.Application.Cells, Excel.Range)
            xlApp.CutCopyMode = False
            rng.Cells.Interior.ColorIndex = Microsoft.Office.Core.XlColorIndex.xlColorIndexNone

            With tgbTurnOffAddin
                .Image = My.Resources.TurnOff
                .Label = "Turned Off"
            End With

        End If
        My.Settings.Save()
    End Sub

End Class
