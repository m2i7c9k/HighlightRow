Partial Class Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.grpHighlighter = Me.Factory.CreateRibbonGroup
        Me.chbHighlightColumn = Me.Factory.CreateRibbonCheckBox
        Me.chbCopyCell = Me.Factory.CreateRibbonCheckBox
        Me.Tab1.SuspendLayout()
        Me.grpHighlighter.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.grpHighlighter)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'grpHighlighter
        '
        Me.grpHighlighter.Items.Add(Me.chbHighlightColumn)
        Me.grpHighlighter.Items.Add(Me.chbCopyCell)
        Me.grpHighlighter.Label = "Highlighter"
        Me.grpHighlighter.Name = "grpHighlighter"
        '
        'chbHighlightColumn
        '
        Me.chbHighlightColumn.Label = "Highlight Column"
        Me.chbHighlightColumn.Name = "chbHighlightColumn"
        '
        'chbCopyCell
        '
        Me.chbCopyCell.Label = "Copy Cell"
        Me.chbCopyCell.Name = "chbCopyCell"
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.grpHighlighter.ResumeLayout(False)
        Me.grpHighlighter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpHighlighter As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents chbHighlightColumn As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents chbCopyCell As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
