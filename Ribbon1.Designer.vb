Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.EIB_TOOLS = Me.Factory.CreateRibbonGroup
        Me.cmd_EIB_ERROR = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.EIB_TOOLS.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.EIB_TOOLS)
        Me.Tab1.Label = "Nephology Tools"
        Me.Tab1.Name = "Tab1"
        '
        'EIB_TOOLS
        '
        Me.EIB_TOOLS.Items.Add(Me.cmd_EIB_ERROR)
        Me.EIB_TOOLS.Label = "EIB Tools"
        Me.EIB_TOOLS.Name = "EIB_TOOLS"
        '
        'cmd_EIB_ERROR
        '
        Me.cmd_EIB_ERROR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.cmd_EIB_ERROR.Image = CType(resources.GetObject("cmd_EIB_ERROR.Image"), System.Drawing.Image)
        Me.cmd_EIB_ERROR.Label = "EIB Error Extractor"
        Me.cmd_EIB_ERROR.Name = "cmd_EIB_ERROR"
        Me.cmd_EIB_ERROR.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.EIB_TOOLS.ResumeLayout(False)
        Me.EIB_TOOLS.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents EIB_TOOLS As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents cmd_EIB_ERROR As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
