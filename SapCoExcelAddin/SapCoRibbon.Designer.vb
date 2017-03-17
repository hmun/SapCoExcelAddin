Partial Class SapCoRibbon
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
        Me.SapCo = Me.Factory.CreateRibbonTab
        Me.SAPCoOmPlan = Me.Factory.CreateRibbonGroup
        Me.ButtonReadAO = Me.Factory.CreateRibbonButton
        Me.SAPCoLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.ButtonReadPC = Me.Factory.CreateRibbonButton
        Me.ButtonReadAI = Me.Factory.CreateRibbonButton
        Me.ButtonReadSK = Me.Factory.CreateRibbonButton
        Me.ButtonPostAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostSK = Me.Factory.CreateRibbonButton
        Me.SapCo.SuspendLayout()
        Me.SAPCoOmPlan.SuspendLayout()
        Me.SAPCoLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCo
        '
        Me.SapCo.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.SapCo.Groups.Add(Me.SAPCoOmPlan)
        Me.SapCo.Groups.Add(Me.SAPCoLogon)
        Me.SapCo.Label = "SAP CO"
        Me.SapCo.Name = "SapCo"
        '
        'SAPCoOmPlan
        '
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadAO)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadPC)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadAI)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostAO)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostPC)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostAI)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadSK)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostSK)
        Me.SAPCoOmPlan.Label = "CO-OM Plan"
        Me.SAPCoOmPlan.Name = "SAPCoOmPlan"
        '
        'ButtonReadAO
        '
        Me.ButtonReadAO.Label = "Read AO"
        Me.ButtonReadAO.Name = "ButtonReadAO"
        Me.ButtonReadAO.ScreenTip = "Read Activity Output"
        '
        'SAPCoLogon
        '
        Me.SAPCoLogon.Items.Add(Me.ButtonLogon)
        Me.SAPCoLogon.Items.Add(Me.ButtonLogoff)
        Me.SAPCoLogon.Label = "Logon"
        Me.SAPCoLogon.Name = "SAPCoLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        '
        'ButtonReadPC
        '
        Me.ButtonReadPC.Label = "Read PC"
        Me.ButtonReadPC.Name = "ButtonReadPC"
        Me.ButtonReadPC.ScreenTip = "Read Primary Cost"
        '
        'ButtonReadAI
        '
        Me.ButtonReadAI.Label = "Read AI"
        Me.ButtonReadAI.Name = "ButtonReadAI"
        Me.ButtonReadAI.ScreenTip = "Read Activity Input"
        '
        'ButtonReadSK
        '
        Me.ButtonReadSK.Label = "Read SK"
        Me.ButtonReadSK.Name = "ButtonReadSK"
        Me.ButtonReadSK.ScreenTip = "Read Statistical Keyfigures"
        '
        'ButtonPostAO
        '
        Me.ButtonPostAO.Label = "Post AO"
        Me.ButtonPostAO.Name = "ButtonPostAO"
        Me.ButtonPostAO.ScreenTip = "Post Activity Output"
        '
        'ButtonPostPC
        '
        Me.ButtonPostPC.Label = "Post PC"
        Me.ButtonPostPC.Name = "ButtonPostPC"
        Me.ButtonPostPC.ScreenTip = "Post Primary Cost"
        '
        'ButtonPostAI
        '
        Me.ButtonPostAI.Label = "Post AI"
        Me.ButtonPostAI.Name = "ButtonPostAI"
        Me.ButtonPostAI.ScreenTip = "Post Activity Input"
        '
        'ButtonPostSK
        '
        Me.ButtonPostSK.Label = "Post SK"
        Me.ButtonPostSK.Name = "ButtonPostSK"
        Me.ButtonPostSK.ScreenTip = "Post Statistical Keyfigures"
        '
        'SapCoRibbon
        '
        Me.Name = "SapCoRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCo)
        Me.SapCo.ResumeLayout(False)
        Me.SapCo.PerformLayout()
        Me.SAPCoOmPlan.ResumeLayout(False)
        Me.SAPCoOmPlan.PerformLayout()
        Me.SAPCoLogon.ResumeLayout(False)
        Me.SAPCoLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCo As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPCoOmPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonReadAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostSK As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapCoRibbon() As SapCoRibbon
        Get
            Return Me.GetRibbon(Of SapCoRibbon)()
        End Get
    End Property
End Class
