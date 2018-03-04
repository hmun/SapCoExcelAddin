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
        Me.ButtonReadPC = Me.Factory.CreateRibbonButton
        Me.ButtonReadAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostSK = Me.Factory.CreateRibbonButton
        Me.ButtonReadSK = Me.Factory.CreateRibbonButton
        Me.SAPCoOmPlanPer = Me.Factory.CreateRibbonGroup
        Me.ButtonReadPerAO = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerPC = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerSK = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerSK = Me.Factory.CreateRibbonButton
        Me.SAPActivityAlloc = Me.Factory.CreateRibbonGroup
        Me.ButtonActivityAllocCheck = Me.Factory.CreateRibbonButton
        Me.ButtonActivityAllocPost = Me.Factory.CreateRibbonButton
        Me.SAPRepstPrimCosts = Me.Factory.CreateRibbonGroup
        Me.ButtonRepstPrimCostsCheck = Me.Factory.CreateRibbonButton
        Me.ButtonRepstPrimCostsPost = Me.Factory.CreateRibbonButton
        Me.SAPCoLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapCo.SuspendLayout()
        Me.SAPCoOmPlan.SuspendLayout()
        Me.SAPCoOmPlanPer.SuspendLayout()
        Me.SAPActivityAlloc.SuspendLayout()
        Me.SAPRepstPrimCosts.SuspendLayout()
        Me.SAPCoLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCo
        '
        Me.SapCo.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.SapCo.Groups.Add(Me.SAPCoOmPlan)
        Me.SapCo.Groups.Add(Me.SAPCoOmPlanPer)
        Me.SapCo.Groups.Add(Me.SAPActivityAlloc)
        Me.SapCo.Groups.Add(Me.SAPRepstPrimCosts)
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
        'ButtonReadSK
        '
        Me.ButtonReadSK.Label = "Read SK"
        Me.ButtonReadSK.Name = "ButtonReadSK"
        Me.ButtonReadSK.ScreenTip = "Read Statistical Keyfigures"
        '
        'SAPCoOmPlanPer
        '
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerAO)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerPC)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerAI)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerAO)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerPC)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerAI)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerSK)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerSK)
        Me.SAPCoOmPlanPer.Label = "CO-OM Plan Periodic"
        Me.SAPCoOmPlanPer.Name = "SAPCoOmPlanPer"
        '
        'ButtonReadPerAO
        '
        Me.ButtonReadPerAO.Label = "Read Per AO"
        Me.ButtonReadPerAO.Name = "ButtonReadPerAO"
        Me.ButtonReadPerAO.ScreenTip = "Read Activity Output"
        '
        'ButtonReadPerPC
        '
        Me.ButtonReadPerPC.Label = "Read Per PC"
        Me.ButtonReadPerPC.Name = "ButtonReadPerPC"
        Me.ButtonReadPerPC.ScreenTip = "Read Primary Cost"
        '
        'ButtonReadPerAI
        '
        Me.ButtonReadPerAI.Label = "Read Per AI"
        Me.ButtonReadPerAI.Name = "ButtonReadPerAI"
        Me.ButtonReadPerAI.ScreenTip = "Read Activity Input"
        '
        'ButtonPostPerAO
        '
        Me.ButtonPostPerAO.Label = "Post Per AO"
        Me.ButtonPostPerAO.Name = "ButtonPostPerAO"
        Me.ButtonPostPerAO.ScreenTip = "Post Activity Output"
        '
        'ButtonPostPerPC
        '
        Me.ButtonPostPerPC.Label = "Post Per PC"
        Me.ButtonPostPerPC.Name = "ButtonPostPerPC"
        Me.ButtonPostPerPC.ScreenTip = "Post Primary Cost"
        '
        'ButtonPostPerAI
        '
        Me.ButtonPostPerAI.Label = "Post Per AI"
        Me.ButtonPostPerAI.Name = "ButtonPostPerAI"
        Me.ButtonPostPerAI.ScreenTip = "Post Activity Input"
        '
        'ButtonPostPerSK
        '
        Me.ButtonPostPerSK.Label = "Post Per SK"
        Me.ButtonPostPerSK.Name = "ButtonPostPerSK"
        Me.ButtonPostPerSK.ScreenTip = "Post Statistical Keyfigures"
        '
        'ButtonReadPerSK
        '
        Me.ButtonReadPerSK.Label = "Read Per SK"
        Me.ButtonReadPerSK.Name = "ButtonReadPerSK"
        Me.ButtonReadPerSK.ScreenTip = "Read Statistical Keyfigures"
        '
        'SAPActivityAlloc
        '
        Me.SAPActivityAlloc.Items.Add(Me.ButtonActivityAllocCheck)
        Me.SAPActivityAlloc.Items.Add(Me.ButtonActivityAllocPost)
        Me.SAPActivityAlloc.Label = "CO ActivityAlloc"
        Me.SAPActivityAlloc.Name = "SAPActivityAlloc"
        '
        'ButtonActivityAllocCheck
        '
        Me.ButtonActivityAllocCheck.Label = "ActivityAlloc Check"
        Me.ButtonActivityAllocCheck.Name = "ButtonActivityAllocCheck"
        Me.ButtonActivityAllocCheck.ScreenTip = "Check Activity Allocation Document"
        '
        'ButtonActivityAllocPost
        '
        Me.ButtonActivityAllocPost.Label = "ActivityAlloc Post"
        Me.ButtonActivityAllocPost.Name = "ButtonActivityAllocPost"
        Me.ButtonActivityAllocPost.ScreenTip = "Post Activity Allocation Document"
        '
        'SAPRepstPrimCosts
        '
        Me.SAPRepstPrimCosts.Items.Add(Me.ButtonRepstPrimCostsCheck)
        Me.SAPRepstPrimCosts.Items.Add(Me.ButtonRepstPrimCostsPost)
        Me.SAPRepstPrimCosts.Label = "CO RepstPrimCosts"
        Me.SAPRepstPrimCosts.Name = "SAPRepstPrimCosts"
        '
        'ButtonRepstPrimCostsCheck
        '
        Me.ButtonRepstPrimCostsCheck.Label = "RepstPrimCosts Check"
        Me.ButtonRepstPrimCostsCheck.Name = "ButtonRepstPrimCostsCheck"
        Me.ButtonRepstPrimCostsCheck.ScreenTip = "Check Repost Primary Costs Document"
        '
        'ButtonRepstPrimCostsPost
        '
        Me.ButtonRepstPrimCostsPost.Label = "RepstPrimCosts Post"
        Me.ButtonRepstPrimCostsPost.Name = "ButtonRepstPrimCostsPost"
        Me.ButtonRepstPrimCostsPost.ScreenTip = "Post Repost Primary Costs Document"
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
        'SapCoRibbon
        '
        Me.Name = "SapCoRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCo)
        Me.SapCo.ResumeLayout(False)
        Me.SapCo.PerformLayout()
        Me.SAPCoOmPlan.ResumeLayout(False)
        Me.SAPCoOmPlan.PerformLayout()
        Me.SAPCoOmPlanPer.ResumeLayout(False)
        Me.SAPCoOmPlanPer.PerformLayout()
        Me.SAPActivityAlloc.ResumeLayout(False)
        Me.SAPActivityAlloc.PerformLayout()
        Me.SAPRepstPrimCosts.ResumeLayout(False)
        Me.SAPRepstPrimCosts.PerformLayout()
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
    Friend WithEvents SAPActivityAlloc As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonActivityAllocCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonActivityAllocPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPRepstPrimCosts As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonRepstPrimCostsCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonRepstPrimCostsPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoOmPlanPer As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonReadPerAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerSK As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapCoRibbon() As SapCoRibbon
        Get
            Return Me.GetRibbon(Of SapCoRibbon)()
        End Get
    End Property
End Class
