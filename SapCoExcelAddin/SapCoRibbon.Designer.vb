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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapCoRibbon))
        Me.SapCo = Me.Factory.CreateRibbonTab
        Me.SAPActivityAlloc = Me.Factory.CreateRibbonGroup
        Me.ButtonActivityAllocCheck = Me.Factory.CreateRibbonButton
        Me.ButtonActivityAllocPost = Me.Factory.CreateRibbonButton
        Me.SAPRepstPrimCosts = Me.Factory.CreateRibbonGroup
        Me.ButtonRepstPrimCostsCheck = Me.Factory.CreateRibbonButton
        Me.ButtonRepstPrimCostsPost = Me.Factory.CreateRibbonButton
        Me.SAPManCostAlloc = Me.Factory.CreateRibbonGroup
        Me.ButtonManCostAllocCheck = Me.Factory.CreateRibbonButton
        Me.ButtonManCostAllocPost = Me.Factory.CreateRibbonButton
        Me.SAPCOPAActuals = Me.Factory.CreateRibbonGroup
        Me.ButtonCheckCostingBasedData = Me.Factory.CreateRibbonButton
        Me.ButtonPostCostingBasedData = Me.Factory.CreateRibbonButton
        Me.SAPCoLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapCo.SuspendLayout()
        Me.SAPActivityAlloc.SuspendLayout()
        Me.SAPRepstPrimCosts.SuspendLayout()
        Me.SAPManCostAlloc.SuspendLayout()
        Me.SAPCOPAActuals.SuspendLayout()
        Me.SAPCoLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCo
        '
        Me.SapCo.Groups.Add(Me.SAPActivityAlloc)
        Me.SapCo.Groups.Add(Me.SAPRepstPrimCosts)
        Me.SapCo.Groups.Add(Me.SAPManCostAlloc)
        Me.SapCo.Groups.Add(Me.SAPCOPAActuals)
        Me.SapCo.Groups.Add(Me.SAPCoLogon)
        Me.SapCo.Label = "SAP CO"
        Me.SapCo.Name = "SapCo"
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
        Me.ButtonActivityAllocCheck.Image = CType(resources.GetObject("ButtonActivityAllocCheck.Image"), System.Drawing.Image)
        Me.ButtonActivityAllocCheck.Label = "ActivityAlloc Check"
        Me.ButtonActivityAllocCheck.Name = "ButtonActivityAllocCheck"
        Me.ButtonActivityAllocCheck.ScreenTip = "Check Activity Allocation Document"
        Me.ButtonActivityAllocCheck.ShowImage = True
        '
        'ButtonActivityAllocPost
        '
        Me.ButtonActivityAllocPost.Image = CType(resources.GetObject("ButtonActivityAllocPost.Image"), System.Drawing.Image)
        Me.ButtonActivityAllocPost.Label = "ActivityAlloc Post"
        Me.ButtonActivityAllocPost.Name = "ButtonActivityAllocPost"
        Me.ButtonActivityAllocPost.ScreenTip = "Post Activity Allocation Document"
        Me.ButtonActivityAllocPost.ShowImage = True
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
        Me.ButtonRepstPrimCostsCheck.Image = CType(resources.GetObject("ButtonRepstPrimCostsCheck.Image"), System.Drawing.Image)
        Me.ButtonRepstPrimCostsCheck.Label = "RepstPrimCosts Check"
        Me.ButtonRepstPrimCostsCheck.Name = "ButtonRepstPrimCostsCheck"
        Me.ButtonRepstPrimCostsCheck.ScreenTip = "Check Repost Primary Costs Document"
        Me.ButtonRepstPrimCostsCheck.ShowImage = True
        '
        'ButtonRepstPrimCostsPost
        '
        Me.ButtonRepstPrimCostsPost.Image = CType(resources.GetObject("ButtonRepstPrimCostsPost.Image"), System.Drawing.Image)
        Me.ButtonRepstPrimCostsPost.Label = "RepstPrimCosts Post"
        Me.ButtonRepstPrimCostsPost.Name = "ButtonRepstPrimCostsPost"
        Me.ButtonRepstPrimCostsPost.ScreenTip = "Post Repost Primary Costs Document"
        Me.ButtonRepstPrimCostsPost.ShowImage = True
        '
        'SAPManCostAlloc
        '
        Me.SAPManCostAlloc.Items.Add(Me.ButtonManCostAllocCheck)
        Me.SAPManCostAlloc.Items.Add(Me.ButtonManCostAllocPost)
        Me.SAPManCostAlloc.Label = "CO ManCostAlloc"
        Me.SAPManCostAlloc.Name = "SAPManCostAlloc"
        '
        'ButtonManCostAllocCheck
        '
        Me.ButtonManCostAllocCheck.Image = CType(resources.GetObject("ButtonManCostAllocCheck.Image"), System.Drawing.Image)
        Me.ButtonManCostAllocCheck.Label = "ManCostAlloc Check"
        Me.ButtonManCostAllocCheck.Name = "ButtonManCostAllocCheck"
        Me.ButtonManCostAllocCheck.ScreenTip = "Check Manual Cost Allocation"
        Me.ButtonManCostAllocCheck.ShowImage = True
        '
        'ButtonManCostAllocPost
        '
        Me.ButtonManCostAllocPost.Image = CType(resources.GetObject("ButtonManCostAllocPost.Image"), System.Drawing.Image)
        Me.ButtonManCostAllocPost.Label = "ManCostAlloc Post"
        Me.ButtonManCostAllocPost.Name = "ButtonManCostAllocPost"
        Me.ButtonManCostAllocPost.ScreenTip = "Post Manual Cost Allocation"
        Me.ButtonManCostAllocPost.ShowImage = True
        '
        'SAPCOPAActuals
        '
        Me.SAPCOPAActuals.Items.Add(Me.ButtonCheckCostingBasedData)
        Me.SAPCOPAActuals.Items.Add(Me.ButtonPostCostingBasedData)
        Me.SAPCOPAActuals.Label = "CO-PA Actuals"
        Me.SAPCOPAActuals.Name = "SAPCOPAActuals"
        '
        'ButtonCheckCostingBasedData
        '
        Me.ButtonCheckCostingBasedData.Image = CType(resources.GetObject("ButtonCheckCostingBasedData.Image"), System.Drawing.Image)
        Me.ButtonCheckCostingBasedData.Label = "CostingBasedData Check"
        Me.ButtonCheckCostingBasedData.Name = "ButtonCheckCostingBasedData"
        Me.ButtonCheckCostingBasedData.ScreenTip = "Check posting of costing based data"
        Me.ButtonCheckCostingBasedData.ShowImage = True
        '
        'ButtonPostCostingBasedData
        '
        Me.ButtonPostCostingBasedData.Image = CType(resources.GetObject("ButtonPostCostingBasedData.Image"), System.Drawing.Image)
        Me.ButtonPostCostingBasedData.Label = "CostingBasedData Post"
        Me.ButtonPostCostingBasedData.Name = "ButtonPostCostingBasedData"
        Me.ButtonPostCostingBasedData.ScreenTip = "Post costing based data"
        Me.ButtonPostCostingBasedData.ShowImage = True
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
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapCoRibbon
        '
        Me.Name = "SapCoRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCo)
        Me.SapCo.ResumeLayout(False)
        Me.SapCo.PerformLayout()
        Me.SAPActivityAlloc.ResumeLayout(False)
        Me.SAPActivityAlloc.PerformLayout()
        Me.SAPRepstPrimCosts.ResumeLayout(False)
        Me.SAPRepstPrimCosts.PerformLayout()
        Me.SAPManCostAlloc.ResumeLayout(False)
        Me.SAPManCostAlloc.PerformLayout()
        Me.SAPCOPAActuals.ResumeLayout(False)
        Me.SAPCOPAActuals.PerformLayout()
        Me.SAPCoLogon.ResumeLayout(False)
        Me.SAPCoLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCo As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPCoLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPActivityAlloc As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonActivityAllocCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonActivityAllocPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPRepstPrimCosts As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonRepstPrimCostsCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonRepstPrimCostsPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCOPAActuals As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCheckCostingBasedData As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostCostingBasedData As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPManCostAlloc As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonManCostAllocCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonManCostAllocPost As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapCoRibbon() As SapCoRibbon
        Get
            Return Me.GetRibbon(Of SapCoRibbon)()
        End Get
    End Property
End Class
