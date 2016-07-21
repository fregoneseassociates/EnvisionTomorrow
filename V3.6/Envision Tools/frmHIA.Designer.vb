<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHIA
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
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

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHIA))
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.btnLayerWarning = New System.Windows.Forms.Button()
    Me.txtEmail = New System.Windows.Forms.TextBox()
    Me.lblEmail = New System.Windows.Forms.Label()
    Me.cmbScenarios = New System.Windows.Forms.ComboBox()
    Me.lblHIAExcelFile = New System.Windows.Forms.Label()
    Me.Panel2 = New System.Windows.Forms.Panel()
    Me.tbxExcel = New System.Windows.Forms.TextBox()
    Me.btnExcelFile = New System.Windows.Forms.Button()
    Me.rdbScenario = New System.Windows.Forms.RadioButton()
    Me.rdbExisting = New System.Windows.Forms.RadioButton()
    Me.Panel6 = New System.Windows.Forms.Panel()
    Me.cmbLayers = New System.Windows.Forms.ComboBox()
    Me.lblCBGLayer = New System.Windows.Forms.Label()
    Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
    Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
    Me.btnRun = New System.Windows.Forms.ToolStripButton()
    Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
    Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.Panel1.SuspendLayout()
    Me.ToolStrip_InfoTab3.SuspendLayout()
    Me.SuspendLayout()
    '
    'Panel1
    '
    Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel1.Controls.Add(Me.btnLayerWarning)
    Me.Panel1.Controls.Add(Me.txtEmail)
    Me.Panel1.Controls.Add(Me.lblEmail)
    Me.Panel1.Controls.Add(Me.cmbScenarios)
    Me.Panel1.Controls.Add(Me.lblHIAExcelFile)
    Me.Panel1.Controls.Add(Me.Panel2)
    Me.Panel1.Controls.Add(Me.tbxExcel)
    Me.Panel1.Controls.Add(Me.btnExcelFile)
    Me.Panel1.Controls.Add(Me.rdbScenario)
    Me.Panel1.Controls.Add(Me.rdbExisting)
    Me.Panel1.Controls.Add(Me.Panel6)
    Me.Panel1.Controls.Add(Me.cmbLayers)
    Me.Panel1.Controls.Add(Me.lblCBGLayer)
    Me.Panel1.Location = New System.Drawing.Point(1, 1)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(423, 198)
    Me.Panel1.TabIndex = 56
    '
    'btnLayerWarning
    '
    Me.btnLayerWarning.Image = Global.Envision_Tools.My.Resources.Resources.GenericWarning16
    Me.btnLayerWarning.Location = New System.Drawing.Point(168, 48)
    Me.btnLayerWarning.Name = "btnLayerWarning"
    Me.btnLayerWarning.Size = New System.Drawing.Size(20, 23)
    Me.btnLayerWarning.TabIndex = 201
    Me.ToolTip1.SetToolTip(Me.btnLayerWarning, "Click for more information")
    Me.btnLayerWarning.UseVisualStyleBackColor = True
    Me.btnLayerWarning.Visible = False
    '
    'txtEmail
    '
    Me.txtEmail.AllowDrop = True
    Me.txtEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtEmail.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtEmail.Location = New System.Drawing.Point(57, 158)
    Me.txtEmail.Name = "txtEmail"
    Me.txtEmail.Size = New System.Drawing.Size(358, 21)
    Me.txtEmail.TabIndex = 200
    Me.txtEmail.Text = "info@frego.com"
    '
    'lblEmail
    '
    Me.lblEmail.AutoSize = True
    Me.lblEmail.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblEmail.Location = New System.Drawing.Point(6, 161)
    Me.lblEmail.Name = "lblEmail"
    Me.lblEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblEmail.Size = New System.Drawing.Size(43, 13)
    Me.lblEmail.TabIndex = 199
    Me.lblEmail.Text = "Email:"
    '
    'cmbScenarios
    '
    Me.cmbScenarios.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmbScenarios.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmbScenarios.FormattingEnabled = True
    Me.cmbScenarios.Items.AddRange(New Object() {"Scenario 1", "Scenario 2", "Scenario 3", "Scenario 4", "Scenario 5"})
    Me.cmbScenarios.Location = New System.Drawing.Point(275, 5)
    Me.cmbScenarios.Name = "cmbScenarios"
    Me.cmbScenarios.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmbScenarios.Size = New System.Drawing.Size(140, 21)
    Me.cmbScenarios.TabIndex = 3
    Me.cmbScenarios.Text = "Scenario 1"
    '
    'lblHIAExcelFile
    '
    Me.lblHIAExcelFile.AutoSize = True
    Me.lblHIAExcelFile.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblHIAExcelFile.Location = New System.Drawing.Point(6, 98)
    Me.lblHIAExcelFile.Name = "lblHIAExcelFile"
    Me.lblHIAExcelFile.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblHIAExcelFile.Size = New System.Drawing.Size(273, 13)
    Me.lblHIAExcelFile.TabIndex = 198
    Me.lblHIAExcelFile.Text = "Health Assessment Model Excel File (optional):"
    '
    'Panel2
    '
    Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel2.Location = New System.Drawing.Point(4, 84)
    Me.Panel2.Name = "Panel2"
    Me.Panel2.Size = New System.Drawing.Size(411, 4)
    Me.Panel2.TabIndex = 196
    '
    'tbxExcel
    '
    Me.tbxExcel.AllowDrop = True
    Me.tbxExcel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbxExcel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbxExcel.Location = New System.Drawing.Point(6, 116)
    Me.tbxExcel.Name = "tbxExcel"
    Me.tbxExcel.ReadOnly = True
    Me.tbxExcel.Size = New System.Drawing.Size(347, 21)
    Me.tbxExcel.TabIndex = 193
    '
    'btnExcelFile
    '
    Me.btnExcelFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnExcelFile.BackColor = System.Drawing.Color.Transparent
    Me.btnExcelFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
    Me.btnExcelFile.Image = CType(resources.GetObject("btnExcelFile.Image"), System.Drawing.Image)
    Me.btnExcelFile.Location = New System.Drawing.Point(359, 101)
    Me.btnExcelFile.Name = "btnExcelFile"
    Me.btnExcelFile.Size = New System.Drawing.Size(56, 48)
    Me.btnExcelFile.TabIndex = 8
    Me.btnExcelFile.UseVisualStyleBackColor = False
    '
    'rdbScenario
    '
    Me.rdbScenario.AutoSize = True
    Me.rdbScenario.Checked = True
    Me.rdbScenario.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.rdbScenario.Location = New System.Drawing.Point(189, 6)
    Me.rdbScenario.Name = "rdbScenario"
    Me.rdbScenario.Size = New System.Drawing.Size(80, 17)
    Me.rdbScenario.TabIndex = 2
    Me.rdbScenario.TabStop = True
    Me.rdbScenario.Text = "Scenario:"
    Me.rdbScenario.UseVisualStyleBackColor = True
    '
    'rdbExisting
    '
    Me.rdbExisting.AutoSize = True
    Me.rdbExisting.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.rdbExisting.Location = New System.Drawing.Point(9, 6)
    Me.rdbExisting.Name = "rdbExisting"
    Me.rdbExisting.Size = New System.Drawing.Size(133, 17)
    Me.rdbExisting.TabIndex = 1
    Me.rdbExisting.Text = "Existing Conditions"
    Me.rdbExisting.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    Me.rdbExisting.UseVisualStyleBackColor = True
    '
    'Panel6
    '
    Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel6.Location = New System.Drawing.Point(4, 34)
    Me.Panel6.Name = "Panel6"
    Me.Panel6.Size = New System.Drawing.Size(411, 4)
    Me.Panel6.TabIndex = 190
    '
    'cmbLayers
    '
    Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmbLayers.CausesValidation = False
    Me.cmbLayers.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmbLayers.FormattingEnabled = True
    Me.cmbLayers.Items.AddRange(New Object() {"Scenario 1", "Scenario 2", "Scenario 3", "Scenario 4", "Scenario 5"})
    Me.cmbLayers.Location = New System.Drawing.Point(188, 49)
    Me.cmbLayers.Name = "cmbLayers"
    Me.cmbLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmbLayers.Size = New System.Drawing.Size(227, 21)
    Me.cmbLayers.TabIndex = 4
    Me.cmbLayers.Text = "Scenario 1"
    '
    'lblCBGLayer
    '
    Me.lblCBGLayer.AutoSize = True
    Me.lblCBGLayer.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblCBGLayer.Location = New System.Drawing.Point(6, 52)
    Me.lblCBGLayer.Name = "lblCBGLayer"
    Me.lblCBGLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblCBGLayer.Size = New System.Drawing.Size(164, 13)
    Me.lblCBGLayer.TabIndex = 152
    Me.lblCBGLayer.Text = "Census Block Group Layer:"
    '
    'ToolStrip_InfoTab3
    '
    Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
    Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus})
    Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 202)
    Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
    Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
    Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(424, 31)
    Me.ToolStrip_InfoTab3.TabIndex = 57
    Me.ToolStrip_InfoTab3.Text = "ToolStrip4"
    '
    'ToolStripSeparator11
    '
    Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
    Me.ToolStripSeparator11.Size = New System.Drawing.Size(6, 31)
    '
    'btnRun
    '
    Me.btnRun.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnRun.Image = CType(resources.GetObject("btnRun.Image"), System.Drawing.Image)
    Me.btnRun.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.btnRun.Name = "btnRun"
    Me.btnRun.Size = New System.Drawing.Size(63, 28)
    Me.btnRun.Text = "Run"
    '
    'ToolStripSeparator9
    '
    Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
    Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 31)
    '
    'lblStatus
    '
    Me.lblStatus.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
    Me.lblStatus.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblStatus.Margin = New System.Windows.Forms.Padding(6, 1, 0, 2)
    Me.lblStatus.Name = "lblStatus"
    Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblStatus.Size = New System.Drawing.Size(69, 28)
    Me.lblStatus.Text = "status..."
    '
    'frmHIA
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(424, 233)
    Me.Controls.Add(Me.ToolStrip_InfoTab3)
    Me.Controls.Add(Me.Panel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmHIA"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Health Assessment Model"
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.ToolStrip_InfoTab3.ResumeLayout(False)
    Me.ToolStrip_InfoTab3.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
  Friend WithEvents lblCBGLayer As System.Windows.Forms.Label
  Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
  Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
  Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
  Friend WithEvents rdbScenario As System.Windows.Forms.RadioButton
  Friend WithEvents rdbExisting As System.Windows.Forms.RadioButton
  Friend WithEvents Panel6 As System.Windows.Forms.Panel
  Friend WithEvents Panel2 As System.Windows.Forms.Panel
  Friend WithEvents tbxExcel As System.Windows.Forms.TextBox
  Friend WithEvents btnExcelFile As System.Windows.Forms.Button
  Friend WithEvents lblHIAExcelFile As System.Windows.Forms.Label
  Friend WithEvents cmbScenarios As System.Windows.Forms.ComboBox
  Friend WithEvents lblEmail As System.Windows.Forms.Label
  Friend WithEvents txtEmail As System.Windows.Forms.TextBox
  Friend WithEvents btnLayerWarning As System.Windows.Forms.Button
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
