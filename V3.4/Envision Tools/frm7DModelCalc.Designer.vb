<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm7DModelCalc
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm7DModelCalc))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.pnlExcel = New System.Windows.Forms.Panel()
        Me.rdbSelect = New System.Windows.Forms.RadioButton()
        Me.rdbEnvision = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.tbxExcel = New System.Windows.Forms.TextBox()
        Me.btnExcelFile = New System.Windows.Forms.Button()
        Me.rdbTotal = New System.Windows.Forms.RadioButton()
        Me.rdbExisting = New System.Windows.Forms.RadioButton()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.chkUseSelected = New System.Windows.Forms.CheckBox()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.barStatus = New System.Windows.Forms.ToolStripProgressBar()
        Me.cmbScenarios = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.pnlExcel.SuspendLayout()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmbScenarios)
        Me.Panel1.Controls.Add(Me.pnlExcel)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.tbxExcel)
        Me.Panel1.Controls.Add(Me.btnExcelFile)
        Me.Panel1.Controls.Add(Me.rdbTotal)
        Me.Panel1.Controls.Add(Me.rdbExisting)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.chkUseSelected)
        Me.Panel1.Controls.Add(Me.cmbLayers)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(423, 234)
        Me.Panel1.TabIndex = 56
        '
        'pnlExcel
        '
        Me.pnlExcel.Controls.Add(Me.rdbSelect)
        Me.pnlExcel.Controls.Add(Me.rdbEnvision)
        Me.pnlExcel.Location = New System.Drawing.Point(6, 134)
        Me.pnlExcel.Name = "pnlExcel"
        Me.pnlExcel.Size = New System.Drawing.Size(313, 54)
        Me.pnlExcel.TabIndex = 201
        '
        'rdbSelect
        '
        Me.rdbSelect.AutoSize = True
        Me.rdbSelect.Checked = True
        Me.rdbSelect.Location = New System.Drawing.Point(3, 34)
        Me.rdbSelect.Name = "rdbSelect"
        Me.rdbSelect.Size = New System.Drawing.Size(202, 17)
        Me.rdbSelect.TabIndex = 7
        Me.rdbSelect.TabStop = True
        Me.rdbSelect.Text = "Select an HH Travel Model Excel File"
        Me.rdbSelect.UseVisualStyleBackColor = True
        '
        'rdbEnvision
        '
        Me.rdbEnvision.AutoSize = True
        Me.rdbEnvision.Location = New System.Drawing.Point(3, 4)
        Me.rdbEnvision.Name = "rdbEnvision"
        Me.rdbEnvision.Size = New System.Drawing.Size(150, 17)
        Me.rdbEnvision.TabIndex = 6
        Me.rdbEnvision.Text = "Current Envision Excel File"
        Me.rdbEnvision.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 114)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(187, 16)
        Me.Label1.TabIndex = 198
        Me.Label1.Text = "HH Travel Model Excel File:"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Location = New System.Drawing.Point(4, 107)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(411, 4)
        Me.Panel2.TabIndex = 196
        '
        'tbxExcel
        '
        Me.tbxExcel.AllowDrop = True
        Me.tbxExcel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxExcel.Location = New System.Drawing.Point(6, 193)
        Me.tbxExcel.Multiline = True
        Me.tbxExcel.Name = "tbxExcel"
        Me.tbxExcel.ReadOnly = True
        Me.tbxExcel.Size = New System.Drawing.Size(347, 28)
        Me.tbxExcel.TabIndex = 193
        '
        'btnExcelFile
        '
        Me.btnExcelFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExcelFile.BackColor = System.Drawing.Color.Transparent
        Me.btnExcelFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnExcelFile.Image = CType(resources.GetObject("btnExcelFile.Image"), System.Drawing.Image)
        Me.btnExcelFile.Location = New System.Drawing.Point(359, 178)
        Me.btnExcelFile.Name = "btnExcelFile"
        Me.btnExcelFile.Size = New System.Drawing.Size(56, 48)
        Me.btnExcelFile.TabIndex = 8
        Me.btnExcelFile.UseVisualStyleBackColor = False
        '
        'rdbTotal
        '
        Me.rdbTotal.AutoSize = True
        Me.rdbTotal.Checked = True
        Me.rdbTotal.Location = New System.Drawing.Point(123, 4)
        Me.rdbTotal.Name = "rdbTotal"
        Me.rdbTotal.Size = New System.Drawing.Size(141, 17)
        Me.rdbTotal.TabIndex = 2
        Me.rdbTotal.TabStop = True
        Me.rdbTotal.Text = "Calculate Scenario Total"
        Me.rdbTotal.UseVisualStyleBackColor = True
        '
        'rdbExisting
        '
        Me.rdbExisting.AutoSize = True
        Me.rdbExisting.Location = New System.Drawing.Point(9, 4)
        Me.rdbExisting.Name = "rdbExisting"
        Me.rdbExisting.Size = New System.Drawing.Size(108, 17)
        Me.rdbExisting.TabIndex = 1
        Me.rdbExisting.Text = "Calculate Existing"
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
        'chkUseSelected
        '
        Me.chkUseSelected.AutoSize = True
        Me.chkUseSelected.Enabled = False
        Me.chkUseSelected.Location = New System.Drawing.Point(153, 76)
        Me.chkUseSelected.Name = "chkUseSelected"
        Me.chkUseSelected.Size = New System.Drawing.Size(189, 17)
        Me.chkUseSelected.TabIndex = 5
        Me.chkUseSelected.Text = "Only Calculate Selected Feature(s)"
        Me.chkUseSelected.UseVisualStyleBackColor = True
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(153, 49)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLayers.Size = New System.Drawing.Size(262, 21)
        Me.cmbLayers.TabIndex = 4
        Me.cmbLayers.Text = "Select a layer"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(3, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(144, 16)
        Me.Label5.TabIndex = 152
        Me.Label5.Text = "Neighborhood Layer:"
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus, Me.barStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 238)
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
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 28)
        Me.lblStatus.Visible = False
        '
        'barStatus
        '
        Me.barStatus.Name = "barStatus"
        Me.barStatus.Size = New System.Drawing.Size(100, 28)
        '
        'cmbScenarios
        '
        Me.cmbScenarios.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbScenarios.FormattingEnabled = True
        Me.cmbScenarios.Items.AddRange(New Object() {"Scenario 1", "Scenario 2", "Scenario 3", "Scenario 4", "Scenario 5"})
        Me.cmbScenarios.Location = New System.Drawing.Point(275, 3)
        Me.cmbScenarios.Name = "cmbScenarios"
        Me.cmbScenarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbScenarios.Size = New System.Drawing.Size(141, 21)
        Me.cmbScenarios.TabIndex = 3
        Me.cmbScenarios.Text = "Scenario 1"
        '
        'frm7DModelCalc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 269)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frm7DModelCalc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "7D Model Calculator"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnlExcel.ResumeLayout(False)
        Me.pnlExcel.PerformLayout()
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents chkUseSelected As System.Windows.Forms.CheckBox
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents barStatus As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents rdbTotal As System.Windows.Forms.RadioButton
    Friend WithEvents rdbExisting As System.Windows.Forms.RadioButton
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents tbxExcel As System.Windows.Forms.TextBox
    Friend WithEvents btnExcelFile As System.Windows.Forms.Button
    Friend WithEvents rdbSelect As System.Windows.Forms.RadioButton
    Friend WithEvents rdbEnvision As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlExcel As System.Windows.Forms.Panel
    Friend WithEvents cmbScenarios As System.Windows.Forms.ComboBox
End Class
