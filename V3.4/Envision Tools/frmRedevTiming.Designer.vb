<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRedevTiming
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRedevTiming))
        Me.cmbParcelLayers = New System.Windows.Forms.ComboBox
        Me.lblParcelLyr = New System.Windows.Forms.Label
        Me.cmbYearBuilt = New System.Windows.Forms.ComboBox
        Me.lblYearBuilt = New System.Windows.Forms.Label
        Me.cmbImproveValue = New System.Windows.Forms.ComboBox
        Me.cmbLandValue = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblGridCellSize = New System.Windows.Forms.Label
        Me.tbxCurrentYear = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbxLifespan = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbxAppreciation = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.pnlInputs = New System.Windows.Forms.Panel
        Me.rdbBottomQurtile = New System.Windows.Forms.RadioButton
        Me.rdbDepreciation = New System.Windows.Forms.RadioButton
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator
        Me.btnRun = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.btnApplyLegend = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator
        Me.barProgress = New System.Windows.Forms.ToolStripProgressBar
        Me.lblSqFt = New System.Windows.Forms.Label
        Me.cmbSqFt = New System.Windows.Forms.ComboBox
        Me.LblBreaks = New System.Windows.Forms.Label
        Me.tbxBreaks = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.pnlInputs.SuspendLayout()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbParcelLayers
        '
        Me.cmbParcelLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbParcelLayers.FormattingEnabled = True
        Me.cmbParcelLayers.Location = New System.Drawing.Point(168, 49)
        Me.cmbParcelLayers.Name = "cmbParcelLayers"
        Me.cmbParcelLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbParcelLayers.Size = New System.Drawing.Size(414, 21)
        Me.cmbParcelLayers.TabIndex = 85
        '
        'lblParcelLyr
        '
        Me.lblParcelLyr.AutoSize = True
        Me.lblParcelLyr.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblParcelLyr.Location = New System.Drawing.Point(13, 51)
        Me.lblParcelLyr.Name = "lblParcelLyr"
        Me.lblParcelLyr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblParcelLyr.Size = New System.Drawing.Size(142, 16)
        Me.lblParcelLyr.TabIndex = 84
        Me.lblParcelLyr.Text = "Select Parcel Layer:"
        '
        'cmbYearBuilt
        '
        Me.cmbYearBuilt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbYearBuilt.FormattingEnabled = True
        Me.cmbYearBuilt.Location = New System.Drawing.Point(326, 95)
        Me.cmbYearBuilt.Name = "cmbYearBuilt"
        Me.cmbYearBuilt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbYearBuilt.Size = New System.Drawing.Size(256, 21)
        Me.cmbYearBuilt.TabIndex = 87
        '
        'lblYearBuilt
        '
        Me.lblYearBuilt.AutoSize = True
        Me.lblYearBuilt.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYearBuilt.Location = New System.Drawing.Point(85, 96)
        Me.lblYearBuilt.Name = "lblYearBuilt"
        Me.lblYearBuilt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblYearBuilt.Size = New System.Drawing.Size(224, 16)
        Me.lblYearBuilt.TabIndex = 86
        Me.lblYearBuilt.Text = "Select ""Effective Year Built"" field"
        '
        'cmbImproveValue
        '
        Me.cmbImproveValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbImproveValue.FormattingEnabled = True
        Me.cmbImproveValue.Location = New System.Drawing.Point(326, 128)
        Me.cmbImproveValue.Name = "cmbImproveValue"
        Me.cmbImproveValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbImproveValue.Size = New System.Drawing.Size(256, 21)
        Me.cmbImproveValue.TabIndex = 89
        '
        'cmbLandValue
        '
        Me.cmbLandValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLandValue.FormattingEnabled = True
        Me.cmbLandValue.Location = New System.Drawing.Point(326, 163)
        Me.cmbLandValue.Name = "cmbLandValue"
        Me.cmbLandValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLandValue.Size = New System.Drawing.Size(256, 21)
        Me.cmbLandValue.TabIndex = 91
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(140, 164)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(169, 16)
        Me.Label3.TabIndex = 90
        Me.Label3.Text = "Select ""Land Value"" field"
        '
        'lblGridCellSize
        '
        Me.lblGridCellSize.AutoSize = True
        Me.lblGridCellSize.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGridCellSize.Location = New System.Drawing.Point(31, 16)
        Me.lblGridCellSize.Name = "lblGridCellSize"
        Me.lblGridCellSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGridCellSize.Size = New System.Drawing.Size(190, 16)
        Me.lblGridCellSize.TabIndex = 93
        Me.lblGridCellSize.Text = "Enter Current Year (4 digit)"
        Me.lblGridCellSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxCurrentYear
        '
        Me.tbxCurrentYear.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxCurrentYear.Location = New System.Drawing.Point(229, 12)
        Me.tbxCurrentYear.Name = "tbxCurrentYear"
        Me.tbxCurrentYear.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxCurrentYear.Size = New System.Drawing.Size(344, 20)
        Me.tbxCurrentYear.TabIndex = 92
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(65, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(156, 16)
        Me.Label4.TabIndex = 95
        Me.Label4.Text = "Enter Building Lifespan"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxLifespan
        '
        Me.tbxLifespan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxLifespan.Location = New System.Drawing.Point(229, 51)
        Me.tbxLifespan.Name = "tbxLifespan"
        Me.tbxLifespan.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxLifespan.Size = New System.Drawing.Size(344, 20)
        Me.tbxLifespan.TabIndex = 94
        Me.tbxLifespan.Text = "50"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(7, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(215, 16)
        Me.Label5.TabIndex = 97
        Me.Label5.Text = "Enter Annual Land Appreciation"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxAppreciation
        '
        Me.tbxAppreciation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxAppreciation.Location = New System.Drawing.Point(229, 82)
        Me.tbxAppreciation.Name = "tbxAppreciation"
        Me.tbxAppreciation.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxAppreciation.Size = New System.Drawing.Size(265, 20)
        Me.tbxAppreciation.TabIndex = 96
        Me.tbxAppreciation.Text = "2.00"
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(802, 52)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(44, 16)
        Me.Label8.TabIndex = 101
        Me.Label8.Text = "years"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(500, 83)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(81, 16)
        Me.Label9.TabIndex = 102
        Me.Label9.Text = "% per year"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 129)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(296, 16)
        Me.Label2.TabIndex = 104
        Me.Label2.Text = "Select ""Improvement or Building Value"" field"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.LblBreaks)
        Me.Panel1.Controls.Add(Me.tbxBreaks)
        Me.Panel1.Controls.Add(Me.lblSqFt)
        Me.Panel1.Controls.Add(Me.cmbSqFt)
        Me.Panel1.Controls.Add(Me.pnlInputs)
        Me.Panel1.Controls.Add(Me.rdbBottomQurtile)
        Me.Panel1.Controls.Add(Me.rdbDepreciation)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.lblParcelLyr)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cmbParcelLayers)
        Me.Panel1.Controls.Add(Me.lblYearBuilt)
        Me.Panel1.Controls.Add(Me.cmbYearBuilt)
        Me.Panel1.Controls.Add(Me.cmbImproveValue)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cmbLandValue)
        Me.Panel1.Location = New System.Drawing.Point(2, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(595, 365)
        Me.Panel1.TabIndex = 105
        '
        'pnlInputs
        '
        Me.pnlInputs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlInputs.Controls.Add(Me.lblGridCellSize)
        Me.pnlInputs.Controls.Add(Me.Label4)
        Me.pnlInputs.Controls.Add(Me.tbxLifespan)
        Me.pnlInputs.Controls.Add(Me.tbxAppreciation)
        Me.pnlInputs.Controls.Add(Me.Label5)
        Me.pnlInputs.Controls.Add(Me.tbxCurrentYear)
        Me.pnlInputs.Controls.Add(Me.Label8)
        Me.pnlInputs.Controls.Add(Me.Label9)
        Me.pnlInputs.Location = New System.Drawing.Point(3, 251)
        Me.pnlInputs.Name = "pnlInputs"
        Me.pnlInputs.Size = New System.Drawing.Size(584, 109)
        Me.pnlInputs.TabIndex = 126
        '
        'rdbBottomQurtile
        '
        Me.rdbBottomQurtile.AutoSize = True
        Me.rdbBottomQurtile.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbBottomQurtile.Location = New System.Drawing.Point(421, 9)
        Me.rdbBottomQurtile.Name = "rdbBottomQurtile"
        Me.rdbBottomQurtile.Size = New System.Drawing.Size(129, 20)
        Me.rdbBottomQurtile.TabIndex = 125
        Me.rdbBottomQurtile.Text = "Bottom Quartile"
        Me.rdbBottomQurtile.UseVisualStyleBackColor = True
        '
        'rdbDepreciation
        '
        Me.rdbDepreciation.AutoSize = True
        Me.rdbDepreciation.Checked = True
        Me.rdbDepreciation.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbDepreciation.Location = New System.Drawing.Point(281, 9)
        Me.rdbDepreciation.Name = "rdbDepreciation"
        Me.rdbDepreciation.Size = New System.Drawing.Size(108, 20)
        Me.rdbDepreciation.TabIndex = 124
        Me.rdbDepreciation.TabStop = True
        Me.rdbDepreciation.Text = "Depreciation"
        Me.rdbDepreciation.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(5, 11)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(239, 16)
        Me.Label6.TabIndex = 123
        Me.Label6.Text = "Redevelopment Candidate Method:"
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(5, 39)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(582, 4)
        Me.Panel3.TabIndex = 122
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator1, Me.btnApplyLegend, Me.ToolStripSeparator9, Me.barProgress})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 369)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(599, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 106
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
        Me.btnRun.Size = New System.Drawing.Size(65, 28)
        Me.btnRun.Text = "RUN"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 31)
        '
        'btnApplyLegend
        '
        Me.btnApplyLegend.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnApplyLegend.Image = CType(resources.GetObject("btnApplyLegend.Image"), System.Drawing.Image)
        Me.btnApplyLegend.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnApplyLegend.Name = "btnApplyLegend"
        Me.btnApplyLegend.Size = New System.Drawing.Size(76, 28)
        Me.btnApplyLegend.Text = "Apply Legend"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 31)
        '
        'barProgress
        '
        Me.barProgress.Name = "barProgress"
        Me.barProgress.Size = New System.Drawing.Size(250, 28)
        Me.barProgress.Step = 1
        Me.barProgress.Visible = False
        '
        'lblSqFt
        '
        Me.lblSqFt.AutoSize = True
        Me.lblSqFt.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSqFt.Location = New System.Drawing.Point(107, 195)
        Me.lblSqFt.Name = "lblSqFt"
        Me.lblSqFt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSqFt.Size = New System.Drawing.Size(202, 16)
        Me.lblSqFt.TabIndex = 127
        Me.lblSqFt.Text = "Select ""Square Footage"" field"
        '
        'cmbSqFt
        '
        Me.cmbSqFt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSqFt.FormattingEnabled = True
        Me.cmbSqFt.Location = New System.Drawing.Point(326, 194)
        Me.cmbSqFt.Name = "cmbSqFt"
        Me.cmbSqFt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSqFt.Size = New System.Drawing.Size(256, 21)
        Me.cmbSqFt.TabIndex = 128
        '
        'LblBreaks
        '
        Me.LblBreaks.AutoSize = True
        Me.LblBreaks.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblBreaks.Location = New System.Drawing.Point(107, 224)
        Me.LblBreaks.Name = "LblBreaks"
        Me.LblBreaks.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblBreaks.Size = New System.Drawing.Size(202, 16)
        Me.LblBreaks.TabIndex = 130
        Me.LblBreaks.Text = "Enter Number of Class Breaks"
        Me.LblBreaks.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxBreaks
        '
        Me.tbxBreaks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxBreaks.Location = New System.Drawing.Point(326, 223)
        Me.tbxBreaks.Name = "tbxBreaks"
        Me.tbxBreaks.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxBreaks.Size = New System.Drawing.Size(57, 20)
        Me.tbxBreaks.TabIndex = 129
        Me.tbxBreaks.Text = "4"
        '
        'frmRedevTiming
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(599, 400)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmRedevTiming"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Redevelopment Candidate"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnlInputs.ResumeLayout(False)
        Me.pnlInputs.PerformLayout()
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbParcelLayers As System.Windows.Forms.ComboBox
    Friend WithEvents lblParcelLyr As System.Windows.Forms.Label
    Friend WithEvents cmbYearBuilt As System.Windows.Forms.ComboBox
    Friend WithEvents lblYearBuilt As System.Windows.Forms.Label
    Friend WithEvents cmbImproveValue As System.Windows.Forms.ComboBox
    Friend WithEvents cmbLandValue As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblGridCellSize As System.Windows.Forms.Label
    Friend WithEvents tbxCurrentYear As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbxLifespan As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbxAppreciation As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents barProgress As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnApplyLegend As System.Windows.Forms.ToolStripButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents rdbBottomQurtile As System.Windows.Forms.RadioButton
    Friend WithEvents rdbDepreciation As System.Windows.Forms.RadioButton
    Friend WithEvents pnlInputs As System.Windows.Forms.Panel
    Friend WithEvents lblSqFt As System.Windows.Forms.Label
    Friend WithEvents cmbSqFt As System.Windows.Forms.ComboBox
    Friend WithEvents LblBreaks As System.Windows.Forms.Label
    Friend WithEvents tbxBreaks As System.Windows.Forms.TextBox
End Class
