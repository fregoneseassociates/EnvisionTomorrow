<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLocalJobsHousingBalance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLocalJobsHousingBalance))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnApplyLegend = New System.Windows.Forms.Button
        Me.rdbIncomeBalance = New System.Windows.Forms.RadioButton
        Me.rdbJobWorkerLegend = New System.Windows.Forms.RadioButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbxBufferDistance = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbAvgWorkersIncome = New System.Windows.Forms.ComboBox
        Me.lblParcelLyr = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmbParcelLayers = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbEmployedRes = New System.Windows.Forms.ComboBox
        Me.cmbEmployees = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbAvgResIncome = New System.Windows.Forms.ComboBox
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator
        Me.btnRun = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator
        Me.barStatusRun = New System.Windows.Forms.ToolStripProgressBar
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
        Me.Panel1.Controls.Add(Me.btnApplyLegend)
        Me.Panel1.Controls.Add(Me.rdbIncomeBalance)
        Me.Panel1.Controls.Add(Me.rdbJobWorkerLegend)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.tbxBufferDistance)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.cmbAvgWorkersIncome)
        Me.Panel1.Controls.Add(Me.lblParcelLyr)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cmbParcelLayers)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmbEmployedRes)
        Me.Panel1.Controls.Add(Me.cmbEmployees)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cmbAvgResIncome)
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(577, 259)
        Me.Panel1.TabIndex = 1
        '
        'btnApplyLegend
        '
        Me.btnApplyLegend.Location = New System.Drawing.Point(495, 225)
        Me.btnApplyLegend.Name = "btnApplyLegend"
        Me.btnApplyLegend.Size = New System.Drawing.Size(75, 23)
        Me.btnApplyLegend.TabIndex = 125
        Me.btnApplyLegend.Text = "Apply"
        Me.btnApplyLegend.UseVisualStyleBackColor = True
        '
        'rdbIncomeBalance
        '
        Me.rdbIncomeBalance.AutoSize = True
        Me.rdbIncomeBalance.Location = New System.Drawing.Point(342, 228)
        Me.rdbIncomeBalance.Name = "rdbIncomeBalance"
        Me.rdbIncomeBalance.Size = New System.Drawing.Size(102, 17)
        Me.rdbIncomeBalance.TabIndex = 124
        Me.rdbIncomeBalance.Text = "Income Balance"
        Me.rdbIncomeBalance.UseVisualStyleBackColor = True
        '
        'rdbJobWorkerLegend
        '
        Me.rdbJobWorkerLegend.AutoSize = True
        Me.rdbJobWorkerLegend.Checked = True
        Me.rdbJobWorkerLegend.Location = New System.Drawing.Point(201, 228)
        Me.rdbJobWorkerLegend.Name = "rdbJobWorkerLegend"
        Me.rdbJobWorkerLegend.Size = New System.Drawing.Size(135, 17)
        Me.rdbJobWorkerLegend.TabIndex = 123
        Me.rdbJobWorkerLegend.TabStop = True
        Me.rdbJobWorkerLegend.Text = "Job / Workers Balance"
        Me.rdbJobWorkerLegend.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(21, 228)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(152, 16)
        Me.Label5.TabIndex = 122
        Me.Label5.Text = "Legend Classification:"
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(1, 212)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(575, 4)
        Me.Panel3.TabIndex = 121
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(506, 43)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 117
        Me.Label7.Text = "Miles"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(142, 44)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(161, 16)
        Me.Label6.TabIndex = 116
        Me.Label6.Text = "Define Buffer Distance:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxBufferDistance
        '
        Me.tbxBufferDistance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxBufferDistance.Location = New System.Drawing.Point(309, 43)
        Me.tbxBufferDistance.Name = "tbxBufferDistance"
        Me.tbxBufferDistance.Size = New System.Drawing.Size(191, 20)
        Me.tbxBufferDistance.TabIndex = 115
        Me.tbxBufferDistance.Text = "3"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(31, 183)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(262, 16)
        Me.Label4.TabIndex = 113
        Me.Label4.Text = "Select ""Avg. Income of Workers"" field:"
        '
        'cmbAvgWorkersIncome
        '
        Me.cmbAvgWorkersIncome.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAvgWorkersIncome.FormattingEnabled = True
        Me.cmbAvgWorkersIncome.Location = New System.Drawing.Point(309, 182)
        Me.cmbAvgWorkersIncome.Name = "cmbAvgWorkersIncome"
        Me.cmbAvgWorkersIncome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAvgWorkersIncome.Size = New System.Drawing.Size(259, 21)
        Me.cmbAvgWorkersIncome.TabIndex = 114
        '
        'lblParcelLyr
        '
        Me.lblParcelLyr.AutoSize = True
        Me.lblParcelLyr.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblParcelLyr.Location = New System.Drawing.Point(5, 11)
        Me.lblParcelLyr.Name = "lblParcelLyr"
        Me.lblParcelLyr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblParcelLyr.Size = New System.Drawing.Size(202, 16)
        Me.lblParcelLyr.TabIndex = 105
        Me.lblParcelLyr.Text = "Select your polygon dataset:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(137, 113)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(156, 16)
        Me.Label2.TabIndex = 112
        Me.Label2.Text = "Select ""Workers"" field:"
        '
        'cmbParcelLayers
        '
        Me.cmbParcelLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbParcelLayers.FormattingEnabled = True
        Me.cmbParcelLayers.Location = New System.Drawing.Point(213, 9)
        Me.cmbParcelLayers.Name = "cmbParcelLayers"
        Me.cmbParcelLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbParcelLayers.Size = New System.Drawing.Size(355, 21)
        Me.cmbParcelLayers.TabIndex = 106
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(60, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(233, 16)
        Me.Label1.TabIndex = 107
        Me.Label1.Text = "Select ""Employed Residents"" field:"
        '
        'cmbEmployedRes
        '
        Me.cmbEmployedRes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbEmployedRes.FormattingEnabled = True
        Me.cmbEmployedRes.Location = New System.Drawing.Point(309, 77)
        Me.cmbEmployedRes.Name = "cmbEmployedRes"
        Me.cmbEmployedRes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbEmployedRes.Size = New System.Drawing.Size(259, 21)
        Me.cmbEmployedRes.TabIndex = 108
        '
        'cmbEmployees
        '
        Me.cmbEmployees.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbEmployees.FormattingEnabled = True
        Me.cmbEmployees.Location = New System.Drawing.Point(309, 112)
        Me.cmbEmployees.Name = "cmbEmployees"
        Me.cmbEmployees.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbEmployees.Size = New System.Drawing.Size(259, 21)
        Me.cmbEmployees.TabIndex = 109
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(21, 148)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(272, 16)
        Me.Label3.TabIndex = 110
        Me.Label3.Text = "Select ""Avg. Income of Residents"" field:"
        '
        'cmbAvgResIncome
        '
        Me.cmbAvgResIncome.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAvgResIncome.FormattingEnabled = True
        Me.cmbAvgResIncome.Location = New System.Drawing.Point(309, 147)
        Me.cmbAvgResIncome.Name = "cmbAvgResIncome"
        Me.cmbAvgResIncome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAvgResIncome.Size = New System.Drawing.Size(259, 21)
        Me.cmbAvgResIncome.TabIndex = 111
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.barStatusRun})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 263)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(578, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 53
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
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 31)
        '
        'barStatusRun
        '
        Me.barStatusRun.Name = "barStatusRun"
        Me.barStatusRun.Size = New System.Drawing.Size(250, 28)
        Me.barStatusRun.Step = 1
        Me.barStatusRun.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.barStatusRun.Visible = False
        '
        'frmLocalJobsHousingBalance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(578, 294)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmLocalJobsHousingBalance"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Local Jobs-Housing Balance App"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbAvgWorkersIncome As System.Windows.Forms.ComboBox
    Friend WithEvents lblParcelLyr As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbParcelLayers As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbEmployedRes As System.Windows.Forms.ComboBox
    Friend WithEvents cmbEmployees As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbAvgResIncome As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbxBufferDistance As System.Windows.Forms.TextBox
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents barStatusRun As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rdbJobWorkerLegend As System.Windows.Forms.RadioButton
    Friend WithEvents rdbIncomeBalance As System.Windows.Forms.RadioButton
    Friend WithEvents btnApplyLegend As System.Windows.Forms.Button
End Class
