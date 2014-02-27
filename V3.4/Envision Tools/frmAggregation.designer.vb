<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAggregation
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAggregation))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.lblNeighborhoodLyr = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.barStatus = New System.Windows.Forms.ToolStripProgressBar()
        Me.lblCount = New System.Windows.Forms.ToolStripLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkParcelSelected = New System.Windows.Forms.CheckBox()
        Me.rdbNeighborhood = New System.Windows.Forms.RadioButton()
        Me.rdbParcel = New System.Windows.Forms.RadioButton()
        Me.chkSelected = New System.Windows.Forms.CheckBox()
        Me.dgvFields = New System.Windows.Forms.DataGridView()
        Me.Column4 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lbl1toManyFields = New System.Windows.Forms.Label()
        Me.btnUncheckAllWeighted = New System.Windows.Forms.Button()
        Me.btnCheckAllWeighted = New System.Windows.Forms.Button()
        Me.btnUncheckAllSum = New System.Windows.Forms.Button()
        Me.btnCheckAllSum = New System.Windows.Forms.Button()
        Me.btnTrackNone = New System.Windows.Forms.Button()
        Me.btnTrackAll = New System.Windows.Forms.Button()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dgvFields, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(8, 94)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.Size = New System.Drawing.Size(672, 21)
        Me.cmbLayers.TabIndex = 140
        '
        'lblNeighborhoodLyr
        '
        Me.lblNeighborhoodLyr.AutoSize = True
        Me.lblNeighborhoodLyr.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNeighborhoodLyr.Location = New System.Drawing.Point(8, 73)
        Me.lblNeighborhoodLyr.Name = "lblNeighborhoodLyr"
        Me.lblNeighborhoodLyr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNeighborhoodLyr.Size = New System.Drawing.Size(204, 16)
        Me.lblNeighborhoodLyr.TabIndex = 141
        Me.lblNeighborhoodLyr.Text = "Select a Neighborhood Layer:"
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.barStatus, Me.lblCount})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 549)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(696, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 147
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
        'barStatus
        '
        Me.barStatus.Name = "barStatus"
        Me.barStatus.Size = New System.Drawing.Size(250, 28)
        Me.barStatus.Visible = False
        '
        'lblCount
        '
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(40, 28)
        Me.lblCount.Text = " 1 of #"
        Me.lblCount.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(2, 544)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(692, 5)
        Me.Panel1.TabIndex = 143
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.chkParcelSelected)
        Me.Panel2.Controls.Add(Me.rdbNeighborhood)
        Me.Panel2.Controls.Add(Me.rdbParcel)
        Me.Panel2.Controls.Add(Me.chkSelected)
        Me.Panel2.Controls.Add(Me.lblNeighborhoodLyr)
        Me.Panel2.Controls.Add(Me.cmbLayers)
        Me.Panel2.Location = New System.Drawing.Point(1, 1)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(693, 124)
        Me.Panel2.TabIndex = 150
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(375, 16)
        Me.Label1.TabIndex = 155
        Me.Label1.Text = "Select Parcel or Neighborhood Layer option to process:"
        '
        'chkParcelSelected
        '
        Me.chkParcelSelected.AutoSize = True
        Me.chkParcelSelected.Enabled = False
        Me.chkParcelSelected.Location = New System.Drawing.Point(179, 26)
        Me.chkParcelSelected.Name = "chkParcelSelected"
        Me.chkParcelSelected.Size = New System.Drawing.Size(193, 17)
        Me.chkParcelSelected.TabIndex = 154
        Me.chkParcelSelected.Text = "Use Selected Feature(s) Only"
        Me.chkParcelSelected.UseVisualStyleBackColor = True
        Me.chkParcelSelected.Visible = False
        '
        'rdbNeighborhood
        '
        Me.rdbNeighborhood.AutoSize = True
        Me.rdbNeighborhood.Checked = True
        Me.rdbNeighborhood.Location = New System.Drawing.Point(8, 51)
        Me.rdbNeighborhood.Name = "rdbNeighborhood"
        Me.rdbNeighborhood.Size = New System.Drawing.Size(140, 17)
        Me.rdbNeighborhood.TabIndex = 153
        Me.rdbNeighborhood.TabStop = True
        Me.rdbNeighborhood.Text = "Neighborhood Layer"
        Me.rdbNeighborhood.UseVisualStyleBackColor = True
        '
        'rdbParcel
        '
        Me.rdbParcel.AutoSize = True
        Me.rdbParcel.Location = New System.Drawing.Point(8, 25)
        Me.rdbParcel.Name = "rdbParcel"
        Me.rdbParcel.Size = New System.Drawing.Size(96, 17)
        Me.rdbParcel.TabIndex = 152
        Me.rdbParcel.Text = "Parcel Layer"
        Me.rdbParcel.UseVisualStyleBackColor = True
        '
        'chkSelected
        '
        Me.chkSelected.AutoSize = True
        Me.chkSelected.Enabled = False
        Me.chkSelected.Location = New System.Drawing.Point(179, 52)
        Me.chkSelected.Name = "chkSelected"
        Me.chkSelected.Size = New System.Drawing.Size(193, 17)
        Me.chkSelected.TabIndex = 151
        Me.chkSelected.Text = "Use Selected Feature(s) Only"
        Me.chkSelected.UseVisualStyleBackColor = True
        Me.chkSelected.Visible = False
        '
        'dgvFields
        '
        Me.dgvFields.AllowUserToAddRows = False
        Me.dgvFields.AllowUserToDeleteRows = False
        Me.dgvFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFields.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFields.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column4, Me.Column1, Me.Column2, Me.Column3, Me.Column5})
        Me.dgvFields.Location = New System.Drawing.Point(0, 183)
        Me.dgvFields.Name = "dgvFields"
        Me.dgvFields.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvFields.Size = New System.Drawing.Size(694, 289)
        Me.dgvFields.TabIndex = 157
        '
        'Column4
        '
        Me.Column4.FalseValue = "0"
        Me.Column4.HeaderText = "Track"
        Me.Column4.IndeterminateValue = "0"
        Me.Column4.Name = "Column4"
        Me.Column4.TrueValue = "1"
        '
        'Column1
        '
        Me.Column1.FillWeight = 50.0!
        Me.Column1.HeaderText = "Field Name"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Column2
        '
        Me.Column2.HeaderText = "Alias"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column2.Width = 250
        '
        'Column3
        '
        Me.Column3.FalseValue = "0"
        Me.Column3.HeaderText = "Calc Total Area Weighted Average"
        Me.Column3.Name = "Column3"
        Me.Column3.TrueValue = "1"
        '
        'Column5
        '
        Me.Column5.FalseValue = "0"
        Me.Column5.HeaderText = "Calc Total Sum"
        Me.Column5.Name = "Column5"
        Me.Column5.TrueValue = "1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(509, 16)
        Me.Label2.TabIndex = 155
        Me.Label2.Text = "Check all Envision Attribute fields you would like to have aggregated."
        '
        'lbl1toManyFields
        '
        Me.lbl1toManyFields.AutoSize = True
        Me.lbl1toManyFields.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1toManyFields.Location = New System.Drawing.Point(5, 158)
        Me.lbl1toManyFields.Name = "lbl1toManyFields"
        Me.lbl1toManyFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl1toManyFields.Size = New System.Drawing.Size(132, 16)
        Me.lbl1toManyFields.TabIndex = 154
        Me.lbl1toManyFields.Text = "Envision Attributes"
        '
        'btnUncheckAllWeighted
        '
        Me.btnUncheckAllWeighted.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUncheckAllWeighted.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUncheckAllWeighted.Location = New System.Drawing.Point(235, 478)
        Me.btnUncheckAllWeighted.Name = "btnUncheckAllWeighted"
        Me.btnUncheckAllWeighted.Size = New System.Drawing.Size(227, 24)
        Me.btnUncheckAllWeighted.TabIndex = 152
        Me.btnUncheckAllWeighted.Text = "Uncheck All Area Weighted Totals"
        Me.btnUncheckAllWeighted.UseVisualStyleBackColor = True
        '
        'btnCheckAllWeighted
        '
        Me.btnCheckAllWeighted.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCheckAllWeighted.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckAllWeighted.Location = New System.Drawing.Point(0, 478)
        Me.btnCheckAllWeighted.Name = "btnCheckAllWeighted"
        Me.btnCheckAllWeighted.Size = New System.Drawing.Size(229, 24)
        Me.btnCheckAllWeighted.TabIndex = 151
        Me.btnCheckAllWeighted.Text = "Check All Area Weighted Totals"
        Me.btnCheckAllWeighted.UseVisualStyleBackColor = True
        '
        'btnUncheckAllSum
        '
        Me.btnUncheckAllSum.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUncheckAllSum.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUncheckAllSum.Location = New System.Drawing.Point(235, 508)
        Me.btnUncheckAllSum.Name = "btnUncheckAllSum"
        Me.btnUncheckAllSum.Size = New System.Drawing.Size(227, 24)
        Me.btnUncheckAllSum.TabIndex = 159
        Me.btnUncheckAllSum.Text = "Uncheck All Sum Totals"
        Me.btnUncheckAllSum.UseVisualStyleBackColor = True
        '
        'btnCheckAllSum
        '
        Me.btnCheckAllSum.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCheckAllSum.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckAllSum.Location = New System.Drawing.Point(0, 508)
        Me.btnCheckAllSum.Name = "btnCheckAllSum"
        Me.btnCheckAllSum.Size = New System.Drawing.Size(229, 24)
        Me.btnCheckAllSum.TabIndex = 158
        Me.btnCheckAllSum.Text = "Check All Sum Totals"
        Me.btnCheckAllSum.UseVisualStyleBackColor = True
        '
        'btnTrackNone
        '
        Me.btnTrackNone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTrackNone.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTrackNone.Location = New System.Drawing.Point(467, 154)
        Me.btnTrackNone.Name = "btnTrackNone"
        Me.btnTrackNone.Size = New System.Drawing.Size(227, 24)
        Me.btnTrackNone.TabIndex = 161
        Me.btnTrackNone.Text = "Uncheck All for Tracking"
        Me.btnTrackNone.UseVisualStyleBackColor = True
        '
        'btnTrackAll
        '
        Me.btnTrackAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTrackAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTrackAll.Location = New System.Drawing.Point(232, 154)
        Me.btnTrackAll.Name = "btnTrackAll"
        Me.btnTrackAll.Size = New System.Drawing.Size(229, 24)
        Me.btnTrackAll.TabIndex = 160
        Me.btnTrackAll.Text = "Check All for Tracking"
        Me.btnTrackAll.UseVisualStyleBackColor = True
        '
        'frmAggregation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(696, 580)
        Me.Controls.Add(Me.btnTrackNone)
        Me.Controls.Add(Me.btnTrackAll)
        Me.Controls.Add(Me.btnUncheckAllSum)
        Me.Controls.Add(Me.btnCheckAllSum)
        Me.Controls.Add(Me.dgvFields)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbl1toManyFields)
        Me.Controls.Add(Me.btnUncheckAllWeighted)
        Me.Controls.Add(Me.btnCheckAllWeighted)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmAggregation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Neighborhood Aggregation"
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.dgvFields, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents lblNeighborhoodLyr As System.Windows.Forms.Label
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents chkSelected As System.Windows.Forms.CheckBox
    Friend WithEvents barStatus As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents dgvFields As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lbl1toManyFields As System.Windows.Forms.Label
    Friend WithEvents btnUncheckAllWeighted As System.Windows.Forms.Button
    Friend WithEvents btnCheckAllWeighted As System.Windows.Forms.Button
    Friend WithEvents btnUncheckAllSum As System.Windows.Forms.Button
    Friend WithEvents btnCheckAllSum As System.Windows.Forms.Button
    Friend WithEvents btnTrackNone As System.Windows.Forms.Button
    Friend WithEvents btnTrackAll As System.Windows.Forms.Button
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents rdbNeighborhood As System.Windows.Forms.RadioButton
    Friend WithEvents rdbParcel As System.Windows.Forms.RadioButton
    Friend WithEvents chkParcelSelected As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblCount As System.Windows.Forms.ToolStripLabel
End Class
