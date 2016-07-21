<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFieldVariables
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
    Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Me.btnUncheckAll = New System.Windows.Forms.Button()
    Me.btnCheckAll = New System.Windows.Forms.Button()
    Me.btnApply = New System.Windows.Forms.Button()
    Me.lbl1toManyFields = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.TextBox1 = New System.Windows.Forms.TextBox()
    Me.dgvFields = New System.Windows.Forms.DataGridView()
    Me.Column3 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
    Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.dataGridViewColumnTotalOutputFldName = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.dataGridViewColumnCalcTot = New System.Windows.Forms.DataGridViewCheckBoxColumn()
    Me.Column6 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
    Me.chkLandUse = New System.Windows.Forms.CheckBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.prgStatus = New System.Windows.Forms.ProgressBar()
    Me.chkTravelBehavior = New System.Windows.Forms.CheckBox()
    Me.chkPublicHealth = New System.Windows.Forms.CheckBox()
    CType(Me.dgvFields, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'btnUncheckAll
    '
    Me.btnUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnUncheckAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnUncheckAll.Location = New System.Drawing.Point(673, 33)
    Me.btnUncheckAll.Name = "btnUncheckAll"
    Me.btnUncheckAll.Size = New System.Drawing.Size(129, 24)
    Me.btnUncheckAll.TabIndex = 118
    Me.btnUncheckAll.Text = "Uncheck All"
    Me.btnUncheckAll.UseVisualStyleBackColor = True
    '
    'btnCheckAll
    '
    Me.btnCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnCheckAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnCheckAll.Location = New System.Drawing.Point(534, 33)
    Me.btnCheckAll.Name = "btnCheckAll"
    Me.btnCheckAll.Size = New System.Drawing.Size(129, 24)
    Me.btnCheckAll.TabIndex = 117
    Me.btnCheckAll.Text = "Check All"
    Me.btnCheckAll.UseVisualStyleBackColor = True
    '
    'btnApply
    '
    Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnApply.BackColor = System.Drawing.Color.Khaki
    Me.btnApply.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnApply.Location = New System.Drawing.Point(674, 403)
    Me.btnApply.Name = "btnApply"
    Me.btnApply.Size = New System.Drawing.Size(129, 35)
    Me.btnApply.TabIndex = 119
    Me.btnApply.Text = "Apply"
    Me.btnApply.UseVisualStyleBackColor = False
    '
    'lbl1toManyFields
    '
    Me.lbl1toManyFields.AutoSize = True
    Me.lbl1toManyFields.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lbl1toManyFields.Location = New System.Drawing.Point(10, 37)
    Me.lbl1toManyFields.Name = "lbl1toManyFields"
    Me.lbl1toManyFields.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lbl1toManyFields.Size = New System.Drawing.Size(132, 16)
    Me.lbl1toManyFields.TabIndex = 120
    Me.lbl1toManyFields.Text = "Envision Attributes"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label2.Location = New System.Drawing.Point(10, 10)
    Me.Label2.Name = "Label2"
    Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label2.Size = New System.Drawing.Size(527, 16)
    Me.Label2.TabIndex = 121
    Me.Label2.Text = "Check all Envision Attribute fields you would like to have Envision track."
    '
    'TextBox1
    '
    Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
    Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
    Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.TextBox1.Location = New System.Drawing.Point(13, 412)
    Me.TextBox1.Multiline = True
    Me.TextBox1.Name = "TextBox1"
    Me.TextBox1.Size = New System.Drawing.Size(789, 29)
    Me.TextBox1.TabIndex = 122
    Me.TextBox1.Text = "The following fields are required Envision fields: DevType, Vacant_Acres, Devd_Ac" & _
    "es, and Constrained_Acre."
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
    DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    Me.dgvFields.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
    Me.dgvFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvFields.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column3, Me.Column1, Me.Column2, Me.Column4, Me.dataGridViewColumnTotalOutputFldName, Me.dataGridViewColumnCalcTot, Me.Column6})
    DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
    DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
    DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    Me.dgvFields.DefaultCellStyle = DataGridViewCellStyle2
    Me.dgvFields.Location = New System.Drawing.Point(13, 61)
    Me.dgvFields.MultiSelect = False
    Me.dgvFields.Name = "dgvFields"
    DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
    DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    Me.dgvFields.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
    Me.dgvFields.Size = New System.Drawing.Size(789, 234)
    Me.dgvFields.TabIndex = 123
    '
    'Column3
    '
    Me.Column3.FalseValue = "0"
    Me.Column3.FillWeight = 50.0!
    Me.Column3.HeaderText = "Track"
    Me.Column3.Name = "Column3"
    Me.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
    Me.Column3.TrueValue = "1"
    Me.Column3.Width = 40
    '
    'Column1
    '
    Me.Column1.FillWeight = 50.0!
    Me.Column1.HeaderText = "Attribute Field Name"
    Me.Column1.Name = "Column1"
    Me.Column1.ReadOnly = True
    Me.Column1.Width = 110
    '
    'Column2
    '
    Me.Column2.HeaderText = "Alias"
    Me.Column2.Name = "Column2"
    Me.Column2.ReadOnly = True
    Me.Column2.Width = 250
    '
    'Column4
    '
    Me.Column4.HeaderText = "New Output Field Name"
    Me.Column4.Name = "Column4"
    Me.Column4.ReadOnly = True
    Me.Column4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    '
    'dataGridViewColumnTotalOutputFldName
    '
    Me.dataGridViewColumnTotalOutputFldName.HeaderText = "Total Output Field Name"
    Me.dataGridViewColumnTotalOutputFldName.Name = "dataGridViewColumnTotalOutputFldName"
    Me.dataGridViewColumnTotalOutputFldName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    Me.dataGridViewColumnTotalOutputFldName.Width = 120
    '
    'dataGridViewColumnCalcTot
    '
    Me.dataGridViewColumnCalcTot.FillWeight = 50.0!
    Me.dataGridViewColumnCalcTot.HeaderText = "Calculate Total Output"
    Me.dataGridViewColumnCalcTot.Name = "dataGridViewColumnCalcTot"
    Me.dataGridViewColumnCalcTot.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
    Me.dataGridViewColumnCalcTot.Width = 55
    '
    'Column6
    '
    Me.Column6.FalseValue = "0"
    Me.Column6.FillWeight = 50.0!
    Me.Column6.HeaderText = "Write Outputs Only"
    Me.Column6.Name = "Column6"
    Me.Column6.TrueValue = "1"
    Me.Column6.Width = 50
    '
    'chkLandUse
    '
    Me.chkLandUse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.chkLandUse.AutoSize = True
    Me.chkLandUse.Location = New System.Drawing.Point(15, 324)
    Me.chkLandUse.Name = "chkLandUse"
    Me.chkLandUse.Size = New System.Drawing.Size(397, 17)
    Me.chkLandUse.TabIndex = 124
    Me.chkLandUse.Text = "Land Use --- (HU, EMP, SF, TH, MF, MH, RET, OFF, IND, PUB, EDU,HOTEL)"
    Me.chkLandUse.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(10, 299)
    Me.Label1.Name = "Label1"
    Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label1.Size = New System.Drawing.Size(241, 16)
    Me.Label1.TabIndex = 125
    Me.Label1.Text = "Quick Reference Field Groupings"
    '
    'prgStatus
    '
    Me.prgStatus.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.prgStatus.Location = New System.Drawing.Point(0, 444)
    Me.prgStatus.Name = "prgStatus"
    Me.prgStatus.Size = New System.Drawing.Size(814, 30)
    Me.prgStatus.TabIndex = 126
    '
    'chkTravelBehavior
    '
    Me.chkTravelBehavior.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.chkTravelBehavior.AutoSize = True
    Me.chkTravelBehavior.Location = New System.Drawing.Point(15, 353)
    Me.chkTravelBehavior.Name = "chkTravelBehavior"
    Me.chkTravelBehavior.Size = New System.Drawing.Size(507, 17)
    Me.chkTravelBehavior.TabIndex = 127
    Me.chkTravelBehavior.Text = "Travel Behavior --- (HH, HU, EMP, POP, AVG_HH_SIZE, AVG_HH_INCOME, AVG_HH_WORKERS" & _
    ")"
    Me.chkTravelBehavior.UseVisualStyleBackColor = True
    '
    'chkPublicHealth
    '
    Me.chkPublicHealth.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.chkPublicHealth.Location = New System.Drawing.Point(15, 374)
    Me.chkPublicHealth.Name = "chkPublicHealth"
    Me.chkPublicHealth.Size = New System.Drawing.Size(653, 32)
    Me.chkPublicHealth.TabIndex = 128
    Me.chkPublicHealth.Text = "Public Health --- (POP, HH, EMP, WORKERS, ppl_acre, NetEMPDen, RET, EMP_MIX, Int_" & _
    "Den_Mi, AVG_HH_SIZE, Pct_Med_HH_INC, Pct_Low_HH_INC, Pct_Hi_HH_INC)"
    Me.chkPublicHealth.UseVisualStyleBackColor = True
    '
    'frmFieldVariables
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(814, 474)
    Me.Controls.Add(Me.chkPublicHealth)
    Me.Controls.Add(Me.chkTravelBehavior)
    Me.Controls.Add(Me.prgStatus)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.chkLandUse)
    Me.Controls.Add(Me.dgvFields)
    Me.Controls.Add(Me.btnApply)
    Me.Controls.Add(Me.TextBox1)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.lbl1toManyFields)
    Me.Controls.Add(Me.btnUncheckAll)
    Me.Controls.Add(Me.btnCheckAll)
    Me.Name = "frmFieldVariables"
    Me.ShowIcon = False
    Me.ShowInTaskbar = False
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
    Me.Text = "Envision Parcel Layer Field Managment"
    CType(Me.dgvFields, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
    Friend WithEvents btnUncheckAll As System.Windows.Forms.Button
    Friend WithEvents btnCheckAll As System.Windows.Forms.Button
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents lbl1toManyFields As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents dgvFields As System.Windows.Forms.DataGridView
    Friend WithEvents chkLandUse As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents prgStatus As System.Windows.Forms.ProgressBar
  Friend WithEvents Column3 As System.Windows.Forms.DataGridViewCheckBoxColumn
  Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents dataGridViewColumnTotalOutputFldName As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents dataGridViewColumnCalcTot As System.Windows.Forms.DataGridViewCheckBoxColumn
  Friend WithEvents Column6 As System.Windows.Forms.DataGridViewCheckBoxColumn
  Friend WithEvents chkTravelBehavior As System.Windows.Forms.CheckBox
  Friend WithEvents chkPublicHealth As System.Windows.Forms.CheckBox
End Class
