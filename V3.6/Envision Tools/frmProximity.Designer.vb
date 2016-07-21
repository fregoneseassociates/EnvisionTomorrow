<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProximity
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProximity))
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.barStatus = New System.Windows.Forms.ToolStripProgressBar()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.InfoMenu_Constraints = New System.Windows.Forms.ToolStrip()
        Me.btnCheckAll = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnUncheckAll = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnSumExisting = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.dgvProximity = New System.Windows.Forms.DataGridView()
        Me.Column4 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.itmSelectOption = New System.Windows.Forms.ToolStripDropDownButton()
        Me.UseCentroidWithinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UseIntersectsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.InfoMenu_Constraints.SuspendLayout()
        CType(Me.dgvProximity, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus, Me.barStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 423)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(664, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 59
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
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.InfoMenu_Constraints)
        Me.Panel1.Controls.Add(Me.dgvProximity)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(663, 419)
        Me.Panel1.TabIndex = 60
        '
        'InfoMenu_Constraints
        '
        Me.InfoMenu_Constraints.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.InfoMenu_Constraints.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnCheckAll, Me.ToolStripSeparator2, Me.btnUncheckAll, Me.ToolStripSeparator1, Me.btnSumExisting, Me.ToolStripSeparator3, Me.itmSelectOption})
        Me.InfoMenu_Constraints.Location = New System.Drawing.Point(0, 0)
        Me.InfoMenu_Constraints.Name = "InfoMenu_Constraints"
        Me.InfoMenu_Constraints.Size = New System.Drawing.Size(659, 31)
        Me.InfoMenu_Constraints.TabIndex = 106
        Me.InfoMenu_Constraints.Text = "ToolStrip1"
        '
        'btnCheckAll
        '
        Me.btnCheckAll.Image = CType(resources.GetObject("btnCheckAll.Image"), System.Drawing.Image)
        Me.btnCheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCheckAll.Name = "btnCheckAll"
        Me.btnCheckAll.Size = New System.Drawing.Size(121, 28)
        Me.btnCheckAll.Text = "Check All Layers"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 31)
        '
        'btnUncheckAll
        '
        Me.btnUncheckAll.Image = CType(resources.GetObject("btnUncheckAll.Image"), System.Drawing.Image)
        Me.btnUncheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnUncheckAll.Name = "btnUncheckAll"
        Me.btnUncheckAll.Size = New System.Drawing.Size(134, 28)
        Me.btnUncheckAll.Text = "Uncheck All Layers"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 31)
        '
        'btnSumExisting
        '
        Me.btnSumExisting.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnSumExisting.Image = CType(resources.GetObject("btnSumExisting.Image"), System.Drawing.Image)
        Me.btnSumExisting.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSumExisting.Name = "btnSumExisting"
        Me.btnSumExisting.Size = New System.Drawing.Size(169, 28)
        Me.btnSumExisting.Text = "Sum Existing Conditions - YES"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 31)
        '
        'dgvProximity
        '
        Me.dgvProximity.AllowUserToAddRows = False
        Me.dgvProximity.AllowUserToDeleteRows = False
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.Color.Blue
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.Color.White
        Me.dgvProximity.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvProximity.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvProximity.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvProximity.BackgroundColor = System.Drawing.SystemColors.Control
        Me.dgvProximity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProximity.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvProximity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProximity.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column4, Me.Column5, Me.Column1, Me.Column2, Me.Column3})
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvProximity.DefaultCellStyle = DataGridViewCellStyle11
        Me.dgvProximity.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvProximity.Location = New System.Drawing.Point(0, 34)
        Me.dgvProximity.MultiSelect = False
        Me.dgvProximity.Name = "dgvProximity"
        Me.dgvProximity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dgvProximity.RowHeadersWidth = 20
        Me.dgvProximity.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle12.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.Color.Black
        Me.dgvProximity.RowsDefaultCellStyle = DataGridViewCellStyle12
        Me.dgvProximity.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProximity.Size = New System.Drawing.Size(661, 383)
        Me.dgvProximity.TabIndex = 105
        '
        'Column4
        '
        Me.Column4.FalseValue = "0"
        Me.Column4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column4.HeaderText = "RUN"
        Me.Column4.Name = "Column4"
        Me.Column4.TrueValue = "1"
        '
        'Column5
        '
        Me.Column5.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.Column5.HeaderText = "Amenity Layer"
        Me.Column5.Name = "Column5"
        '
        'Column1
        '
        Me.Column1.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.Column1.HeaderText = "Buffer Option"
        Me.Column1.Items.AddRange(New Object() {"Custom", "1/4 MILE", "1/2 MILE", "1 MILE"})
        Me.Column1.Name = "Column1"
        Me.Column1.Visible = False
        '
        'Column2
        '
        Me.Column2.HeaderText = "Buffer Distance"
        Me.Column2.Name = "Column2"
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Column3
        '
        Me.Column3.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.Column3.HeaderText = "Distance Measure"
        Me.Column3.Items.AddRange(New Object() {"MILES", "FEET", "METERS", "KILOMETERS", "YARDS"})
        Me.Column3.Name = "Column3"
        '
        'itmSelectOption
        '
        Me.itmSelectOption.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.itmSelectOption.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UseCentroidWithinToolStripMenuItem, Me.UseIntersectsToolStripMenuItem})
        Me.itmSelectOption.Image = CType(resources.GetObject("itmSelectOption.Image"), System.Drawing.Image)
        Me.itmSelectOption.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.itmSelectOption.Name = "itmSelectOption"
        Me.itmSelectOption.Size = New System.Drawing.Size(94, 28)
        Me.itmSelectOption.Text = "Intersect - ON"
        '
        'UseCentroidWithinToolStripMenuItem
        '
        Me.UseCentroidWithinToolStripMenuItem.Name = "UseCentroidWithinToolStripMenuItem"
        Me.UseCentroidWithinToolStripMenuItem.Size = New System.Drawing.Size(188, 22)
        Me.UseCentroidWithinToolStripMenuItem.Text = "Use - Centroid Within"
        '
        'UseIntersectsToolStripMenuItem
        '
        Me.UseIntersectsToolStripMenuItem.Name = "UseIntersectsToolStripMenuItem"
        Me.UseIntersectsToolStripMenuItem.Size = New System.Drawing.Size(188, 22)
        Me.UseIntersectsToolStripMenuItem.Text = "Use - Intersect"
        '
        'frmProximity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(664, 454)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Name = "frmProximity"
        Me.Text = "Proximity Summary Tool"
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.InfoMenu_Constraints.ResumeLayout(False)
        Me.InfoMenu_Constraints.PerformLayout()
        CType(Me.dgvProximity, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents barStatus As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents dgvProximity As System.Windows.Forms.DataGridView
    Friend WithEvents InfoMenu_Constraints As System.Windows.Forms.ToolStrip
    Friend WithEvents btnCheckAll As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnUncheckAll As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnSumExisting As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents itmSelectOption As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents UseCentroidWithinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UseIntersectsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
