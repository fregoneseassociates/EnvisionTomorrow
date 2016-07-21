<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBufferSum
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBufferSum))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.tabSteps = New System.Windows.Forms.TabControl
        Me.tabStep1 = New System.Windows.Forms.TabPage
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.cmbLayers = New System.Windows.Forms.ComboBox
        Me.rdbNeighborhoodLayer = New System.Windows.Forms.RadioButton
        Me.Label2 = New System.Windows.Forms.Label
        Me.rdbParcelLayer = New System.Windows.Forms.RadioButton
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.rdbSelected = New System.Windows.Forms.RadioButton
        Me.rdbPartial = New System.Windows.Forms.RadioButton
        Me.rdbFull = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.chkNumericalFlds = New System.Windows.Forms.CheckedListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ToolStrip_InfoTab1 = New System.Windows.Forms.ToolStrip
        Me.btnNext1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.tabStep2 = New System.Windows.Forms.TabPage
        Me.dgvBuffers = New System.Windows.Forms.DataGridView
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.lbl1toManyFields = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.ToolStrip_InfoTab2 = New System.Windows.Forms.ToolStrip
        Me.btnNext2 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.btnPrevious1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator
        Me.tabStep3 = New System.Windows.Forms.TabPage
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator
        Me.btnPrevious2 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripProgressBar2 = New System.Windows.Forms.ToolStripProgressBar
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel
        Me.dgvMatrix = New System.Windows.Forms.DataGridView
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewCheckBoxColumn2 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column2 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column4 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column6 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column7 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column8 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel8 = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator
        Me.btnRun = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator
        Me.btnPrevious3 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator
        Me.barStatusRun = New System.Windows.Forms.ToolStripProgressBar
        Me.lblRunStatus = New System.Windows.Forms.ToolStripLabel
        Me.tabSteps.SuspendLayout()
        Me.tabStep1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.ToolStrip_InfoTab1.SuspendLayout()
        Me.tabStep2.SuspendLayout()
        CType(Me.dgvBuffers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip_InfoTab2.SuspendLayout()
        Me.tabStep3.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        CType(Me.dgvMatrix, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel8.SuspendLayout()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabSteps
        '
        Me.tabSteps.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.tabSteps.Controls.Add(Me.tabStep1)
        Me.tabSteps.Controls.Add(Me.tabStep2)
        Me.tabSteps.Controls.Add(Me.tabStep3)
        Me.tabSteps.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabSteps.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabSteps.Location = New System.Drawing.Point(0, 0)
        Me.tabSteps.Multiline = True
        Me.tabSteps.Name = "tabSteps"
        Me.tabSteps.Padding = New System.Drawing.Point(5, 20)
        Me.tabSteps.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tabSteps.SelectedIndex = 0
        Me.tabSteps.Size = New System.Drawing.Size(1203, 609)
        Me.tabSteps.TabIndex = 133
        '
        'tabStep1
        '
        Me.tabStep1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabStep1.Controls.Add(Me.Panel4)
        Me.tabStep1.Controls.Add(Me.Panel6)
        Me.tabStep1.Controls.Add(Me.Panel2)
        Me.tabStep1.Controls.Add(Me.rdbSelected)
        Me.tabStep1.Controls.Add(Me.rdbPartial)
        Me.tabStep1.Controls.Add(Me.rdbFull)
        Me.tabStep1.Controls.Add(Me.Label3)
        Me.tabStep1.Controls.Add(Me.chkNumericalFlds)
        Me.tabStep1.Controls.Add(Me.Label1)
        Me.tabStep1.Controls.Add(Me.Panel3)
        Me.tabStep1.Controls.Add(Me.ToolStrip_InfoTab1)
        Me.tabStep1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabStep1.Location = New System.Drawing.Point(4, 62)
        Me.tabStep1.Name = "tabStep1"
        Me.tabStep1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep1.Size = New System.Drawing.Size(1195, 543)
        Me.tabStep1.TabIndex = 0
        Me.tabStep1.Text = "Step 1: Select Fields to Sum"
        Me.tabStep1.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Location = New System.Drawing.Point(3, 54)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1188, 4)
        Me.Panel4.TabIndex = 44
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.Controls.Add(Me.cmbLayers)
        Me.Panel6.Controls.Add(Me.rdbNeighborhoodLayer)
        Me.Panel6.Controls.Add(Me.Label2)
        Me.Panel6.Controls.Add(Me.rdbParcelLayer)
        Me.Panel6.Location = New System.Drawing.Point(0, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1185, 51)
        Me.Panel6.TabIndex = 140
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.Enabled = False
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(445, 16)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.Size = New System.Drawing.Size(735, 23)
        Me.cmbLayers.TabIndex = 139
        '
        'rdbNeighborhoodLayer
        '
        Me.rdbNeighborhoodLayer.AutoSize = True
        Me.rdbNeighborhoodLayer.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbNeighborhoodLayer.Location = New System.Drawing.Point(283, 18)
        Me.rdbNeighborhoodLayer.Name = "rdbNeighborhoodLayer"
        Me.rdbNeighborhoodLayer.Size = New System.Drawing.Size(156, 20)
        Me.rdbNeighborhoodLayer.TabIndex = 137
        Me.rdbNeighborhoodLayer.Text = "Neighborhood Layer"
        Me.rdbNeighborhoodLayer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rdbNeighborhoodLayer.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(125, 16)
        Me.Label2.TabIndex = 138
        Me.Label2.Text = "Processing Layer:"
        '
        'rdbParcelLayer
        '
        Me.rdbParcelLayer.AutoSize = True
        Me.rdbParcelLayer.Checked = True
        Me.rdbParcelLayer.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbParcelLayer.Location = New System.Drawing.Point(150, 18)
        Me.rdbParcelLayer.Name = "rdbParcelLayer"
        Me.rdbParcelLayer.Size = New System.Drawing.Size(107, 20)
        Me.rdbParcelLayer.TabIndex = 136
        Me.rdbParcelLayer.TabStop = True
        Me.rdbParcelLayer.Text = "Parcel Layer"
        Me.rdbParcelLayer.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Location = New System.Drawing.Point(3, 107)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1188, 4)
        Me.Panel2.TabIndex = 43
        '
        'rdbSelected
        '
        Me.rdbSelected.AutoSize = True
        Me.rdbSelected.Enabled = False
        Me.rdbSelected.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbSelected.Location = New System.Drawing.Point(656, 79)
        Me.rdbSelected.Name = "rdbSelected"
        Me.rdbSelected.Size = New System.Drawing.Size(147, 20)
        Me.rdbSelected.TabIndex = 136
        Me.rdbSelected.TabStop = True
        Me.rdbSelected.Text = "Selected Features"
        Me.rdbSelected.UseVisualStyleBackColor = True
        '
        'rdbPartial
        '
        Me.rdbPartial.AutoSize = True
        Me.rdbPartial.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbPartial.Location = New System.Drawing.Point(409, 79)
        Me.rdbPartial.Name = "rdbPartial"
        Me.rdbPartial.Size = New System.Drawing.Size(233, 20)
        Me.rdbPartial.TabIndex = 135
        Me.rdbPartial.TabStop = True
        Me.rdbPartial.Text = "Partial (Dev Type Applied Only)"
        Me.rdbPartial.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rdbPartial.UseVisualStyleBackColor = True
        '
        'rdbFull
        '
        Me.rdbFull.AutoSize = True
        Me.rdbFull.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbFull.Location = New System.Drawing.Point(263, 79)
        Me.rdbFull.Name = "rdbFull"
        Me.rdbFull.Size = New System.Drawing.Size(143, 20)
        Me.rdbFull.TabIndex = 134
        Me.rdbFull.TabStop = True
        Me.rdbFull.Text = "Full (All Features)"
        Me.rdbFull.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(254, 16)
        Me.Label3.TabIndex = 133
        Me.Label3.Text = "Processing Feature Selection Option:"
        '
        'chkNumericalFlds
        '
        Me.chkNumericalFlds.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkNumericalFlds.CheckOnClick = True
        Me.chkNumericalFlds.FormattingEnabled = True
        Me.chkNumericalFlds.Location = New System.Drawing.Point(6, 158)
        Me.chkNumericalFlds.Name = "chkNumericalFlds"
        Me.chkNumericalFlds.Size = New System.Drawing.Size(1179, 340)
        Me.chkNumericalFlds.TabIndex = 100
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(567, 16)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Select the Numerical Fields you would like to have summarized by a buffer distance:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(1, 498)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1188, 4)
        Me.Panel3.TabIndex = 42
        '
        'ToolStrip_InfoTab1
        '
        Me.ToolStrip_InfoTab1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnNext1, Me.ToolStripSeparator1, Me.ToolStripProgressBar1, Me.ToolStripLabel1})
        Me.ToolStrip_InfoTab1.Location = New System.Drawing.Point(3, 505)
        Me.ToolStrip_InfoTab1.Name = "ToolStrip_InfoTab1"
        Me.ToolStrip_InfoTab1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab1.Size = New System.Drawing.Size(1185, 31)
        Me.ToolStrip_InfoTab1.TabIndex = 41
        Me.ToolStrip_InfoTab1.Text = "ToolStrip3"
        '
        'btnNext1
        '
        Me.btnNext1.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext1.Image = CType(resources.GetObject("btnNext1.Image"), System.Drawing.Image)
        Me.btnNext1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnNext1.Name = "btnNext1"
        Me.btnNext1.Size = New System.Drawing.Size(70, 28)
        Me.btnNext1.Text = "Next"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 28)
        Me.ToolStripProgressBar1.Step = 1
        Me.ToolStripProgressBar1.Visible = False
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ToolStripLabel1.Size = New System.Drawing.Size(84, 28)
        Me.ToolStripLabel1.Text = "Processing:"
        Me.ToolStripLabel1.Visible = False
        '
        'tabStep2
        '
        Me.tabStep2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabStep2.Controls.Add(Me.dgvBuffers)
        Me.tabStep2.Controls.Add(Me.lbl1toManyFields)
        Me.tabStep2.Controls.Add(Me.Panel5)
        Me.tabStep2.Controls.Add(Me.ToolStrip_InfoTab2)
        Me.tabStep2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabStep2.Location = New System.Drawing.Point(4, 62)
        Me.tabStep2.Name = "tabStep2"
        Me.tabStep2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tabStep2.Size = New System.Drawing.Size(1195, 543)
        Me.tabStep2.TabIndex = 1
        Me.tabStep2.Text = "Step 2: Define Buffers"
        Me.tabStep2.UseVisualStyleBackColor = True
        '
        'dgvBuffers
        '
        Me.dgvBuffers.AllowUserToAddRows = False
        Me.dgvBuffers.AllowUserToDeleteRows = False
        Me.dgvBuffers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvBuffers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBuffers.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvBuffers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBuffers.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewCheckBoxColumn1, Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3})
        Me.dgvBuffers.Location = New System.Drawing.Point(6, 30)
        Me.dgvBuffers.MultiSelect = False
        Me.dgvBuffers.Name = "dgvBuffers"
        Me.dgvBuffers.Size = New System.Drawing.Size(1179, 459)
        Me.dgvBuffers.TabIndex = 127
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.FillWeight = 50.0!
        Me.DataGridViewCheckBoxColumn1.HeaderText = "Track"
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "BUFFER"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Buffer Distance (miles)"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Deafult Field Append Text (i.e _1MI)"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'lbl1toManyFields
        '
        Me.lbl1toManyFields.AutoSize = True
        Me.lbl1toManyFields.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1toManyFields.Location = New System.Drawing.Point(2, 5)
        Me.lbl1toManyFields.Name = "lbl1toManyFields"
        Me.lbl1toManyFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl1toManyFields.Size = New System.Drawing.Size(132, 16)
        Me.lbl1toManyFields.TabIndex = 126
        Me.lbl1toManyFields.Text = "Envision Attributes"
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Location = New System.Drawing.Point(-1, 498)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(1193, 4)
        Me.Panel5.TabIndex = 59
        '
        'ToolStrip_InfoTab2
        '
        Me.ToolStrip_InfoTab2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab2.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnNext2, Me.ToolStripSeparator3, Me.btnPrevious1, Me.ToolStripSeparator4})
        Me.ToolStrip_InfoTab2.Location = New System.Drawing.Point(3, 505)
        Me.ToolStrip_InfoTab2.Name = "ToolStrip_InfoTab2"
        Me.ToolStrip_InfoTab2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab2.Size = New System.Drawing.Size(1185, 31)
        Me.ToolStrip_InfoTab2.TabIndex = 36
        Me.ToolStrip_InfoTab2.Text = "ToolStrip2"
        '
        'btnNext2
        '
        Me.btnNext2.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext2.Image = CType(resources.GetObject("btnNext2.Image"), System.Drawing.Image)
        Me.btnNext2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnNext2.Name = "btnNext2"
        Me.btnNext2.Size = New System.Drawing.Size(70, 28)
        Me.btnNext2.Text = "Next"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 31)
        '
        'btnPrevious1
        '
        Me.btnPrevious1.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious1.Image = CType(resources.GetObject("btnPrevious1.Image"), System.Drawing.Image)
        Me.btnPrevious1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrevious1.Name = "btnPrevious1"
        Me.btnPrevious1.Size = New System.Drawing.Size(99, 28)
        Me.btnPrevious1.Text = "Previous"
        Me.btnPrevious1.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 31)
        '
        'tabStep3
        '
        Me.tabStep3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabStep3.Controls.Add(Me.ToolStrip2)
        Me.tabStep3.Controls.Add(Me.dgvMatrix)
        Me.tabStep3.Controls.Add(Me.Panel8)
        Me.tabStep3.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.tabStep3.Location = New System.Drawing.Point(4, 62)
        Me.tabStep3.Name = "tabStep3"
        Me.tabStep3.Padding = New System.Windows.Forms.Padding(3)
        Me.tabStep3.Size = New System.Drawing.Size(1195, 543)
        Me.tabStep3.TabIndex = 3
        Me.tabStep3.Text = "Step 3: Define Output Fields"
        Me.tabStep3.UseVisualStyleBackColor = True
        '
        'ToolStrip2
        '
        Me.ToolStrip2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip2.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator2, Me.ToolStripButton1, Me.ToolStripSeparator8, Me.btnPrevious2, Me.ToolStripSeparator5, Me.ToolStripProgressBar2, Me.ToolStripLabel2})
        Me.ToolStrip2.Location = New System.Drawing.Point(3, 481)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip2.Size = New System.Drawing.Size(806, 31)
        Me.ToolStrip2.TabIndex = 128
        Me.ToolStrip2.Text = "ToolStrip4"
        Me.ToolStrip2.Visible = False
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Enabled = False
        Me.ToolStripButton1.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(65, 28)
        Me.ToolStripButton1.Text = "RUN"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(6, 31)
        '
        'btnPrevious2
        '
        Me.btnPrevious2.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious2.Image = CType(resources.GetObject("btnPrevious2.Image"), System.Drawing.Image)
        Me.btnPrevious2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrevious2.Name = "btnPrevious2"
        Me.btnPrevious2.Size = New System.Drawing.Size(99, 28)
        Me.btnPrevious2.Text = "Previous"
        Me.btnPrevious2.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripProgressBar2
        '
        Me.ToolStripProgressBar2.Name = "ToolStripProgressBar2"
        Me.ToolStripProgressBar2.Size = New System.Drawing.Size(100, 28)
        Me.ToolStripProgressBar2.Step = 1
        Me.ToolStripProgressBar2.Visible = False
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ToolStripLabel2.Size = New System.Drawing.Size(84, 28)
        Me.ToolStripLabel2.Text = "Processing:"
        Me.ToolStripLabel2.Visible = False
        '
        'dgvMatrix
        '
        Me.dgvMatrix.AllowUserToAddRows = False
        Me.dgvMatrix.AllowUserToDeleteRows = False
        Me.dgvMatrix.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMatrix.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMatrix.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvMatrix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMatrix.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.DataGridViewCheckBoxColumn2, Me.DataGridViewTextBoxColumn4, Me.Column2, Me.DataGridViewTextBoxColumn5, Me.Column4, Me.DataGridViewTextBoxColumn6, Me.Column6, Me.Column3, Me.Column7, Me.Column5, Me.Column8, Me.Column9})
        Me.dgvMatrix.Location = New System.Drawing.Point(6, 6)
        Me.dgvMatrix.MultiSelect = False
        Me.dgvMatrix.Name = "dgvMatrix"
        Me.dgvMatrix.Size = New System.Drawing.Size(1179, 486)
        Me.dgvMatrix.TabIndex = 127
        '
        'Column1
        '
        Me.Column1.HeaderText = "Input Field"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 102
        '
        'DataGridViewCheckBoxColumn2
        '
        Me.DataGridViewCheckBoxColumn2.FillWeight = 50.0!
        Me.DataGridViewCheckBoxColumn2.HeaderText = "1/4 Mile"
        Me.DataGridViewCheckBoxColumn2.Name = "DataGridViewCheckBoxColumn2"
        Me.DataGridViewCheckBoxColumn2.Width = 67
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn4.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 112
        '
        'Column2
        '
        Me.Column2.FillWeight = 50.0!
        Me.Column2.HeaderText = "1/2 Mile"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 67
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn5.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 112
        '
        'Column4
        '
        Me.Column4.FillWeight = 50.0!
        Me.Column4.HeaderText = "1 Mile"
        Me.Column4.Name = "Column4"
        Me.Column4.Width = 51
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn6.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 112
        '
        'Column6
        '
        Me.Column6.FillWeight = 50.0!
        Me.Column6.HeaderText = "User Defined 1"
        Me.Column6.Name = "Column6"
        Me.Column6.Width = 96
        '
        'Column3
        '
        Me.Column3.FillWeight = 50.0!
        Me.Column3.HeaderText = "Output Field"
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 112
        '
        'Column7
        '
        Me.Column7.FillWeight = 50.0!
        Me.Column7.HeaderText = "User Defined 2"
        Me.Column7.Name = "Column7"
        Me.Column7.Width = 96
        '
        'Column5
        '
        Me.Column5.FillWeight = 50.0!
        Me.Column5.HeaderText = "Output Field"
        Me.Column5.Name = "Column5"
        Me.Column5.Width = 112
        '
        'Column8
        '
        Me.Column8.HeaderText = "User Defined 3"
        Me.Column8.Name = "Column8"
        Me.Column8.Width = 96
        '
        'Column9
        '
        Me.Column9.HeaderText = "Output Field"
        Me.Column9.Name = "Column9"
        Me.Column9.Width = 112
        '
        'Panel8
        '
        Me.Panel8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel8.Controls.Add(Me.Panel1)
        Me.Panel8.Location = New System.Drawing.Point(2, 498)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(1184, 4)
        Me.Panel8.TabIndex = 53
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(-2, -2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1184, 4)
        Me.Panel1.TabIndex = 54
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.btnPrevious3, Me.ToolStripSeparator10, Me.barStatusRun, Me.lblRunStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(3, 505)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(1185, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 52
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
        'btnPrevious3
        '
        Me.btnPrevious3.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious3.Image = CType(resources.GetObject("btnPrevious3.Image"), System.Drawing.Image)
        Me.btnPrevious3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrevious3.Name = "btnPrevious3"
        Me.btnPrevious3.Size = New System.Drawing.Size(99, 28)
        Me.btnPrevious3.Text = "Previous"
        Me.btnPrevious3.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(6, 31)
        '
        'barStatusRun
        '
        Me.barStatusRun.Name = "barStatusRun"
        Me.barStatusRun.Size = New System.Drawing.Size(100, 28)
        Me.barStatusRun.Step = 1
        Me.barStatusRun.Visible = False
        '
        'lblRunStatus
        '
        Me.lblRunStatus.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRunStatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblRunStatus.Name = "lblRunStatus"
        Me.lblRunStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRunStatus.Size = New System.Drawing.Size(84, 28)
        Me.lblRunStatus.Text = "Processing:"
        Me.lblRunStatus.Visible = False
        '
        'frmBufferSum
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1203, 609)
        Me.Controls.Add(Me.tabSteps)
        Me.Name = "frmBufferSum"
        Me.Text = "Buffer Summaries"
        Me.tabSteps.ResumeLayout(False)
        Me.tabStep1.ResumeLayout(False)
        Me.tabStep1.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ToolStrip_InfoTab1.ResumeLayout(False)
        Me.ToolStrip_InfoTab1.PerformLayout()
        Me.tabStep2.ResumeLayout(False)
        Me.tabStep2.PerformLayout()
        CType(Me.dgvBuffers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip_InfoTab2.ResumeLayout(False)
        Me.ToolStrip_InfoTab2.PerformLayout()
        Me.tabStep3.ResumeLayout(False)
        Me.tabStep3.PerformLayout()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        CType(Me.dgvMatrix, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel8.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabSteps As System.Windows.Forms.TabControl
    Friend WithEvents tabStep1 As System.Windows.Forms.TabPage
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip_InfoTab1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnNext1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tabStep2 As System.Windows.Forms.TabPage
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip_InfoTab2 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnNext2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnPrevious1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tabStep3 As System.Windows.Forms.TabPage
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnPrevious3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents barStatusRun As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents lblRunStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents chkNumericalFlds As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvBuffers As System.Windows.Forms.DataGridView
    Friend WithEvents lbl1toManyFields As System.Windows.Forms.Label
    Friend WithEvents dgvMatrix As System.Windows.Forms.DataGridView
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnPrevious2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripProgressBar2 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents DataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents rdbSelected As System.Windows.Forms.RadioButton
    Friend WithEvents rdbPartial As System.Windows.Forms.RadioButton
    Friend WithEvents rdbFull As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents rdbNeighborhoodLayer As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rdbParcelLayer As System.Windows.Forms.RadioButton
End Class
