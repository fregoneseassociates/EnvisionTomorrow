<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFeatureSum
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFeatureSum))
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgvPOINT = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn2 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblPoint = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.barStatusRun = New System.Windows.Forms.ToolStripProgressBar()
        Me.lblRunStatus = New System.Windows.Forms.ToolStripLabel()
        Me.dgvARC = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewComboBoxColumn1 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewCheckBoxColumn3 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn4 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblArc = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.rdbSelected = New System.Windows.Forms.RadioButton()
        Me.rdbPartial = New System.Windows.Forms.RadioButton()
        Me.rdbFull = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkAcres = New System.Windows.Forms.CheckBox()
        Me.lblAcres = New System.Windows.Forms.Label()
        Me.cmbAcresFld = New System.Windows.Forms.ComboBox()
        Me.cmbSqMi = New System.Windows.Forms.ComboBox()
        Me.lblSqMi = New System.Windows.Forms.Label()
        Me.chkSqMi = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tbxProjection = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnSelectPrj = New System.Windows.Forms.Button()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.rdbNeighborhoodLayer = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rdbParcelLayer = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        CType(Me.dgvPOINT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        CType(Me.dgvARC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvPOINT
        '
        Me.dgvPOINT.AllowUserToAddRows = False
        Me.dgvPOINT.AllowUserToDeleteRows = False
        Me.dgvPOINT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPOINT.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPOINT.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvPOINT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPOINT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn6, Me.DataGridViewCheckBoxColumn2, Me.DataGridViewTextBoxColumn4, Me.Column2, Me.DataGridViewTextBoxColumn5, Me.Column4, Me.Column1})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvPOINT.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvPOINT.Enabled = False
        Me.dgvPOINT.Location = New System.Drawing.Point(7, 237)
        Me.dgvPOINT.MultiSelect = False
        Me.dgvPOINT.Name = "dgvPOINT"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPOINT.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvPOINT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvPOINT.ShowCellErrors = False
        Me.dgvPOINT.ShowEditingIcon = False
        Me.dgvPOINT.Size = New System.Drawing.Size(847, 155)
        Me.dgvPOINT.TabIndex = 130
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn6.HeaderText = "Layer Name"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        Me.DataGridViewTextBoxColumn6.Width = 91
        '
        'DataGridViewCheckBoxColumn2
        '
        Me.DataGridViewCheckBoxColumn2.HeaderText = "Track Feature Count"
        Me.DataGridViewCheckBoxColumn2.Name = "DataGridViewCheckBoxColumn2"
        Me.DataGridViewCheckBoxColumn2.Width = 111
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn4.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 90
        '
        'Column2
        '
        Me.Column2.HeaderText = "Track Sq Mile Density"
        Me.Column2.Name = "Column2"
        Me.Column2.Visible = False
        Me.Column2.Width = 82
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn5.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Visible = False
        Me.DataGridViewTextBoxColumn5.Width = 90
        '
        'Column4
        '
        Me.Column4.HeaderText = "Track Acre Density"
        Me.Column4.Name = "Column4"
        Me.Column4.Visible = False
        Me.Column4.Width = 102
        '
        'Column1
        '
        Me.Column1.HeaderText = "Output Field"
        Me.Column1.Name = "Column1"
        Me.Column1.Visible = False
        Me.Column1.Width = 90
        '
        'lblPoint
        '
        Me.lblPoint.AutoSize = True
        Me.lblPoint.Enabled = False
        Me.lblPoint.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoint.Location = New System.Drawing.Point(8, 217)
        Me.lblPoint.Name = "lblPoint"
        Me.lblPoint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPoint.Size = New System.Drawing.Size(337, 16)
        Me.lblPoint.TabIndex = 129
        Me.lblPoint.Text = "POINT layer feature count and density summaries"
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.barStatusRun, Me.lblRunStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 693)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(861, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 128
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
        'dgvARC
        '
        Me.dgvARC.AllowUserToAddRows = False
        Me.dgvARC.AllowUserToDeleteRows = False
        Me.dgvARC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvARC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvARC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvARC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvARC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewCheckBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewComboBoxColumn1, Me.DataGridViewCheckBoxColumn3, Me.DataGridViewTextBoxColumn3, Me.DataGridViewCheckBoxColumn4, Me.DataGridViewTextBoxColumn7})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvARC.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvARC.Location = New System.Drawing.Point(7, 440)
        Me.dgvARC.MultiSelect = False
        Me.dgvARC.Name = "dgvARC"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvARC.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvARC.Size = New System.Drawing.Size(847, 180)
        Me.dgvARC.TabIndex = 132
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Layer Name"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 91
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.HeaderText = "Track Feature Length Sum"
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        Me.DataGridViewCheckBoxColumn1.Width = 119
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn2.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 90
        '
        'DataGridViewComboBoxColumn1
        '
        Me.DataGridViewComboBoxColumn1.HeaderText = "Output Units"
        Me.DataGridViewComboBoxColumn1.Items.AddRange(New Object() {"Feet", "Miles"})
        Me.DataGridViewComboBoxColumn1.Name = "DataGridViewComboBoxColumn1"
        Me.DataGridViewComboBoxColumn1.Width = 72
        '
        'DataGridViewCheckBoxColumn3
        '
        Me.DataGridViewCheckBoxColumn3.HeaderText = "Track Sq Mile Density"
        Me.DataGridViewCheckBoxColumn3.Name = "DataGridViewCheckBoxColumn3"
        Me.DataGridViewCheckBoxColumn3.Visible = False
        Me.DataGridViewCheckBoxColumn3.Width = 82
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn3.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Visible = False
        Me.DataGridViewTextBoxColumn3.Width = 90
        '
        'DataGridViewCheckBoxColumn4
        '
        Me.DataGridViewCheckBoxColumn4.HeaderText = "Track Acre Density"
        Me.DataGridViewCheckBoxColumn4.Name = "DataGridViewCheckBoxColumn4"
        Me.DataGridViewCheckBoxColumn4.Visible = False
        Me.DataGridViewCheckBoxColumn4.Width = 102
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Output Field"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.Visible = False
        Me.DataGridViewTextBoxColumn7.Width = 90
        '
        'lblArc
        '
        Me.lblArc.AutoSize = True
        Me.lblArc.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArc.Location = New System.Drawing.Point(8, 420)
        Me.lblArc.Name = "lblArc"
        Me.lblArc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblArc.Size = New System.Drawing.Size(325, 16)
        Me.lblArc.TabIndex = 131
        Me.lblArc.Text = "LINE layer feature count and density summaries"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Location = New System.Drawing.Point(5, 203)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(852, 4)
        Me.Panel2.TabIndex = 137
        '
        'rdbSelected
        '
        Me.rdbSelected.AutoSize = True
        Me.rdbSelected.Enabled = False
        Me.rdbSelected.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbSelected.Location = New System.Drawing.Point(688, 108)
        Me.rdbSelected.Name = "rdbSelected"
        Me.rdbSelected.Size = New System.Drawing.Size(147, 20)
        Me.rdbSelected.TabIndex = 141
        Me.rdbSelected.TabStop = True
        Me.rdbSelected.Text = "Selected Features"
        Me.rdbSelected.UseVisualStyleBackColor = True
        '
        'rdbPartial
        '
        Me.rdbPartial.AutoSize = True
        Me.rdbPartial.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbPartial.Location = New System.Drawing.Point(430, 108)
        Me.rdbPartial.Name = "rdbPartial"
        Me.rdbPartial.Size = New System.Drawing.Size(244, 20)
        Me.rdbPartial.TabIndex = 140
        Me.rdbPartial.TabStop = True
        Me.rdbPartial.Text = "Partial (Dev Type Features Only)"
        Me.rdbPartial.UseVisualStyleBackColor = True
        '
        'rdbFull
        '
        Me.rdbFull.AutoSize = True
        Me.rdbFull.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbFull.Location = New System.Drawing.Point(273, 108)
        Me.rdbFull.Name = "rdbFull"
        Me.rdbFull.Size = New System.Drawing.Size(143, 20)
        Me.rdbFull.TabIndex = 139
        Me.rdbFull.TabStop = True
        Me.rdbFull.Text = "Full (All Features)"
        Me.rdbFull.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(4, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(254, 16)
        Me.Label3.TabIndex = 138
        Me.Label3.Text = "Processing Feature Selection Option:"
        '
        'chkAcres
        '
        Me.chkAcres.AutoSize = True
        Me.chkAcres.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.chkAcres.Location = New System.Drawing.Point(5, 144)
        Me.chkAcres.Name = "chkAcres"
        Me.chkAcres.Size = New System.Drawing.Size(166, 20)
        Me.chkAcres.TabIndex = 142
        Me.chkAcres.Text = "Acre Density Options"
        Me.chkAcres.UseVisualStyleBackColor = True
        '
        'lblAcres
        '
        Me.lblAcres.AutoSize = True
        Me.lblAcres.Enabled = False
        Me.lblAcres.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAcres.Location = New System.Drawing.Point(220, 145)
        Me.lblAcres.Name = "lblAcres"
        Me.lblAcres.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAcres.Size = New System.Drawing.Size(176, 16)
        Me.lblAcres.TabIndex = 143
        Me.lblAcres.Text = "Selected the Acres Field:"
        '
        'cmbAcresFld
        '
        Me.cmbAcresFld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAcresFld.Enabled = False
        Me.cmbAcresFld.FormattingEnabled = True
        Me.cmbAcresFld.Location = New System.Drawing.Point(402, 142)
        Me.cmbAcresFld.Name = "cmbAcresFld"
        Me.cmbAcresFld.Size = New System.Drawing.Size(452, 21)
        Me.cmbAcresFld.TabIndex = 144
        '
        'cmbSqMi
        '
        Me.cmbSqMi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSqMi.Enabled = False
        Me.cmbSqMi.FormattingEnabled = True
        Me.cmbSqMi.Location = New System.Drawing.Point(448, 168)
        Me.cmbSqMi.Name = "cmbSqMi"
        Me.cmbSqMi.Size = New System.Drawing.Size(406, 21)
        Me.cmbSqMi.TabIndex = 147
        '
        'lblSqMi
        '
        Me.lblSqMi.AutoSize = True
        Me.lblSqMi.Enabled = False
        Me.lblSqMi.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSqMi.Location = New System.Drawing.Point(220, 171)
        Me.lblSqMi.Name = "lblSqMi"
        Me.lblSqMi.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSqMi.Size = New System.Drawing.Size(222, 16)
        Me.lblSqMi.TabIndex = 146
        Me.lblSqMi.Text = "Selected the Square Miles Field:"
        '
        'chkSqMi
        '
        Me.chkSqMi.AutoSize = True
        Me.chkSqMi.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.chkSqMi.Location = New System.Drawing.Point(5, 170)
        Me.chkSqMi.Name = "chkSqMi"
        Me.chkSqMi.Size = New System.Drawing.Size(212, 20)
        Me.chkSqMi.TabIndex = 145
        Me.chkSqMi.Text = "Square Mile Density Options"
        Me.chkSqMi.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(1, 402)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(852, 4)
        Me.Panel1.TabIndex = 138
        '
        'tbxProjection
        '
        Me.tbxProjection.AllowDrop = True
        Me.tbxProjection.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxProjection.Location = New System.Drawing.Point(7, 651)
        Me.tbxProjection.Multiline = True
        Me.tbxProjection.Name = "tbxProjection"
        Me.tbxProjection.ReadOnly = True
        Me.tbxProjection.Size = New System.Drawing.Size(785, 28)
        Me.tbxProjection.TabIndex = 148
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(5, 631)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(188, 16)
        Me.Label6.TabIndex = 149
        Me.Label6.Text = "Clip Processing Projection: "
        '
        'btnSelectPrj
        '
        Me.btnSelectPrj.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelectPrj.BackColor = System.Drawing.Color.Transparent
        Me.btnSelectPrj.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnSelectPrj.Image = CType(resources.GetObject("btnSelectPrj.Image"), System.Drawing.Image)
        Me.btnSelectPrj.Location = New System.Drawing.Point(798, 636)
        Me.btnSelectPrj.Name = "btnSelectPrj"
        Me.btnSelectPrj.Size = New System.Drawing.Size(56, 48)
        Me.btnSelectPrj.TabIndex = 150
        Me.btnSelectPrj.UseVisualStyleBackColor = False
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.Controls.Add(Me.cmbLayers)
        Me.Panel6.Controls.Add(Me.rdbNeighborhoodLayer)
        Me.Panel6.Controls.Add(Me.Label4)
        Me.Panel6.Controls.Add(Me.rdbParcelLayer)
        Me.Panel6.Location = New System.Drawing.Point(0, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(857, 51)
        Me.Panel6.TabIndex = 151
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.Enabled = False
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(448, 16)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.Size = New System.Drawing.Size(404, 21)
        Me.cmbLayers.TabIndex = 139
        '
        'rdbNeighborhoodLayer
        '
        Me.rdbNeighborhoodLayer.AutoSize = True
        Me.rdbNeighborhoodLayer.Enabled = False
        Me.rdbNeighborhoodLayer.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbNeighborhoodLayer.Location = New System.Drawing.Point(273, 18)
        Me.rdbNeighborhoodLayer.Name = "rdbNeighborhoodLayer"
        Me.rdbNeighborhoodLayer.Size = New System.Drawing.Size(167, 20)
        Me.rdbNeighborhoodLayer.TabIndex = 137
        Me.rdbNeighborhoodLayer.Text = "Neighborhood Layer :"
        Me.rdbNeighborhoodLayer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rdbNeighborhoodLayer.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(125, 16)
        Me.Label4.TabIndex = 138
        Me.Label4.Text = "Processing Layer:"
        '
        'rdbParcelLayer
        '
        Me.rdbParcelLayer.AutoSize = True
        Me.rdbParcelLayer.Checked = True
        Me.rdbParcelLayer.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbParcelLayer.Location = New System.Drawing.Point(140, 18)
        Me.rdbParcelLayer.Name = "rdbParcelLayer"
        Me.rdbParcelLayer.Size = New System.Drawing.Size(107, 20)
        Me.rdbParcelLayer.TabIndex = 136
        Me.rdbParcelLayer.TabStop = True
        Me.rdbParcelLayer.Text = "Parcel Layer"
        Me.rdbParcelLayer.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(3, 92)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(852, 4)
        Me.Panel3.TabIndex = 138
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Location = New System.Drawing.Point(1, 686)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(860, 4)
        Me.Panel4.TabIndex = 139
        '
        'frmFeatureSum
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(861, 724)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.tbxProjection)
        Me.Controls.Add(Me.btnSelectPrj)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmbSqMi)
        Me.Controls.Add(Me.lblSqMi)
        Me.Controls.Add(Me.chkSqMi)
        Me.Controls.Add(Me.cmbAcresFld)
        Me.Controls.Add(Me.lblAcres)
        Me.Controls.Add(Me.chkAcres)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.rdbSelected)
        Me.Controls.Add(Me.rdbPartial)
        Me.Controls.Add(Me.rdbFull)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgvARC)
        Me.Controls.Add(Me.lblArc)
        Me.Controls.Add(Me.dgvPOINT)
        Me.Controls.Add(Me.lblPoint)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Name = "frmFeatureSum"
        Me.Text = "Feature Summary"
        CType(Me.dgvPOINT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        CType(Me.dgvARC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvPOINT As System.Windows.Forms.DataGridView
    Friend WithEvents lblPoint As System.Windows.Forms.Label
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents barStatusRun As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents lblRunStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents dgvARC As System.Windows.Forms.DataGridView
    Friend WithEvents lblArc As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents rdbSelected As System.Windows.Forms.RadioButton
    Friend WithEvents rdbPartial As System.Windows.Forms.RadioButton
    Friend WithEvents rdbFull As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkAcres As System.Windows.Forms.CheckBox
    Friend WithEvents lblAcres As System.Windows.Forms.Label
    Friend WithEvents cmbAcresFld As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSqMi As System.Windows.Forms.ComboBox
    Friend WithEvents lblSqMi As System.Windows.Forms.Label
    Friend WithEvents chkSqMi As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents tbxProjection As System.Windows.Forms.TextBox
    Friend WithEvents btnSelectPrj As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn3 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn4 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents rdbNeighborhoodLayer As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents rdbParcelLayer As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
End Class
