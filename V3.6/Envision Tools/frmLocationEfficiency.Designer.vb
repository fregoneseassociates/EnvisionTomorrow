<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLocationEfficiency
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
    Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLocationEfficiency))
    Me.InfoStationLocationSettings = New System.Windows.Forms.TabControl()
    Me.tabWeights = New System.Windows.Forms.TabPage()
    Me.lblAlias = New System.Windows.Forms.Label()
    Me.txtAlias = New System.Windows.Forms.TextBox()
    Me.dgvWeights = New System.Windows.Forms.DataGridView()
    Me.DataGridViewButtonColumn2 = New System.Windows.Forms.DataGridViewButtonColumn()
    Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.DataGridViewTextBoxColumnAlias = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.DataGridViewCheckBoxColumnKernel = New System.Windows.Forms.DataGridViewCheckBoxColumn()
    Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
    Me.lblTotalInfluence = New System.Windows.Forms.Label()
    Me.tbxTotalInfluence = New System.Windows.Forms.TextBox()
    Me.lblInfluence = New System.Windows.Forms.Label()
    Me.tbxInfluence = New System.Windows.Forms.TextBox()
    Me.btnAddFactor = New System.Windows.Forms.Button()
    Me.cmbValuesFld = New System.Windows.Forms.ComboBox()
    Me.lblValueField = New System.Windows.Forms.Label()
    Me.cmbLayers = New System.Windows.Forms.ComboBox()
    Me.lblDataLayer = New System.Windows.Forms.Label()
    Me.Panel4 = New System.Windows.Forms.Panel()
    Me.InfoGroup1 = New System.Windows.Forms.ToolStrip()
    Me.btnNext = New System.Windows.Forms.ToolStripButton()
    Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
    Me.tabSettings = New System.Windows.Forms.TabPage()
    Me.ChkDataHasOverlappingPolys = New System.Windows.Forms.CheckBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.cboZoneField = New System.Windows.Forms.ComboBox()
    Me.cboZoneLayer = New System.Windows.Forms.ComboBox()
    Me.Panel2 = New System.Windows.Forms.Panel()
    Me.lblSearchRadiusNote = New System.Windows.Forms.Label()
    Me.txtSearchRadius = New System.Windows.Forms.TextBox()
    Me.lblSearchRadius = New System.Windows.Forms.Label()
    Me.Panel9 = New System.Windows.Forms.Panel()
    Me.txtCellSize = New System.Windows.Forms.TextBox()
    Me.lblCellSizeFeet = New System.Windows.Forms.Label()
    Me.cmbExtentLayers = New System.Windows.Forms.ComboBox()
    Me.rdbExtentLayer = New System.Windows.Forms.RadioButton()
    Me.rdbExtentView = New System.Windows.Forms.RadioButton()
    Me.Label24 = New System.Windows.Forms.Label()
    Me.Panel3 = New System.Windows.Forms.Panel()
    Me.tbxProjectName = New System.Windows.Forms.TextBox()
    Me.lblNewFileGDBName = New System.Windows.Forms.Label()
    Me.tbxWorkspace = New System.Windows.Forms.TextBox()
    Me.btnWorkspace = New System.Windows.Forms.Button()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.InfoSettings = New System.Windows.Forms.ToolStrip()
    Me.btnRun = New System.Windows.Forms.ToolStripButton()
    Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
    Me.btnPrevious = New System.Windows.Forms.ToolStripButton()
    Me.ToolStripSeparator14 = New System.Windows.Forms.ToolStripSeparator()
    Me.barStatusRun = New System.Windows.Forms.ToolStripProgressBar()
    Me.lblRunStatus = New System.Windows.Forms.ToolStripLabel()
    Me.Panel7 = New System.Windows.Forms.Panel()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.InfoStationLocationSettings.SuspendLayout()
    Me.tabWeights.SuspendLayout()
    CType(Me.dgvWeights, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.InfoGroup1.SuspendLayout()
    Me.tabSettings.SuspendLayout()
    Me.InfoSettings.SuspendLayout()
    Me.Panel7.SuspendLayout()
    Me.SuspendLayout()
    '
    'InfoStationLocationSettings
    '
    Me.InfoStationLocationSettings.Appearance = System.Windows.Forms.TabAppearance.Buttons
    Me.InfoStationLocationSettings.Controls.Add(Me.tabWeights)
    Me.InfoStationLocationSettings.Controls.Add(Me.tabSettings)
    Me.InfoStationLocationSettings.Dock = System.Windows.Forms.DockStyle.Fill
    Me.InfoStationLocationSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.InfoStationLocationSettings.Location = New System.Drawing.Point(0, 0)
    Me.InfoStationLocationSettings.Multiline = True
    Me.InfoStationLocationSettings.Name = "InfoStationLocationSettings"
    Me.InfoStationLocationSettings.SelectedIndex = 0
    Me.InfoStationLocationSettings.Size = New System.Drawing.Size(805, 391)
    Me.InfoStationLocationSettings.TabIndex = 43
    '
    'tabWeights
    '
    Me.tabWeights.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.tabWeights.Controls.Add(Me.lblAlias)
    Me.tabWeights.Controls.Add(Me.txtAlias)
    Me.tabWeights.Controls.Add(Me.dgvWeights)
    Me.tabWeights.Controls.Add(Me.lblTotalInfluence)
    Me.tabWeights.Controls.Add(Me.tbxTotalInfluence)
    Me.tabWeights.Controls.Add(Me.lblInfluence)
    Me.tabWeights.Controls.Add(Me.tbxInfluence)
    Me.tabWeights.Controls.Add(Me.btnAddFactor)
    Me.tabWeights.Controls.Add(Me.cmbValuesFld)
    Me.tabWeights.Controls.Add(Me.lblValueField)
    Me.tabWeights.Controls.Add(Me.cmbLayers)
    Me.tabWeights.Controls.Add(Me.lblDataLayer)
    Me.tabWeights.Controls.Add(Me.Panel4)
    Me.tabWeights.Controls.Add(Me.InfoGroup1)
    Me.tabWeights.Location = New System.Drawing.Point(4, 28)
    Me.tabWeights.Name = "tabWeights"
    Me.tabWeights.Padding = New System.Windows.Forms.Padding(3)
    Me.tabWeights.Size = New System.Drawing.Size(797, 359)
    Me.tabWeights.TabIndex = 2
    Me.tabWeights.Text = "Weights"
    Me.tabWeights.UseVisualStyleBackColor = True
    '
    'lblAlias
    '
    Me.lblAlias.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblAlias.AutoSize = True
    Me.lblAlias.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblAlias.Location = New System.Drawing.Point(420, 3)
    Me.lblAlias.Name = "lblAlias"
    Me.lblAlias.Size = New System.Drawing.Size(41, 14)
    Me.lblAlias.TabIndex = 74
    Me.lblAlias.Text = "Alias:"
    '
    'txtAlias
    '
    Me.txtAlias.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtAlias.Location = New System.Drawing.Point(423, 23)
    Me.txtAlias.Name = "txtAlias"
    Me.txtAlias.Size = New System.Drawing.Size(177, 22)
    Me.txtAlias.TabIndex = 73
    '
    'dgvWeights
    '
    Me.dgvWeights.AllowUserToAddRows = False
    Me.dgvWeights.AllowUserToDeleteRows = False
    Me.dgvWeights.AllowUserToOrderColumns = True
    DataGridViewCellStyle9.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    DataGridViewCellStyle9.ForeColor = System.Drawing.Color.Black
    Me.dgvWeights.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle9
    Me.dgvWeights.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvWeights.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
    Me.dgvWeights.BackgroundColor = System.Drawing.SystemColors.Control
    DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
    DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    Me.dgvWeights.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
    Me.dgvWeights.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvWeights.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewButtonColumn2, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumnAlias, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewCheckBoxColumnKernel, Me.DataGridViewTextBoxColumn8})
    DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
    DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.MenuHighlight
    DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    Me.dgvWeights.DefaultCellStyle = DataGridViewCellStyle11
    Me.dgvWeights.Location = New System.Drawing.Point(4, 51)
    Me.dgvWeights.MultiSelect = False
    Me.dgvWeights.Name = "dgvWeights"
    Me.dgvWeights.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.dgvWeights.RowHeadersWidth = 20
    Me.dgvWeights.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
    DataGridViewCellStyle12.BackColor = System.Drawing.Color.White
    DataGridViewCellStyle12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    DataGridViewCellStyle12.ForeColor = System.Drawing.Color.Black
    Me.dgvWeights.RowsDefaultCellStyle = DataGridViewCellStyle12
    Me.dgvWeights.Size = New System.Drawing.Size(786, 235)
    Me.dgvWeights.TabIndex = 72
    '
    'DataGridViewButtonColumn2
    '
    Me.DataGridViewButtonColumn2.FillWeight = 6.526463!
    Me.DataGridViewButtonColumn2.HeaderText = ""
    Me.DataGridViewButtonColumn2.Name = "DataGridViewButtonColumn2"
    Me.DataGridViewButtonColumn2.ReadOnly = True
    Me.DataGridViewButtonColumn2.Text = "Remove"
    Me.DataGridViewButtonColumn2.UseColumnTextForButtonValue = True
    '
    'DataGridViewTextBoxColumn5
    '
    Me.DataGridViewTextBoxColumn5.FillWeight = 26.05293!
    Me.DataGridViewTextBoxColumn5.HeaderText = "Layer"
    Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
    Me.DataGridViewTextBoxColumn5.ReadOnly = True
    '
    'DataGridViewTextBoxColumnAlias
    '
    Me.DataGridViewTextBoxColumnAlias.FillWeight = 13.05293!
    Me.DataGridViewTextBoxColumnAlias.HeaderText = "Alias"
    Me.DataGridViewTextBoxColumnAlias.Name = "DataGridViewTextBoxColumnAlias"
    '
    'DataGridViewTextBoxColumn6
    '
    Me.DataGridViewTextBoxColumn6.FillWeight = 13.05293!
    Me.DataGridViewTextBoxColumn6.HeaderText = "Value Field"
    Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
    Me.DataGridViewTextBoxColumn6.ReadOnly = True
    '
    'DataGridViewTextBoxColumn7
    '
    Me.DataGridViewTextBoxColumn7.FillWeight = 13.05293!
    Me.DataGridViewTextBoxColumn7.HeaderText = "Influence %"
    Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
    '
    'DataGridViewCheckBoxColumnKernel
    '
    Me.DataGridViewCheckBoxColumnKernel.FillWeight = 6.0!
    Me.DataGridViewCheckBoxColumnKernel.HeaderText = "Density?"
    Me.DataGridViewCheckBoxColumnKernel.Name = "DataGridViewCheckBoxColumnKernel"
    Me.DataGridViewCheckBoxColumnKernel.ToolTipText = "Select to use Kernel Density processing.  Unselect to use Natural Neighbor Interp" & _
    "olation.  Applies only to point and line layers with a selected field."
    '
    'DataGridViewTextBoxColumn8
    '
    Me.DataGridViewTextBoxColumn8.HeaderText = "Layer Index"
    Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
    Me.DataGridViewTextBoxColumn8.ReadOnly = True
    Me.DataGridViewTextBoxColumn8.Visible = False
    '
    'lblTotalInfluence
    '
    Me.lblTotalInfluence.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblTotalInfluence.AutoSize = True
    Me.lblTotalInfluence.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblTotalInfluence.ForeColor = System.Drawing.SystemColors.ControlText
    Me.lblTotalInfluence.Location = New System.Drawing.Point(573, 293)
    Me.lblTotalInfluence.Name = "lblTotalInfluence"
    Me.lblTotalInfluence.Size = New System.Drawing.Size(121, 14)
    Me.lblTotalInfluence.TabIndex = 71
    Me.lblTotalInfluence.Text = "Total Influence %:"
    '
    'tbxTotalInfluence
    '
    Me.tbxTotalInfluence.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbxTotalInfluence.BackColor = System.Drawing.SystemColors.Control
    Me.tbxTotalInfluence.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbxTotalInfluence.ForeColor = System.Drawing.SystemColors.WindowText
    Me.tbxTotalInfluence.Location = New System.Drawing.Point(700, 289)
    Me.tbxTotalInfluence.Name = "tbxTotalInfluence"
    Me.tbxTotalInfluence.ReadOnly = True
    Me.tbxTotalInfluence.RightToLeft = System.Windows.Forms.RightToLeft.Yes
    Me.tbxTotalInfluence.Size = New System.Drawing.Size(90, 23)
    Me.tbxTotalInfluence.TabIndex = 70
    Me.tbxTotalInfluence.Text = "0"
    '
    'lblInfluence
    '
    Me.lblInfluence.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblInfluence.AutoSize = True
    Me.lblInfluence.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblInfluence.Location = New System.Drawing.Point(607, 3)
    Me.lblInfluence.Name = "lblInfluence"
    Me.lblInfluence.Size = New System.Drawing.Size(87, 14)
    Me.lblInfluence.TabIndex = 69
    Me.lblInfluence.Text = "Influence %:"
    '
    'tbxInfluence
    '
    Me.tbxInfluence.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbxInfluence.Location = New System.Drawing.Point(610, 23)
    Me.tbxInfluence.Name = "tbxInfluence"
    Me.tbxInfluence.Size = New System.Drawing.Size(84, 22)
    Me.tbxInfluence.TabIndex = 68
    '
    'btnAddFactor
    '
    Me.btnAddFactor.Enabled = False
    Me.btnAddFactor.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnAddFactor.Location = New System.Drawing.Point(700, 22)
    Me.btnAddFactor.Name = "btnAddFactor"
    Me.btnAddFactor.Size = New System.Drawing.Size(91, 23)
    Me.btnAddFactor.TabIndex = 67
    Me.btnAddFactor.Text = "Add Factor"
    Me.btnAddFactor.UseVisualStyleBackColor = True
    '
    'cmbValuesFld
    '
    Me.cmbValuesFld.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmbValuesFld.FormattingEnabled = True
    Me.cmbValuesFld.Location = New System.Drawing.Point(237, 23)
    Me.cmbValuesFld.Name = "cmbValuesFld"
    Me.cmbValuesFld.Size = New System.Drawing.Size(177, 24)
    Me.cmbValuesFld.TabIndex = 66
    '
    'lblValueField
    '
    Me.lblValueField.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblValueField.AutoSize = True
    Me.lblValueField.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblValueField.Location = New System.Drawing.Point(234, 3)
    Me.lblValueField.Name = "lblValueField"
    Me.lblValueField.Size = New System.Drawing.Size(143, 14)
    Me.lblValueField.TabIndex = 65
    Me.lblValueField.Text = "Select the value field:"
    '
    'cmbLayers
    '
    Me.cmbLayers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmbLayers.FormattingEnabled = True
    Me.cmbLayers.Location = New System.Drawing.Point(4, 23)
    Me.cmbLayers.Name = "cmbLayers"
    Me.cmbLayers.Size = New System.Drawing.Size(224, 24)
    Me.cmbLayers.TabIndex = 64
    '
    'lblDataLayer
    '
    Me.lblDataLayer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblDataLayer.AutoSize = True
    Me.lblDataLayer.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblDataLayer.Location = New System.Drawing.Point(1, 3)
    Me.lblDataLayer.Name = "lblDataLayer"
    Me.lblDataLayer.Size = New System.Drawing.Size(130, 14)
    Me.lblDataLayer.TabIndex = 62
    Me.lblDataLayer.Text = "Select a data layer:"
    '
    'Panel4
    '
    Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel4.Location = New System.Drawing.Point(2, 314)
    Me.Panel4.Name = "Panel4"
    Me.Panel4.Size = New System.Drawing.Size(790, 4)
    Me.Panel4.TabIndex = 44
    '
    'InfoGroup1
    '
    Me.InfoGroup1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.InfoGroup1.ImageScalingSize = New System.Drawing.Size(24, 24)
    Me.InfoGroup1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnNext, Me.ToolStripSeparator13})
    Me.InfoGroup1.Location = New System.Drawing.Point(3, 321)
    Me.InfoGroup1.Name = "InfoGroup1"
    Me.InfoGroup1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
    Me.InfoGroup1.Size = New System.Drawing.Size(787, 31)
    Me.InfoGroup1.TabIndex = 37
    Me.InfoGroup1.Text = "ToolStrip4"
    '
    'btnNext
    '
    Me.btnNext.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnNext.Image = CType(resources.GetObject("btnNext.Image"), System.Drawing.Image)
    Me.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.btnNext.Name = "btnNext"
    Me.btnNext.Size = New System.Drawing.Size(70, 28)
    Me.btnNext.Text = "Next"
    '
    'ToolStripSeparator13
    '
    Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
    Me.ToolStripSeparator13.Size = New System.Drawing.Size(6, 31)
    '
    'tabSettings
    '
    Me.tabSettings.AutoScroll = True
    Me.tabSettings.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.tabSettings.Controls.Add(Me.ChkDataHasOverlappingPolys)
    Me.tabSettings.Controls.Add(Me.Label1)
    Me.tabSettings.Controls.Add(Me.Label4)
    Me.tabSettings.Controls.Add(Me.Label3)
    Me.tabSettings.Controls.Add(Me.cboZoneField)
    Me.tabSettings.Controls.Add(Me.cboZoneLayer)
    Me.tabSettings.Controls.Add(Me.Panel2)
    Me.tabSettings.Controls.Add(Me.lblSearchRadiusNote)
    Me.tabSettings.Controls.Add(Me.txtSearchRadius)
    Me.tabSettings.Controls.Add(Me.lblSearchRadius)
    Me.tabSettings.Controls.Add(Me.Panel9)
    Me.tabSettings.Controls.Add(Me.txtCellSize)
    Me.tabSettings.Controls.Add(Me.lblCellSizeFeet)
    Me.tabSettings.Controls.Add(Me.cmbExtentLayers)
    Me.tabSettings.Controls.Add(Me.rdbExtentLayer)
    Me.tabSettings.Controls.Add(Me.rdbExtentView)
    Me.tabSettings.Controls.Add(Me.Label24)
    Me.tabSettings.Controls.Add(Me.Panel3)
    Me.tabSettings.Controls.Add(Me.tbxProjectName)
    Me.tabSettings.Controls.Add(Me.lblNewFileGDBName)
    Me.tabSettings.Controls.Add(Me.tbxWorkspace)
    Me.tabSettings.Controls.Add(Me.btnWorkspace)
    Me.tabSettings.Controls.Add(Me.Label6)
    Me.tabSettings.Controls.Add(Me.InfoSettings)
    Me.tabSettings.Controls.Add(Me.Panel7)
    Me.tabSettings.Location = New System.Drawing.Point(4, 28)
    Me.tabSettings.Name = "tabSettings"
    Me.tabSettings.Padding = New System.Windows.Forms.Padding(3)
    Me.tabSettings.Size = New System.Drawing.Size(797, 359)
    Me.tabSettings.TabIndex = 5
    Me.tabSettings.Text = "Process Settings"
    Me.tabSettings.UseVisualStyleBackColor = True
    '
    'ChkDataHasOverlappingPolys
    '
    Me.ChkDataHasOverlappingPolys.AutoSize = True
    Me.ChkDataHasOverlappingPolys.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
    Me.ChkDataHasOverlappingPolys.Location = New System.Drawing.Point(9, 193)
    Me.ChkDataHasOverlappingPolys.Name = "ChkDataHasOverlappingPolys"
    Me.ChkDataHasOverlappingPolys.Size = New System.Drawing.Size(294, 20)
    Me.ChkDataHasOverlappingPolys.TabIndex = 146
    Me.ChkDataHasOverlappingPolys.Text = "Feature zone data has overlapping polygons"
    Me.ChkDataHasOverlappingPolys.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(6, 135)
    Me.Label1.Name = "Label1"
    Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label1.Size = New System.Drawing.Size(322, 16)
    Me.Label1.TabIndex = 145
    Me.Label1.Text = "Select inputs for zonal statistics (optional):"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label4.Location = New System.Drawing.Point(525, 166)
    Me.Label4.Name = "Label4"
    Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label4.Size = New System.Drawing.Size(79, 16)
    Me.Label4.TabIndex = 144
    Me.Label4.Text = "Zone field:"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label3.Location = New System.Drawing.Point(6, 166)
    Me.Label3.Name = "Label3"
    Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label3.Size = New System.Drawing.Size(198, 16)
    Me.Label3.TabIndex = 143
    Me.Label3.Text = "Raster or feature zone data:"
    '
    'cboZoneField
    '
    Me.cboZoneField.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboZoneField.FormattingEnabled = True
    Me.cboZoneField.Location = New System.Drawing.Point(608, 163)
    Me.cboZoneField.Name = "cboZoneField"
    Me.cboZoneField.Size = New System.Drawing.Size(184, 24)
    Me.cboZoneField.TabIndex = 142
    '
    'cboZoneLayer
    '
    Me.cboZoneLayer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboZoneLayer.FormattingEnabled = True
    Me.cboZoneLayer.Location = New System.Drawing.Point(209, 163)
    Me.cboZoneLayer.Name = "cboZoneLayer"
    Me.cboZoneLayer.Size = New System.Drawing.Size(295, 24)
    Me.cboZoneLayer.TabIndex = 140
    '
    'Panel2
    '
    Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel2.Location = New System.Drawing.Point(1, 119)
    Me.Panel2.Name = "Panel2"
    Me.Panel2.Size = New System.Drawing.Size(795, 4)
    Me.Panel2.TabIndex = 138
    '
    'lblSearchRadiusNote
    '
    Me.lblSearchRadiusNote.AutoSize = True
    Me.lblSearchRadiusNote.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblSearchRadiusNote.Location = New System.Drawing.Point(526, 89)
    Me.lblSearchRadiusNote.Name = "lblSearchRadiusNote"
    Me.lblSearchRadiusNote.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblSearchRadiusNote.Size = New System.Drawing.Size(262, 16)
    Me.lblSearchRadiusNote.TabIndex = 137
    Me.lblSearchRadiusNote.Text = "(for point and line input density calcs)"
    '
    'txtSearchRadius
    '
    Me.txtSearchRadius.AllowDrop = True
    Me.txtSearchRadius.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtSearchRadius.BackColor = System.Drawing.Color.White
    Me.txtSearchRadius.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtSearchRadius.Location = New System.Drawing.Point(471, 86)
    Me.txtSearchRadius.Name = "txtSearchRadius"
    Me.txtSearchRadius.Size = New System.Drawing.Size(51, 22)
    Me.txtSearchRadius.TabIndex = 136
    Me.txtSearchRadius.Text = "1320"
    '
    'lblSearchRadius
    '
    Me.lblSearchRadius.AutoSize = True
    Me.lblSearchRadius.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblSearchRadius.Location = New System.Drawing.Point(317, 89)
    Me.lblSearchRadius.Name = "lblSearchRadius"
    Me.lblSearchRadius.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblSearchRadius.Size = New System.Drawing.Size(148, 16)
    Me.lblSearchRadius.TabIndex = 135
    Me.lblSearchRadius.Text = "Search radius (feet):"
    '
    'Panel9
    '
    Me.Panel9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel9.Location = New System.Drawing.Point(1, 71)
    Me.Panel9.Name = "Panel9"
    Me.Panel9.Size = New System.Drawing.Size(795, 4)
    Me.Panel9.TabIndex = 134
    '
    'txtCellSize
    '
    Me.txtCellSize.AllowDrop = True
    Me.txtCellSize.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtCellSize.BackColor = System.Drawing.Color.White
    Me.txtCellSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtCellSize.Location = New System.Drawing.Point(218, 86)
    Me.txtCellSize.Name = "txtCellSize"
    Me.txtCellSize.Size = New System.Drawing.Size(51, 22)
    Me.txtCellSize.TabIndex = 133
    Me.txtCellSize.Text = "264"
    '
    'lblCellSizeFeet
    '
    Me.lblCellSizeFeet.AutoSize = True
    Me.lblCellSizeFeet.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblCellSizeFeet.Location = New System.Drawing.Point(6, 89)
    Me.lblCellSizeFeet.Name = "lblCellSizeFeet"
    Me.lblCellSizeFeet.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblCellSizeFeet.Size = New System.Drawing.Size(206, 16)
    Me.lblCellSizeFeet.TabIndex = 132
    Me.lblCellSizeFeet.Text = "Raster output cell size (feet):"
    '
    'cmbExtentLayers
    '
    Me.cmbExtentLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmbExtentLayers.Enabled = False
    Me.cmbExtentLayers.FormattingEnabled = True
    Me.cmbExtentLayers.Location = New System.Drawing.Point(330, 31)
    Me.cmbExtentLayers.Name = "cmbExtentLayers"
    Me.cmbExtentLayers.Size = New System.Drawing.Size(462, 24)
    Me.cmbExtentLayers.TabIndex = 129
    '
    'rdbExtentLayer
    '
    Me.rdbExtentLayer.AutoSize = True
    Me.rdbExtentLayer.Location = New System.Drawing.Point(165, 33)
    Me.rdbExtentLayer.Name = "rdbExtentLayer"
    Me.rdbExtentLayer.Size = New System.Drawing.Size(154, 20)
    Me.rdbExtentLayer.TabIndex = 128
    Me.rdbExtentLayer.Text = "Selected layer extent:"
    Me.rdbExtentLayer.UseVisualStyleBackColor = True
    '
    'rdbExtentView
    '
    Me.rdbExtentView.AutoSize = True
    Me.rdbExtentView.Location = New System.Drawing.Point(9, 33)
    Me.rdbExtentView.Name = "rdbExtentView"
    Me.rdbExtentView.Size = New System.Drawing.Size(136, 20)
    Me.rdbExtentView.TabIndex = 126
    Me.rdbExtentView.Text = "Current view extent"
    Me.rdbExtentView.UseVisualStyleBackColor = True
    '
    'Label24
    '
    Me.Label24.AutoSize = True
    Me.Label24.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label24.Location = New System.Drawing.Point(6, 7)
    Me.Label24.Name = "Label24"
    Me.Label24.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label24.Size = New System.Drawing.Size(290, 16)
    Me.Label24.TabIndex = 125
    Me.Label24.Text = "Select the extent for the raster output:"
    '
    'Panel3
    '
    Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel3.Location = New System.Drawing.Point(1, 220)
    Me.Panel3.Name = "Panel3"
    Me.Panel3.Size = New System.Drawing.Size(795, 4)
    Me.Panel3.TabIndex = 124
    '
    'tbxProjectName
    '
    Me.tbxProjectName.AllowDrop = True
    Me.tbxProjectName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbxProjectName.BackColor = System.Drawing.Color.White
    Me.tbxProjectName.Enabled = False
    Me.tbxProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbxProjectName.Location = New System.Drawing.Point(233, 284)
    Me.tbxProjectName.Multiline = True
    Me.tbxProjectName.Name = "tbxProjectName"
    Me.tbxProjectName.Size = New System.Drawing.Size(559, 22)
    Me.tbxProjectName.TabIndex = 103
    Me.tbxProjectName.Text = "LOCATON_EFFICIENCY"
    '
    'lblNewFileGDBName
    '
    Me.lblNewFileGDBName.AutoSize = True
    Me.lblNewFileGDBName.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblNewFileGDBName.Location = New System.Drawing.Point(2, 287)
    Me.lblNewFileGDBName.Name = "lblNewFileGDBName"
    Me.lblNewFileGDBName.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblNewFileGDBName.Size = New System.Drawing.Size(214, 16)
    Me.lblNewFileGDBName.TabIndex = 104
    Me.lblNewFileGDBName.Text = "New file geodatabase name:"
    '
    'tbxWorkspace
    '
    Me.tbxWorkspace.AllowDrop = True
    Me.tbxWorkspace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbxWorkspace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbxWorkspace.Location = New System.Drawing.Point(171, 242)
    Me.tbxWorkspace.Multiline = True
    Me.tbxWorkspace.Name = "tbxWorkspace"
    Me.tbxWorkspace.ReadOnly = True
    Me.tbxWorkspace.Size = New System.Drawing.Size(558, 22)
    Me.tbxWorkspace.TabIndex = 100
    '
    'btnWorkspace
    '
    Me.btnWorkspace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnWorkspace.BackColor = System.Drawing.Color.Transparent
    Me.btnWorkspace.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.btnWorkspace.Image = CType(resources.GetObject("btnWorkspace.Image"), System.Drawing.Image)
    Me.btnWorkspace.Location = New System.Drawing.Point(735, 229)
    Me.btnWorkspace.Name = "btnWorkspace"
    Me.btnWorkspace.Size = New System.Drawing.Size(56, 48)
    Me.btnWorkspace.TabIndex = 102
    Me.btnWorkspace.UseVisualStyleBackColor = False
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label6.Location = New System.Drawing.Point(2, 245)
    Me.Label6.Name = "Label6"
    Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label6.Size = New System.Drawing.Size(163, 16)
    Me.Label6.TabIndex = 101
    Me.Label6.Text = "Workspace directory:"
    '
    'InfoSettings
    '
    Me.InfoSettings.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.InfoSettings.ImageScalingSize = New System.Drawing.Size(24, 24)
    Me.InfoSettings.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnRun, Me.ToolStripSeparator11, Me.btnPrevious, Me.ToolStripSeparator14, Me.barStatusRun, Me.lblRunStatus})
    Me.InfoSettings.Location = New System.Drawing.Point(3, 321)
    Me.InfoSettings.Name = "InfoSettings"
    Me.InfoSettings.RightToLeft = System.Windows.Forms.RightToLeft.Yes
    Me.InfoSettings.Size = New System.Drawing.Size(787, 31)
    Me.InfoSettings.TabIndex = 53
    Me.InfoSettings.Text = "ToolStrip4"
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
    'ToolStripSeparator11
    '
    Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
    Me.ToolStripSeparator11.Size = New System.Drawing.Size(6, 31)
    '
    'btnPrevious
    '
    Me.btnPrevious.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btnPrevious.Image = CType(resources.GetObject("btnPrevious.Image"), System.Drawing.Image)
    Me.btnPrevious.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.btnPrevious.Name = "btnPrevious"
    Me.btnPrevious.Size = New System.Drawing.Size(99, 28)
    Me.btnPrevious.Text = "Previous"
    Me.btnPrevious.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
    '
    'ToolStripSeparator14
    '
    Me.ToolStripSeparator14.Name = "ToolStripSeparator14"
    Me.ToolStripSeparator14.Size = New System.Drawing.Size(6, 31)
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
    'Panel7
    '
    Me.Panel7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel7.Controls.Add(Me.Panel1)
    Me.Panel7.Location = New System.Drawing.Point(2, 313)
    Me.Panel7.Name = "Panel7"
    Me.Panel7.Size = New System.Drawing.Size(795, 4)
    Me.Panel7.TabIndex = 52
    '
    'Panel1
    '
    Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel1.Location = New System.Drawing.Point(-2, -4)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(706, 4)
    Me.Panel1.TabIndex = 53
    '
    'frmLocationEfficiency
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(805, 391)
    Me.Controls.Add(Me.InfoStationLocationSettings)
    Me.MaximizeBox = False
    Me.Name = "frmLocationEfficiency"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Location Efficiency Tool"
    Me.InfoStationLocationSettings.ResumeLayout(False)
    Me.tabWeights.ResumeLayout(False)
    Me.tabWeights.PerformLayout()
    CType(Me.dgvWeights, System.ComponentModel.ISupportInitialize).EndInit()
    Me.InfoGroup1.ResumeLayout(False)
    Me.InfoGroup1.PerformLayout()
    Me.tabSettings.ResumeLayout(False)
    Me.tabSettings.PerformLayout()
    Me.InfoSettings.ResumeLayout(False)
    Me.InfoSettings.PerformLayout()
    Me.Panel7.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents InfoStationLocationSettings As System.Windows.Forms.TabControl
  Friend WithEvents tabWeights As System.Windows.Forms.TabPage
  Friend WithEvents Panel4 As System.Windows.Forms.Panel
  Friend WithEvents InfoGroup1 As System.Windows.Forms.ToolStrip
  Friend WithEvents btnNext As System.Windows.Forms.ToolStripButton
  Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents tabSettings As System.Windows.Forms.TabPage
  Friend WithEvents InfoSettings As System.Windows.Forms.ToolStrip
  Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
  Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents btnPrevious As System.Windows.Forms.ToolStripButton
  Friend WithEvents ToolStripSeparator14 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents barStatusRun As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents lblRunStatus As System.Windows.Forms.ToolStripLabel
  Friend WithEvents Panel7 As System.Windows.Forms.Panel
  Friend WithEvents tbxWorkspace As System.Windows.Forms.TextBox
  Friend WithEvents btnWorkspace As System.Windows.Forms.Button
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents tbxProjectName As System.Windows.Forms.TextBox
  Friend WithEvents lblNewFileGDBName As System.Windows.Forms.Label
  Friend WithEvents lblTotalInfluence As System.Windows.Forms.Label
  Friend WithEvents tbxTotalInfluence As System.Windows.Forms.TextBox
  Friend WithEvents lblInfluence As System.Windows.Forms.Label
  Friend WithEvents tbxInfluence As System.Windows.Forms.TextBox
  Friend WithEvents btnAddFactor As System.Windows.Forms.Button
  Friend WithEvents cmbValuesFld As System.Windows.Forms.ComboBox
  Friend WithEvents lblValueField As System.Windows.Forms.Label
  Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
  Friend WithEvents lblDataLayer As System.Windows.Forms.Label
  Friend WithEvents dgvWeights As System.Windows.Forms.DataGridView
  Friend WithEvents Panel3 As System.Windows.Forms.Panel
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents cmbExtentLayers As System.Windows.Forms.ComboBox
  Friend WithEvents rdbExtentLayer As System.Windows.Forms.RadioButton
  Friend WithEvents rdbExtentView As System.Windows.Forms.RadioButton
  Friend WithEvents Label24 As System.Windows.Forms.Label
  Friend WithEvents txtCellSize As System.Windows.Forms.TextBox
  Friend WithEvents lblCellSizeFeet As System.Windows.Forms.Label
  Friend WithEvents Panel9 As System.Windows.Forms.Panel
  Friend WithEvents lblSearchRadiusNote As System.Windows.Forms.Label
  Friend WithEvents txtSearchRadius As System.Windows.Forms.TextBox
  Friend WithEvents lblSearchRadius As System.Windows.Forms.Label
  Friend WithEvents lblAlias As System.Windows.Forms.Label
  Friend WithEvents txtAlias As System.Windows.Forms.TextBox
  Friend WithEvents DataGridViewButtonColumn2 As System.Windows.Forms.DataGridViewButtonColumn
  Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumnAlias As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewCheckBoxColumnKernel As System.Windows.Forms.DataGridViewCheckBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Panel2 As System.Windows.Forms.Panel
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents cboZoneField As System.Windows.Forms.ComboBox
  Friend WithEvents cboZoneLayer As System.Windows.Forms.ComboBox
  Friend WithEvents ChkDataHasOverlappingPolys As System.Windows.Forms.CheckBox
End Class
