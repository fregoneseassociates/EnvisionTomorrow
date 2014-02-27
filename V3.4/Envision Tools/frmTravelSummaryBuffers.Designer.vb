<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTravelSummaryBuffers
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTravelSummaryBuffers))
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbxWorkspace = New System.Windows.Forms.TextBox()
        Me.chkUseSelected = New System.Windows.Forms.CheckBox()
        Me.cmbFieldsID = New System.Windows.Forms.ComboBox()
        Me.lblFieldsId = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.barStatusRun = New System.Windows.Forms.ToolStripProgressBar()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.lblStNetowrk = New System.Windows.Forms.Label()
        Me.tbxNetworkLyr = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnWorkspace = New System.Windows.Forms.Button()
        Me.btnNetworkLyr = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnUnselect = New System.Windows.Forms.Button()
        Me.btnCheckAll = New System.Windows.Forms.Button()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.chkTransit = New System.Windows.Forms.CheckBox()
        Me.chkOneMiAuto = New System.Windows.Forms.CheckBox()
        Me.chkHalfMiAuto = New System.Windows.Forms.CheckBox()
        Me.chkQtrMiAuto = New System.Windows.Forms.CheckBox()
        Me.chk30min = New System.Windows.Forms.CheckBox()
        Me.chk20min = New System.Windows.Forms.CheckBox()
        Me.chk10min = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rdbStandard = New System.Windows.Forms.RadioButton()
        Me.rdbNetwork = New System.Windows.Forms.RadioButton()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(5, 151)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(230, 16)
        Me.Label6.TabIndex = 114
        Me.Label6.Text = "Select an Output File Geodatase:"
        '
        'tbxWorkspace
        '
        Me.tbxWorkspace.AllowDrop = True
        Me.tbxWorkspace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxWorkspace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxWorkspace.Location = New System.Drawing.Point(5, 179)
        Me.tbxWorkspace.Multiline = True
        Me.tbxWorkspace.Name = "tbxWorkspace"
        Me.tbxWorkspace.ReadOnly = True
        Me.tbxWorkspace.Size = New System.Drawing.Size(480, 28)
        Me.tbxWorkspace.TabIndex = 113
        '
        'chkUseSelected
        '
        Me.chkUseSelected.AutoSize = True
        Me.chkUseSelected.Enabled = False
        Me.chkUseSelected.Location = New System.Drawing.Point(231, 78)
        Me.chkUseSelected.Name = "chkUseSelected"
        Me.chkUseSelected.Size = New System.Drawing.Size(140, 17)
        Me.chkUseSelected.TabIndex = 143
        Me.chkUseSelected.Text = "Use Selected Feature(s)"
        Me.chkUseSelected.UseVisualStyleBackColor = True
        '
        'cmbFieldsID
        '
        Me.cmbFieldsID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbFieldsID.Enabled = False
        Me.cmbFieldsID.FormattingEnabled = True
        Me.cmbFieldsID.Location = New System.Drawing.Point(228, 106)
        Me.cmbFieldsID.Name = "cmbFieldsID"
        Me.cmbFieldsID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbFieldsID.Size = New System.Drawing.Size(326, 21)
        Me.cmbFieldsID.TabIndex = 142
        '
        'lblFieldsId
        '
        Me.lblFieldsId.AutoSize = True
        Me.lblFieldsId.Enabled = False
        Me.lblFieldsId.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldsId.Location = New System.Drawing.Point(50, 107)
        Me.lblFieldsId.Name = "lblFieldsId"
        Me.lblFieldsId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFieldsId.Size = New System.Drawing.Size(175, 16)
        Me.lblFieldsId.TabIndex = 141
        Me.lblFieldsId.Text = "Select a UNIQUE ID field:"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Location = New System.Drawing.Point(9, 140)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(545, 4)
        Me.Panel2.TabIndex = 139
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(228, 49)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLayers.Size = New System.Drawing.Size(326, 21)
        Me.cmbLayers.TabIndex = 140
        Me.cmbLayers.Text = "Select a layer"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(5, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(214, 16)
        Me.Label5.TabIndex = 138
        Me.Label5.Text = "Input Layer (Point or Polygon):"
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.barStatusRun, Me.lblStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 513)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(571, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 135
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
        'lblStatus
        '
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 28)
        '
        'lblStNetowrk
        '
        Me.lblStNetowrk.AutoSize = True
        Me.lblStNetowrk.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStNetowrk.Location = New System.Drawing.Point(6, 241)
        Me.lblStNetowrk.Name = "lblStNetowrk"
        Me.lblStNetowrk.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStNetowrk.Size = New System.Drawing.Size(175, 16)
        Me.lblStNetowrk.TabIndex = 132
        Me.lblStNetowrk.Text = "Select a Street Network:"
        '
        'tbxNetworkLyr
        '
        Me.tbxNetworkLyr.AllowDrop = True
        Me.tbxNetworkLyr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxNetworkLyr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxNetworkLyr.Location = New System.Drawing.Point(5, 260)
        Me.tbxNetworkLyr.Multiline = True
        Me.tbxNetworkLyr.Name = "tbxNetworkLyr"
        Me.tbxNetworkLyr.ReadOnly = True
        Me.tbxNetworkLyr.Size = New System.Drawing.Size(480, 28)
        Me.tbxNetworkLyr.TabIndex = 131
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(9, 229)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(545, 4)
        Me.Panel1.TabIndex = 140
        '
        'btnWorkspace
        '
        Me.btnWorkspace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnWorkspace.BackColor = System.Drawing.Color.Transparent
        Me.btnWorkspace.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnWorkspace.Image = CType(resources.GetObject("btnWorkspace.Image"), System.Drawing.Image)
        Me.btnWorkspace.Location = New System.Drawing.Point(498, 166)
        Me.btnWorkspace.Name = "btnWorkspace"
        Me.btnWorkspace.Size = New System.Drawing.Size(56, 48)
        Me.btnWorkspace.TabIndex = 115
        Me.btnWorkspace.UseVisualStyleBackColor = False
        '
        'btnNetworkLyr
        '
        Me.btnNetworkLyr.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNetworkLyr.BackColor = System.Drawing.Color.Transparent
        Me.btnNetworkLyr.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNetworkLyr.Image = CType(resources.GetObject("btnNetworkLyr.Image"), System.Drawing.Image)
        Me.btnNetworkLyr.Location = New System.Drawing.Point(498, 247)
        Me.btnNetworkLyr.Name = "btnNetworkLyr"
        Me.btnNetworkLyr.Size = New System.Drawing.Size(56, 48)
        Me.btnNetworkLyr.TabIndex = 133
        Me.btnNetworkLyr.UseVisualStyleBackColor = False
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.TextBox1)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.btnUnselect)
        Me.Panel3.Controls.Add(Me.btnCheckAll)
        Me.Panel3.Controls.Add(Me.Panel9)
        Me.Panel3.Controls.Add(Me.chkTransit)
        Me.Panel3.Controls.Add(Me.chkOneMiAuto)
        Me.Panel3.Controls.Add(Me.chkHalfMiAuto)
        Me.Panel3.Controls.Add(Me.chkQtrMiAuto)
        Me.Panel3.Controls.Add(Me.chk30min)
        Me.Panel3.Controls.Add(Me.chk20min)
        Me.Panel3.Controls.Add(Me.chk10min)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.Panel5)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.rdbStandard)
        Me.Panel3.Controls.Add(Me.rdbNetwork)
        Me.Panel3.Controls.Add(Me.Panel4)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.btnNetworkLyr)
        Me.Panel3.Controls.Add(Me.cmbLayers)
        Me.Panel3.Controls.Add(Me.btnWorkspace)
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.lblFieldsId)
        Me.Panel3.Controls.Add(Me.tbxWorkspace)
        Me.Panel3.Controls.Add(Me.cmbFieldsID)
        Me.Panel3.Controls.Add(Me.chkUseSelected)
        Me.Panel3.Controls.Add(Me.tbxNetworkLyr)
        Me.Panel3.Controls.Add(Me.lblStNetowrk)
        Me.Panel3.Location = New System.Drawing.Point(1, 1)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(566, 509)
        Me.Panel3.TabIndex = 148
        '
        'btnUnselect
        '
        Me.btnUnselect.Location = New System.Drawing.Point(377, 373)
        Me.btnUnselect.Name = "btnUnselect"
        Me.btnUnselect.Size = New System.Drawing.Size(75, 23)
        Me.btnUnselect.TabIndex = 160
        Me.btnUnselect.Text = "Unselect All"
        Me.btnUnselect.UseVisualStyleBackColor = True
        '
        'btnCheckAll
        '
        Me.btnCheckAll.Location = New System.Drawing.Point(296, 373)
        Me.btnCheckAll.Name = "btnCheckAll"
        Me.btnCheckAll.Size = New System.Drawing.Size(75, 23)
        Me.btnCheckAll.TabIndex = 157
        Me.btnCheckAll.Text = "Select All"
        Me.btnCheckAll.UseVisualStyleBackColor = True
        '
        'Panel9
        '
        Me.Panel9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel9.Controls.Add(Me.Panel10)
        Me.Panel9.Location = New System.Drawing.Point(9, 499)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(545, 4)
        Me.Panel9.TabIndex = 156
        '
        'Panel10
        '
        Me.Panel10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Location = New System.Drawing.Point(-2, 99)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(545, 4)
        Me.Panel10.TabIndex = 141
        '
        'chkTransit
        '
        Me.chkTransit.AutoSize = True
        Me.chkTransit.Checked = True
        Me.chkTransit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTransit.Location = New System.Drawing.Point(302, 404)
        Me.chkTransit.Name = "chkTransit"
        Me.chkTransit.Size = New System.Drawing.Size(58, 17)
        Me.chkTransit.TabIndex = 155
        Me.chkTransit.Text = "Transit"
        Me.chkTransit.UseVisualStyleBackColor = True
        '
        'chkOneMiAuto
        '
        Me.chkOneMiAuto.AutoSize = True
        Me.chkOneMiAuto.Checked = True
        Me.chkOneMiAuto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOneMiAuto.Location = New System.Drawing.Point(162, 471)
        Me.chkOneMiAuto.Name = "chkOneMiAuto"
        Me.chkOneMiAuto.Size = New System.Drawing.Size(93, 17)
        Me.chkOneMiAuto.TabIndex = 154
        Me.chkOneMiAuto.Text = "One Mile Auto"
        Me.chkOneMiAuto.UseVisualStyleBackColor = True
        '
        'chkHalfMiAuto
        '
        Me.chkHalfMiAuto.AutoSize = True
        Me.chkHalfMiAuto.Checked = True
        Me.chkHalfMiAuto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHalfMiAuto.Location = New System.Drawing.Point(162, 437)
        Me.chkHalfMiAuto.Name = "chkHalfMiAuto"
        Me.chkHalfMiAuto.Size = New System.Drawing.Size(92, 17)
        Me.chkHalfMiAuto.TabIndex = 153
        Me.chkHalfMiAuto.Text = "Half Mile Auto"
        Me.chkHalfMiAuto.UseVisualStyleBackColor = True
        '
        'chkQtrMiAuto
        '
        Me.chkQtrMiAuto.AutoSize = True
        Me.chkQtrMiAuto.Checked = True
        Me.chkQtrMiAuto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkQtrMiAuto.Location = New System.Drawing.Point(162, 404)
        Me.chkQtrMiAuto.Name = "chkQtrMiAuto"
        Me.chkQtrMiAuto.Size = New System.Drawing.Size(108, 17)
        Me.chkQtrMiAuto.TabIndex = 152
        Me.chkQtrMiAuto.Text = "Quarter Mile Auto"
        Me.chkQtrMiAuto.UseVisualStyleBackColor = True
        '
        'chk30min
        '
        Me.chk30min.AutoSize = True
        Me.chk30min.Checked = True
        Me.chk30min.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk30min.Location = New System.Drawing.Point(9, 466)
        Me.chk30min.Name = "chk30min"
        Me.chk30min.Size = New System.Drawing.Size(98, 17)
        Me.chk30min.TabIndex = 151
        Me.chk30min.Text = "30 Minute Auto"
        Me.chk30min.UseVisualStyleBackColor = True
        '
        'chk20min
        '
        Me.chk20min.AutoSize = True
        Me.chk20min.Checked = True
        Me.chk20min.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk20min.Location = New System.Drawing.Point(9, 432)
        Me.chk20min.Name = "chk20min"
        Me.chk20min.Size = New System.Drawing.Size(98, 17)
        Me.chk20min.TabIndex = 150
        Me.chk20min.Text = "20 Minute Auto"
        Me.chk20min.UseVisualStyleBackColor = True
        '
        'chk10min
        '
        Me.chk10min.AutoSize = True
        Me.chk10min.Checked = True
        Me.chk10min.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk10min.Location = New System.Drawing.Point(9, 399)
        Me.chk10min.Name = "chk10min"
        Me.chk10min.Size = New System.Drawing.Size(98, 17)
        Me.chk10min.TabIndex = 149
        Me.chk10min.Text = "10 Minute Auto"
        Me.chk10min.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 376)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(283, 16)
        Me.Label2.TabIndex = 148
        Me.Label2.Text = "Select the Buffer Layer(s) to be created:"
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.Panel7)
        Me.Panel5.Controls.Add(Me.Panel6)
        Me.Panel5.Location = New System.Drawing.Point(9, 360)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(545, 4)
        Me.Panel5.TabIndex = 147
        '
        'Panel7
        '
        Me.Panel7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel7.Controls.Add(Me.Panel8)
        Me.Panel7.Location = New System.Drawing.Point(-2, -2)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(545, 4)
        Me.Panel7.TabIndex = 148
        '
        'Panel8
        '
        Me.Panel8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel8.Location = New System.Drawing.Point(-2, 99)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(545, 4)
        Me.Panel8.TabIndex = 141
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Location = New System.Drawing.Point(-2, 99)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(545, 4)
        Me.Panel6.TabIndex = 141
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(165, 16)
        Me.Label1.TabIndex = 146
        Me.Label1.Text = "Select the Buffer Type:"
        '
        'rdbStandard
        '
        Me.rdbStandard.AutoSize = True
        Me.rdbStandard.Checked = True
        Me.rdbStandard.Location = New System.Drawing.Point(183, 11)
        Me.rdbStandard.Name = "rdbStandard"
        Me.rdbStandard.Size = New System.Drawing.Size(85, 17)
        Me.rdbStandard.TabIndex = 145
        Me.rdbStandard.TabStop = True
        Me.rdbStandard.Text = "STANDARD"
        Me.rdbStandard.UseVisualStyleBackColor = True
        '
        'rdbNetwork
        '
        Me.rdbNetwork.AutoSize = True
        Me.rdbNetwork.Location = New System.Drawing.Point(290, 11)
        Me.rdbNetwork.Name = "rdbNetwork"
        Me.rdbNetwork.Size = New System.Drawing.Size(81, 17)
        Me.rdbNetwork.TabIndex = 144
        Me.rdbNetwork.Text = "NETWORK"
        Me.rdbNetwork.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rdbNetwork.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Location = New System.Drawing.Point(9, 39)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(545, 4)
        Me.Panel4.TabIndex = 141
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(500, 302)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 48)
        Me.Button1.TabIndex = 163
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(7, 315)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(480, 28)
        Me.TextBox1.TabIndex = 161
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 296)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(179, 16)
        Me.Label3.TabIndex = 162
        Me.Label3.Text = "Select a Transit Network:"
        '
        'frmTravelSummaryBuffers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(571, 544)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Name = "frmTravelSummaryBuffers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create Network Service Area Buffers"
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel9.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnWorkspace As System.Windows.Forms.Button
    Friend WithEvents tbxWorkspace As System.Windows.Forms.TextBox
    Friend WithEvents chkUseSelected As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFieldsID As System.Windows.Forms.ComboBox
    Friend WithEvents lblFieldsId As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents barStatusRun As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents lblStNetowrk As System.Windows.Forms.Label
    Friend WithEvents btnNetworkLyr As System.Windows.Forms.Button
    Friend WithEvents tbxNetworkLyr As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rdbStandard As System.Windows.Forms.RadioButton
    Friend WithEvents rdbNetwork As System.Windows.Forms.RadioButton
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents btnCheckAll As System.Windows.Forms.Button
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents chkTransit As System.Windows.Forms.CheckBox
    Friend WithEvents chkOneMiAuto As System.Windows.Forms.CheckBox
    Friend WithEvents chkHalfMiAuto As System.Windows.Forms.CheckBox
    Friend WithEvents chkQtrMiAuto As System.Windows.Forms.CheckBox
    Friend WithEvents chk30min As System.Windows.Forms.CheckBox
    Friend WithEvents chk20min As System.Windows.Forms.CheckBox
    Friend WithEvents chk10min As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents btnUnselect As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
