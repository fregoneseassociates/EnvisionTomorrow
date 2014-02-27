<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSumTransportationPnts
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSumTransportationPnts))
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.barStatus = New System.Windows.Forms.ToolStripProgressBar()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.lblEmp = New System.Windows.Forms.Label()
        Me.lblArea = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.cmbEmpFld = New System.Windows.Forms.ComboBox()
        Me.cmbArea = New System.Windows.Forms.ComboBox()
        Me.lblAreaFlds = New System.Windows.Forms.Label()
        Me.cmbFieldId = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.tbxWorkspace = New System.Windows.Forms.TextBox()
        Me.btnWorkspace = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbRailStops = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbTransitStops = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmb4WayIntersections = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbAllIntersections = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkUseSelected = New System.Windows.Forms.CheckBox()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.rdbExisting = New System.Windows.Forms.RadioButton()
        Me.rdbTotal = New System.Windows.Forms.RadioButton()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus, Me.barStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 443)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(577, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 58
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
        Me.Panel1.Controls.Add(Me.rdbTotal)
        Me.Panel1.Controls.Add(Me.rdbExisting)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.lblEmp)
        Me.Panel1.Controls.Add(Me.lblArea)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.cmbEmpFld)
        Me.Panel1.Controls.Add(Me.cmbArea)
        Me.Panel1.Controls.Add(Me.lblAreaFlds)
        Me.Panel1.Controls.Add(Me.cmbFieldId)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.tbxWorkspace)
        Me.Panel1.Controls.Add(Me.btnWorkspace)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.cmbRailStops)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.cmbTransitStops)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cmb4WayIntersections)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cmbAllIntersections)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.chkUseSelected)
        Me.Panel1.Controls.Add(Me.cmbLayers)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(576, 439)
        Me.Panel1.TabIndex = 59
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Location = New System.Drawing.Point(5, 286)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(563, 4)
        Me.Panel5.TabIndex = 186
        '
        'lblEmp
        '
        Me.lblEmp.AutoSize = True
        Me.lblEmp.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmp.Location = New System.Drawing.Point(241, 174)
        Me.lblEmp.Name = "lblEmp"
        Me.lblEmp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmp.Size = New System.Drawing.Size(125, 16)
        Me.lblEmp.TabIndex = 182
        Me.lblEmp.Text = "Employment field:"
        '
        'lblArea
        '
        Me.lblArea.AutoSize = True
        Me.lblArea.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArea.Location = New System.Drawing.Point(3, 174)
        Me.lblArea.Name = "lblArea"
        Me.lblArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblArea.Size = New System.Drawing.Size(76, 16)
        Me.lblArea.TabIndex = 179
        Me.lblArea.Text = "Area field:"
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Location = New System.Drawing.Point(4, 135)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(563, 4)
        Me.Panel4.TabIndex = 171
        '
        'cmbEmpFld
        '
        Me.cmbEmpFld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbEmpFld.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbEmpFld.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbEmpFld.FormattingEnabled = True
        Me.cmbEmpFld.Location = New System.Drawing.Point(370, 173)
        Me.cmbEmpFld.Name = "cmbEmpFld"
        Me.cmbEmpFld.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbEmpFld.Size = New System.Drawing.Size(190, 21)
        Me.cmbEmpFld.TabIndex = 7
        Me.cmbEmpFld.Text = "Select a field"
        '
        'cmbArea
        '
        Me.cmbArea.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbArea.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbArea.FormattingEnabled = True
        Me.cmbArea.Location = New System.Drawing.Point(85, 173)
        Me.cmbArea.Name = "cmbArea"
        Me.cmbArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbArea.Size = New System.Drawing.Size(142, 21)
        Me.cmbArea.TabIndex = 5
        Me.cmbArea.Text = "Select a field"
        '
        'lblAreaFlds
        '
        Me.lblAreaFlds.AutoSize = True
        Me.lblAreaFlds.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAreaFlds.Location = New System.Drawing.Point(2, 147)
        Me.lblAreaFlds.Name = "lblAreaFlds"
        Me.lblAreaFlds.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAreaFlds.Size = New System.Drawing.Size(162, 16)
        Me.lblAreaFlds.TabIndex = 173
        Me.lblAreaFlds.Text = "Select the Input fields:"
        '
        'cmbFieldId
        '
        Me.cmbFieldId.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbFieldId.Enabled = False
        Me.cmbFieldId.FormattingEnabled = True
        Me.cmbFieldId.Location = New System.Drawing.Point(202, 108)
        Me.cmbFieldId.Name = "cmbFieldId"
        Me.cmbFieldId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbFieldId.Size = New System.Drawing.Size(358, 21)
        Me.cmbFieldId.TabIndex = 3
        Me.cmbFieldId.Text = "Select a field"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Enabled = False
        Me.Label7.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(1, 108)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(182, 16)
        Me.Label7.TabIndex = 171
        Me.Label7.Text = "Select the unique ID field:"
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(4, 208)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(563, 4)
        Me.Panel3.TabIndex = 170
        '
        'tbxWorkspace
        '
        Me.tbxWorkspace.AllowDrop = True
        Me.tbxWorkspace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxWorkspace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxWorkspace.Location = New System.Drawing.Point(14, 243)
        Me.tbxWorkspace.Multiline = True
        Me.tbxWorkspace.Name = "tbxWorkspace"
        Me.tbxWorkspace.ReadOnly = True
        Me.tbxWorkspace.Size = New System.Drawing.Size(484, 28)
        Me.tbxWorkspace.TabIndex = 9
        '
        'btnWorkspace
        '
        Me.btnWorkspace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnWorkspace.BackColor = System.Drawing.Color.Transparent
        Me.btnWorkspace.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnWorkspace.Image = CType(resources.GetObject("btnWorkspace.Image"), System.Drawing.Image)
        Me.btnWorkspace.Location = New System.Drawing.Point(511, 230)
        Me.btnWorkspace.Name = "btnWorkspace"
        Me.btnWorkspace.Size = New System.Drawing.Size(56, 48)
        Me.btnWorkspace.TabIndex = 10
        Me.btnWorkspace.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(14, 223)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(380, 16)
        Me.Label6.TabIndex = 167
        Me.Label6.Text = "Select a Network or Standard Buffers File Geodatabase:"
        '
        'cmbRailStops
        '
        Me.cmbRailStops.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbRailStops.FormattingEnabled = True
        Me.cmbRailStops.Location = New System.Drawing.Point(293, 404)
        Me.cmbRailStops.Name = "cmbRailStops"
        Me.cmbRailStops.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbRailStops.Size = New System.Drawing.Size(267, 21)
        Me.cmbRailStops.TabIndex = 13
        Me.cmbRailStops.Text = "Select a layer"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(1, 404)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(260, 16)
        Me.Label4.TabIndex = 164
        Me.Label4.Text = "Rail Transportation Stops Point Layer:"
        '
        'cmbTransitStops
        '
        Me.cmbTransitStops.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbTransitStops.FormattingEnabled = True
        Me.cmbTransitStops.Location = New System.Drawing.Point(195, 370)
        Me.cmbTransitStops.Name = "cmbTransitStops"
        Me.cmbTransitStops.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbTransitStops.Size = New System.Drawing.Size(365, 21)
        Me.cmbTransitStops.TabIndex = 12
        Me.cmbTransitStops.Text = "Select a layer"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(1, 370)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(175, 16)
        Me.Label3.TabIndex = 162
        Me.Label3.Text = "Transit Stop Point Layer:"
        '
        'cmb4WayIntersections
        '
        Me.cmb4WayIntersections.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmb4WayIntersections.FormattingEnabled = True
        Me.cmb4WayIntersections.Location = New System.Drawing.Point(403, 337)
        Me.cmb4WayIntersections.Name = "cmb4WayIntersections"
        Me.cmb4WayIntersections.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmb4WayIntersections.Size = New System.Drawing.Size(157, 21)
        Me.cmb4WayIntersections.TabIndex = 14
        Me.cmb4WayIntersections.Text = "Select a layer"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(1, 337)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(383, 16)
        Me.Label2.TabIndex = 160
        Me.Label2.Text = "Transportation Network 4-Way Intersection Point Layer:"
        '
        'cmbAllIntersections
        '
        Me.cmbAllIntersections.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAllIntersections.FormattingEnabled = True
        Me.cmbAllIntersections.Location = New System.Drawing.Point(354, 303)
        Me.cmbAllIntersections.Name = "cmbAllIntersections"
        Me.cmbAllIntersections.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAllIntersections.Size = New System.Drawing.Size(206, 21)
        Me.cmbAllIntersections.TabIndex = 11
        Me.cmbAllIntersections.Text = "Select a layer"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(1, 303)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(334, 16)
        Me.Label1.TabIndex = 158
        Me.Label1.Text = "Transportation Network Intersection Point Layer:"
        '
        'chkUseSelected
        '
        Me.chkUseSelected.AutoSize = True
        Me.chkUseSelected.Location = New System.Drawing.Point(177, 56)
        Me.chkUseSelected.Name = "chkUseSelected"
        Me.chkUseSelected.Size = New System.Drawing.Size(189, 17)
        Me.chkUseSelected.TabIndex = 1
        Me.chkUseSelected.Text = "Only Calculate Selected Feature(s)"
        Me.chkUseSelected.UseVisualStyleBackColor = True
        Me.chkUseSelected.Visible = False
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(17, 76)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLayers.Size = New System.Drawing.Size(543, 21)
        Me.cmbLayers.TabIndex = 2
        Me.cmbLayers.Text = "Select a layer"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(14, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(144, 16)
        Me.Label5.TabIndex = 152
        Me.Label5.Text = "Neighborhood Layer:"
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Location = New System.Drawing.Point(5, 39)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(563, 4)
        Me.Panel6.TabIndex = 187
        '
        'rdbExisting
        '
        Me.rdbExisting.AutoSize = True
        Me.rdbExisting.Location = New System.Drawing.Point(167, 9)
        Me.rdbExisting.Name = "rdbExisting"
        Me.rdbExisting.Size = New System.Drawing.Size(108, 17)
        Me.rdbExisting.TabIndex = 188
        Me.rdbExisting.Text = "Calculate Existing"
        Me.rdbExisting.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rdbExisting.UseVisualStyleBackColor = True
        '
        'rdbTotal
        '
        Me.rdbTotal.AutoSize = True
        Me.rdbTotal.Checked = True
        Me.rdbTotal.Location = New System.Drawing.Point(14, 9)
        Me.rdbTotal.Name = "rdbTotal"
        Me.rdbTotal.Size = New System.Drawing.Size(141, 17)
        Me.rdbTotal.TabIndex = 189
        Me.rdbTotal.TabStop = True
        Me.rdbTotal.Text = "Calculate Scenario Total"
        Me.rdbTotal.UseVisualStyleBackColor = True
        '
        'frmSumTransportationPnts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(577, 474)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Name = "frmSumTransportationPnts"
        Me.Text = "Transportation Locations Summary"
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
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
    Friend WithEvents cmbRailStops As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbTransitStops As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmb4WayIntersections As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbAllIntersections As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkUseSelected As System.Windows.Forms.CheckBox
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbxWorkspace As System.Windows.Forms.TextBox
    Friend WithEvents btnWorkspace As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents cmbFieldId As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmbArea As System.Windows.Forms.ComboBox
    Friend WithEvents lblAreaFlds As System.Windows.Forms.Label
    Friend WithEvents cmbEmpFld As System.Windows.Forms.ComboBox
    Friend WithEvents lblEmp As System.Windows.Forms.Label
    Friend WithEvents lblArea As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents rdbTotal As System.Windows.Forms.RadioButton
    Friend WithEvents rdbExisting As System.Windows.Forms.RadioButton
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
End Class
