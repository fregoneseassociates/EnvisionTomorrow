<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmJobSummary
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmJobSummary))
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnWorkspace = New System.Windows.Forms.Button()
        Me.tbxWorkspace = New System.Windows.Forms.TextBox()
        Me.chkUseSelected = New System.Windows.Forms.CheckBox()
        Me.cmbFieldsID = New System.Windows.Forms.ComboBox()
        Me.lblFieldsId = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmbLayers = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 182)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(560, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 56
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
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.btnWorkspace)
        Me.Panel1.Controls.Add(Me.tbxWorkspace)
        Me.Panel1.Controls.Add(Me.chkUseSelected)
        Me.Panel1.Controls.Add(Me.cmbFieldsID)
        Me.Panel1.Controls.Add(Me.lblFieldsId)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.cmbLayers)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Location = New System.Drawing.Point(1, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(559, 178)
        Me.Panel1.TabIndex = 55
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(4, 105)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(373, 16)
        Me.Label6.TabIndex = 148
        Me.Label6.Text = "Select a Standard or Network Buffer File Geodatabase:"
        '
        'btnWorkspace
        '
        Me.btnWorkspace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnWorkspace.BackColor = System.Drawing.Color.Transparent
        Me.btnWorkspace.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnWorkspace.Image = CType(resources.GetObject("btnWorkspace.Image"), System.Drawing.Image)
        Me.btnWorkspace.Location = New System.Drawing.Point(487, 120)
        Me.btnWorkspace.Name = "btnWorkspace"
        Me.btnWorkspace.Size = New System.Drawing.Size(56, 48)
        Me.btnWorkspace.TabIndex = 149
        Me.btnWorkspace.UseVisualStyleBackColor = False
        '
        'tbxWorkspace
        '
        Me.tbxWorkspace.AllowDrop = True
        Me.tbxWorkspace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxWorkspace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxWorkspace.Location = New System.Drawing.Point(4, 133)
        Me.tbxWorkspace.Multiline = True
        Me.tbxWorkspace.Name = "tbxWorkspace"
        Me.tbxWorkspace.ReadOnly = True
        Me.tbxWorkspace.Size = New System.Drawing.Size(471, 28)
        Me.tbxWorkspace.TabIndex = 147
        '
        'chkUseSelected
        '
        Me.chkUseSelected.AutoSize = True
        Me.chkUseSelected.Enabled = False
        Me.chkUseSelected.Location = New System.Drawing.Point(230, 34)
        Me.chkUseSelected.Name = "chkUseSelected"
        Me.chkUseSelected.Size = New System.Drawing.Size(140, 17)
        Me.chkUseSelected.TabIndex = 157
        Me.chkUseSelected.Text = "Use Selected Feature(s)"
        Me.chkUseSelected.UseVisualStyleBackColor = True
        '
        'cmbFieldsID
        '
        Me.cmbFieldsID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbFieldsID.Enabled = False
        Me.cmbFieldsID.FormattingEnabled = True
        Me.cmbFieldsID.Location = New System.Drawing.Point(227, 62)
        Me.cmbFieldsID.Name = "cmbFieldsID"
        Me.cmbFieldsID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbFieldsID.Size = New System.Drawing.Size(316, 21)
        Me.cmbFieldsID.TabIndex = 156
        '
        'lblFieldsId
        '
        Me.lblFieldsId.AutoSize = True
        Me.lblFieldsId.Enabled = False
        Me.lblFieldsId.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldsId.Location = New System.Drawing.Point(49, 63)
        Me.lblFieldsId.Name = "lblFieldsId"
        Me.lblFieldsId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFieldsId.Size = New System.Drawing.Size(175, 16)
        Me.lblFieldsId.TabIndex = 155
        Me.lblFieldsId.Text = "Select a UNIQUE ID field:"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Location = New System.Drawing.Point(7, 94)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(543, 4)
        Me.Panel2.TabIndex = 153
        '
        'cmbLayers
        '
        Me.cmbLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLayers.FormattingEnabled = True
        Me.cmbLayers.Location = New System.Drawing.Point(227, 5)
        Me.cmbLayers.Name = "cmbLayers"
        Me.cmbLayers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLayers.Size = New System.Drawing.Size(316, 21)
        Me.cmbLayers.TabIndex = 154
        Me.cmbLayers.Text = "Select a layer"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(10, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(214, 16)
        Me.Label5.TabIndex = 152
        Me.Label5.Text = "Input Layer (Point or Polygon):"
        '
        'frmJobSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 213)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmJobSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sum Jobs"
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnWorkspace As System.Windows.Forms.Button
    Friend WithEvents tbxWorkspace As System.Windows.Forms.TextBox
    Friend WithEvents chkUseSelected As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFieldsID As System.Windows.Forms.ComboBox
    Friend WithEvents lblFieldsId As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmbLayers As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
End Class
