<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDevFeasibility
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDevFeasibility))
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbDevTypes = New System.Windows.Forms.ComboBox()
        Me.ToolStrip_InfoTab3 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRun = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.tbxLandValueAppreciation = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxTargetRent = New System.Windows.Forms.TextBox()
        Me.lblTarget = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbLandValueFld = New System.Windows.Forms.ComboBox()
        Me.cmbSqFtFld = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbImprovementValueFld = New System.Windows.Forms.ComboBox()
        Me.chkParcelSelected = New System.Windows.Forms.CheckBox()
        Me.ToolStrip_InfoTab3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(4, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(106, 16)
        Me.Label5.TabIndex = 141
        Me.Label5.Text = "Building Type: "
        '
        'cmbDevTypes
        '
        Me.cmbDevTypes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbDevTypes.FormattingEnabled = True
        Me.cmbDevTypes.Location = New System.Drawing.Point(116, 12)
        Me.cmbDevTypes.Name = "cmbDevTypes"
        Me.cmbDevTypes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbDevTypes.Size = New System.Drawing.Size(348, 21)
        Me.cmbDevTypes.TabIndex = 142
        Me.cmbDevTypes.Text = "<Select Building Type>"
        '
        'ToolStrip_InfoTab3
        '
        Me.ToolStrip_InfoTab3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip_InfoTab3.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip_InfoTab3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator11, Me.btnRun, Me.ToolStripSeparator9, Me.lblStatus})
        Me.ToolStrip_InfoTab3.Location = New System.Drawing.Point(0, 224)
        Me.ToolStrip_InfoTab3.Name = "ToolStrip_InfoTab3"
        Me.ToolStrip_InfoTab3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip_InfoTab3.Size = New System.Drawing.Size(476, 31)
        Me.ToolStrip_InfoTab3.TabIndex = 143
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
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 28)
        '
        'tbxLandValueAppreciation
        '
        Me.tbxLandValueAppreciation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxLandValueAppreciation.Location = New System.Drawing.Point(303, 66)
        Me.tbxLandValueAppreciation.Name = "tbxLandValueAppreciation"
        Me.tbxLandValueAppreciation.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxLandValueAppreciation.Size = New System.Drawing.Size(161, 20)
        Me.tbxLandValueAppreciation.TabIndex = 144
        Me.tbxLandValueAppreciation.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(213, 16)
        Me.Label1.TabIndex = 145
        Me.Label1.Text = "Land Value Appreciation :  (%)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxTargetRent
        '
        Me.tbxTargetRent.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxTargetRent.Location = New System.Drawing.Point(303, 97)
        Me.tbxTargetRent.Name = "tbxTargetRent"
        Me.tbxTargetRent.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxTargetRent.Size = New System.Drawing.Size(161, 20)
        Me.tbxTargetRent.TabIndex = 146
        Me.tbxTargetRent.Text = "0"
        '
        'lblTarget
        '
        Me.lblTarget.AutoSize = True
        Me.lblTarget.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTarget.Location = New System.Drawing.Point(4, 97)
        Me.lblTarget.Name = "lblTarget"
        Me.lblTarget.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTarget.Size = New System.Drawing.Size(263, 16)
        Me.lblTarget.TabIndex = 147
        Me.lblTarget.Text = "Target Achievable Rental Rate : ($/sf)"
        Me.lblTarget.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(4, 129)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(121, 16)
        Me.Label2.TabIndex = 149
        Me.Label2.Text = "Land Value Field:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(4, 161)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(176, 16)
        Me.Label3.TabIndex = 151
        Me.Label3.Text = "Improvement Value Field:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbLandValueFld
        '
        Me.cmbLandValueFld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLandValueFld.FormattingEnabled = True
        Me.cmbLandValueFld.Location = New System.Drawing.Point(209, 128)
        Me.cmbLandValueFld.Name = "cmbLandValueFld"
        Me.cmbLandValueFld.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbLandValueFld.Size = New System.Drawing.Size(255, 21)
        Me.cmbLandValueFld.TabIndex = 152
        '
        'cmbSqFtFld
        '
        Me.cmbSqFtFld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSqFtFld.FormattingEnabled = True
        Me.cmbSqFtFld.Location = New System.Drawing.Point(209, 192)
        Me.cmbSqFtFld.Name = "cmbSqFtFld"
        Me.cmbSqFtFld.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSqFtFld.Size = New System.Drawing.Size(255, 21)
        Me.cmbSqFtFld.TabIndex = 154
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(4, 193)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(199, 16)
        Me.Label4.TabIndex = 155
        Me.Label4.Text = "Parcel Square Footage Field:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbImprovementValueFld
        '
        Me.cmbImprovementValueFld.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbImprovementValueFld.FormattingEnabled = True
        Me.cmbImprovementValueFld.Location = New System.Drawing.Point(209, 160)
        Me.cmbImprovementValueFld.Name = "cmbImprovementValueFld"
        Me.cmbImprovementValueFld.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbImprovementValueFld.Size = New System.Drawing.Size(255, 21)
        Me.cmbImprovementValueFld.TabIndex = 153
        '
        'chkParcelSelected
        '
        Me.chkParcelSelected.AutoSize = True
        Me.chkParcelSelected.Enabled = False
        Me.chkParcelSelected.Location = New System.Drawing.Point(116, 38)
        Me.chkParcelSelected.Name = "chkParcelSelected"
        Me.chkParcelSelected.Size = New System.Drawing.Size(164, 17)
        Me.chkParcelSelected.TabIndex = 156
        Me.chkParcelSelected.Text = "Use Selected Feature(s) Only"
        Me.chkParcelSelected.UseVisualStyleBackColor = True
        '
        'frmDevFeasibility
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(476, 255)
        Me.Controls.Add(Me.chkParcelSelected)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmbSqFtFld)
        Me.Controls.Add(Me.cmbImprovementValueFld)
        Me.Controls.Add(Me.cmbLandValueFld)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbxLandValueAppreciation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbxTargetRent)
        Me.Controls.Add(Me.lblTarget)
        Me.Controls.Add(Me.ToolStrip_InfoTab3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cmbDevTypes)
        Me.Name = "frmDevFeasibility"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Development Feasibility"
        Me.ToolStrip_InfoTab3.ResumeLayout(False)
        Me.ToolStrip_InfoTab3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbDevTypes As System.Windows.Forms.ComboBox
    Friend WithEvents ToolStrip_InfoTab3 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tbxLandValueAppreciation As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbxTargetRent As System.Windows.Forms.TextBox
    Friend WithEvents lblTarget As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbLandValueFld As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSqFtFld As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbImprovementValueFld As System.Windows.Forms.ComboBox
    Friend WithEvents chkParcelSelected As System.Windows.Forms.CheckBox
End Class
