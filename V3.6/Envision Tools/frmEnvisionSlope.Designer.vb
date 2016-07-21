<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEnvisionSlope
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblAnalysisCellSize = New System.Windows.Forms.Label
        Me.tbxCellSize = New System.Windows.Forms.TextBox
        Me.cmbGridLayers = New System.Windows.Forms.ComboBox
        Me.lblBndFeatClass = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnReturnGridCellSize = New System.Windows.Forms.Button
        Me.btnSlopeRun = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblAnalysisCellSize
        '
        Me.lblAnalysisCellSize.AutoSize = True
        Me.lblAnalysisCellSize.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAnalysisCellSize.Location = New System.Drawing.Point(7, 64)
        Me.lblAnalysisCellSize.Name = "lblAnalysisCellSize"
        Me.lblAnalysisCellSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAnalysisCellSize.Size = New System.Drawing.Size(150, 16)
        Me.lblAnalysisCellSize.TabIndex = 75
        Me.lblAnalysisCellSize.Text = "Output Grid Cell Size:"
        Me.lblAnalysisCellSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbxCellSize
        '
        Me.tbxCellSize.Location = New System.Drawing.Point(158, 63)
        Me.tbxCellSize.Name = "tbxCellSize"
        Me.tbxCellSize.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tbxCellSize.Size = New System.Drawing.Size(86, 20)
        Me.tbxCellSize.TabIndex = 74
        '
        'cmbGridLayers
        '
        Me.cmbGridLayers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbGridLayers.FormattingEnabled = True
        Me.cmbGridLayers.Location = New System.Drawing.Point(8, 29)
        Me.cmbGridLayers.Name = "cmbGridLayers"
        Me.cmbGridLayers.Size = New System.Drawing.Size(375, 21)
        Me.cmbGridLayers.TabIndex = 73
        Me.cmbGridLayers.Text = "Select a DEM Grid Layer"
        '
        'lblBndFeatClass
        '
        Me.lblBndFeatClass.AutoSize = True
        Me.lblBndFeatClass.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBndFeatClass.Location = New System.Drawing.Point(7, 8)
        Me.lblBndFeatClass.Name = "lblBndFeatClass"
        Me.lblBndFeatClass.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBndFeatClass.Size = New System.Drawing.Size(237, 16)
        Me.lblBndFeatClass.TabIndex = 72
        Me.lblBndFeatClass.Text = "Digital Elevation Model (DEM) Grid:"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.btnReturnGridCellSize)
        Me.Panel1.Controls.Add(Me.lblBndFeatClass)
        Me.Panel1.Controls.Add(Me.lblAnalysisCellSize)
        Me.Panel1.Controls.Add(Me.cmbGridLayers)
        Me.Panel1.Controls.Add(Me.tbxCellSize)
        Me.Panel1.Location = New System.Drawing.Point(2, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(390, 96)
        Me.Panel1.TabIndex = 76
        '
        'btnReturnGridCellSize
        '
        Me.btnReturnGridCellSize.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReturnGridCellSize.Location = New System.Drawing.Point(250, 58)
        Me.btnReturnGridCellSize.Name = "btnReturnGridCellSize"
        Me.btnReturnGridCellSize.Size = New System.Drawing.Size(133, 31)
        Me.btnReturnGridCellSize.TabIndex = 78
        Me.btnReturnGridCellSize.Text = "Get Grid Cell Size"
        Me.btnReturnGridCellSize.UseVisualStyleBackColor = True
        '
        'btnSlopeRun
        '
        Me.btnSlopeRun.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSlopeRun.Location = New System.Drawing.Point(294, 101)
        Me.btnSlopeRun.Name = "btnSlopeRun"
        Me.btnSlopeRun.Size = New System.Drawing.Size(98, 31)
        Me.btnSlopeRun.TabIndex = 77
        Me.btnSlopeRun.Text = "Run"
        Me.btnSlopeRun.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.Color.Red
        Me.lblStatus.Location = New System.Drawing.Point(12, 111)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(182, 13)
        Me.lblStatus.TabIndex = 78
        Me.lblStatus.Text = "Please Wait, Processing....."
        Me.lblStatus.Visible = False
        '
        'frmEnvisionSlope
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(394, 136)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnSlopeRun)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmEnvisionSlope"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generate Slope and Hillshade Layers"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblAnalysisCellSize As System.Windows.Forms.Label
    Friend WithEvents tbxCellSize As System.Windows.Forms.TextBox
    Friend WithEvents cmbGridLayers As System.Windows.Forms.ComboBox
    Friend WithEvents lblBndFeatClass As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnSlopeRun As System.Windows.Forms.Button
    Friend WithEvents btnReturnGridCellSize As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
End Class
