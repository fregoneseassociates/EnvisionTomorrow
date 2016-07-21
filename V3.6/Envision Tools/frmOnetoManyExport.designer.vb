<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOnetoManyExport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOnetoManyExport))
        Me.lblLstFields = New System.Windows.Forms.Label
        Me.lstFields = New System.Windows.Forms.CheckedListBox
        Me.lbl1toManyFields = New System.Windows.Forms.Label
        Me.cmb1toManyFields = New System.Windows.Forms.ComboBox
        Me.tbxCSVFile = New System.Windows.Forms.TextBox
        Me.btnSaveCSVFile = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.tbxExportTable = New System.Windows.Forms.TextBox
        Me.lblExportTable = New System.Windows.Forms.Label
        Me.btnExport = New System.Windows.Forms.Button
        Me.btnCheckAll = New System.Windows.Forms.Button
        Me.btnUncheckAll = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.grpCSV = New System.Windows.Forms.GroupBox
        Me.grpDBF = New System.Windows.Forms.GroupBox
        Me.tbxDBFFile = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnDBFSaveFile = New System.Windows.Forms.Button
        Me.rdbCSV = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.rdbDBF = New System.Windows.Forms.RadioButton
        Me.rdbTable = New System.Windows.Forms.RadioButton
        Me.grpEnvisionTable = New System.Windows.Forms.GroupBox
        Me.grpCSV.SuspendLayout()
        Me.grpDBF.SuspendLayout()
        Me.grpEnvisionTable.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblLstFields
        '
        Me.lblLstFields.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLstFields.AutoSize = True
        Me.lblLstFields.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLstFields.Location = New System.Drawing.Point(5, 45)
        Me.lblLstFields.Name = "lblLstFields"
        Me.lblLstFields.Size = New System.Drawing.Size(184, 14)
        Me.lblLstFields.TabIndex = 63
        Me.lblLstFields.Text = "Select the Attribute Field(s):"
        Me.lblLstFields.UseWaitCursor = True
        '
        'lstFields
        '
        Me.lstFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFields.FormattingEnabled = True
        Me.lstFields.Location = New System.Drawing.Point(5, 67)
        Me.lstFields.MultiColumn = True
        Me.lstFields.Name = "lstFields"
        Me.lstFields.Size = New System.Drawing.Size(514, 139)
        Me.lstFields.TabIndex = 62
        Me.lstFields.UseWaitCursor = True
        '
        'lbl1toManyFields
        '
        Me.lbl1toManyFields.AutoSize = True
        Me.lbl1toManyFields.Enabled = False
        Me.lbl1toManyFields.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1toManyFields.Location = New System.Drawing.Point(5, 13)
        Me.lbl1toManyFields.Name = "lbl1toManyFields"
        Me.lbl1toManyFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl1toManyFields.Size = New System.Drawing.Size(225, 16)
        Me.lbl1toManyFields.TabIndex = 98
        Me.lbl1toManyFields.Text = "Select One to Many Count Field:"
        Me.lbl1toManyFields.UseWaitCursor = True
        '
        'cmb1toManyFields
        '
        Me.cmb1toManyFields.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmb1toManyFields.Enabled = False
        Me.cmb1toManyFields.FormattingEnabled = True
        Me.cmb1toManyFields.Location = New System.Drawing.Point(236, 12)
        Me.cmb1toManyFields.Name = "cmb1toManyFields"
        Me.cmb1toManyFields.Size = New System.Drawing.Size(283, 21)
        Me.cmb1toManyFields.TabIndex = 99
        Me.cmb1toManyFields.UseWaitCursor = True
        '
        'tbxCSVFile
        '
        Me.tbxCSVFile.AllowDrop = True
        Me.tbxCSVFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxCSVFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxCSVFile.Location = New System.Drawing.Point(6, 42)
        Me.tbxCSVFile.Multiline = True
        Me.tbxCSVFile.Name = "tbxCSVFile"
        Me.tbxCSVFile.ReadOnly = True
        Me.tbxCSVFile.Size = New System.Drawing.Size(444, 28)
        Me.tbxCSVFile.TabIndex = 100
        Me.tbxCSVFile.UseWaitCursor = True
        '
        'btnSaveCSVFile
        '
        Me.btnSaveCSVFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSaveCSVFile.BackColor = System.Drawing.Color.Transparent
        Me.btnSaveCSVFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSaveCSVFile.Image = CType(resources.GetObject("btnSaveCSVFile.Image"), System.Drawing.Image)
        Me.btnSaveCSVFile.Location = New System.Drawing.Point(456, 22)
        Me.btnSaveCSVFile.Name = "btnSaveCSVFile"
        Me.btnSaveCSVFile.Size = New System.Drawing.Size(56, 48)
        Me.btnSaveCSVFile.TabIndex = 102
        Me.btnSaveCSVFile.UseVisualStyleBackColor = False
        Me.btnSaveCSVFile.UseWaitCursor = True
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(3, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(122, 16)
        Me.Label6.TabIndex = 101
        Me.Label6.Text = "Output Filename:"
        Me.Label6.UseWaitCursor = True
        '
        'tbxExportTable
        '
        Me.tbxExportTable.AllowDrop = True
        Me.tbxExportTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxExportTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxExportTable.Location = New System.Drawing.Point(98, 42)
        Me.tbxExportTable.Multiline = True
        Me.tbxExportTable.Name = "tbxExportTable"
        Me.tbxExportTable.Size = New System.Drawing.Size(410, 28)
        Me.tbxExportTable.TabIndex = 108
        Me.tbxExportTable.UseWaitCursor = True
        '
        'lblExportTable
        '
        Me.lblExportTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblExportTable.AutoSize = True
        Me.lblExportTable.Enabled = False
        Me.lblExportTable.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExportTable.Location = New System.Drawing.Point(4, 45)
        Me.lblExportTable.Name = "lblExportTable"
        Me.lblExportTable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExportTable.Size = New System.Drawing.Size(91, 16)
        Me.lblExportTable.TabIndex = 107
        Me.lblExportTable.Text = "Table Name:"
        Me.lblExportTable.UseWaitCursor = True
        '
        'btnExport
        '
        Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExport.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.Location = New System.Drawing.Point(390, 354)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(129, 35)
        Me.btnExport.TabIndex = 109
        Me.btnExport.Text = "Export"
        Me.btnExport.UseVisualStyleBackColor = True
        Me.btnExport.UseWaitCursor = True
        '
        'btnCheckAll
        '
        Me.btnCheckAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckAll.Location = New System.Drawing.Point(203, 40)
        Me.btnCheckAll.Name = "btnCheckAll"
        Me.btnCheckAll.Size = New System.Drawing.Size(129, 24)
        Me.btnCheckAll.TabIndex = 115
        Me.btnCheckAll.Text = "Check All"
        Me.btnCheckAll.UseVisualStyleBackColor = True
        Me.btnCheckAll.UseWaitCursor = True
        '
        'btnUncheckAll
        '
        Me.btnUncheckAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUncheckAll.Location = New System.Drawing.Point(342, 40)
        Me.btnUncheckAll.Name = "btnUncheckAll"
        Me.btnUncheckAll.Size = New System.Drawing.Size(129, 24)
        Me.btnUncheckAll.TabIndex = 116
        Me.btnUncheckAll.Text = "Uncheck All"
        Me.btnUncheckAll.UseVisualStyleBackColor = True
        Me.btnUncheckAll.UseWaitCursor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(5, 355)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(375, 34)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 117
        Me.ProgressBar1.UseWaitCursor = True
        Me.ProgressBar1.Visible = False
        '
        'grpCSV
        '
        Me.grpCSV.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpCSV.Controls.Add(Me.tbxCSVFile)
        Me.grpCSV.Controls.Add(Me.Label6)
        Me.grpCSV.Controls.Add(Me.btnSaveCSVFile)
        Me.grpCSV.Location = New System.Drawing.Point(5, 263)
        Me.grpCSV.Name = "grpCSV"
        Me.grpCSV.Size = New System.Drawing.Size(519, 76)
        Me.grpCSV.TabIndex = 118
        Me.grpCSV.TabStop = False
        Me.grpCSV.Text = "CSV File:"
        Me.grpCSV.UseWaitCursor = True
        '
        'grpDBF
        '
        Me.grpDBF.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpDBF.Controls.Add(Me.tbxDBFFile)
        Me.grpDBF.Controls.Add(Me.Label4)
        Me.grpDBF.Controls.Add(Me.btnDBFSaveFile)
        Me.grpDBF.Location = New System.Drawing.Point(5, 264)
        Me.grpDBF.Name = "grpDBF"
        Me.grpDBF.Size = New System.Drawing.Size(518, 76)
        Me.grpDBF.TabIndex = 119
        Me.grpDBF.TabStop = False
        Me.grpDBF.Text = "dBASE File"
        Me.grpDBF.UseWaitCursor = True
        '
        'tbxDBFFile
        '
        Me.tbxDBFFile.AllowDrop = True
        Me.tbxDBFFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxDBFFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDBFFile.Location = New System.Drawing.Point(6, 42)
        Me.tbxDBFFile.Multiline = True
        Me.tbxDBFFile.Name = "tbxDBFFile"
        Me.tbxDBFFile.ReadOnly = True
        Me.tbxDBFFile.Size = New System.Drawing.Size(443, 28)
        Me.tbxDBFFile.TabIndex = 103
        Me.tbxDBFFile.UseWaitCursor = True
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(122, 16)
        Me.Label4.TabIndex = 104
        Me.Label4.Text = "Output Filename:"
        Me.Label4.UseWaitCursor = True
        '
        'btnDBFSaveFile
        '
        Me.btnDBFSaveFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDBFSaveFile.BackColor = System.Drawing.Color.Transparent
        Me.btnDBFSaveFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDBFSaveFile.Image = CType(resources.GetObject("btnDBFSaveFile.Image"), System.Drawing.Image)
        Me.btnDBFSaveFile.Location = New System.Drawing.Point(455, 22)
        Me.btnDBFSaveFile.Name = "btnDBFSaveFile"
        Me.btnDBFSaveFile.Size = New System.Drawing.Size(56, 48)
        Me.btnDBFSaveFile.TabIndex = 105
        Me.btnDBFSaveFile.UseVisualStyleBackColor = False
        Me.btnDBFSaveFile.UseWaitCursor = True
        '
        'rdbCSV
        '
        Me.rdbCSV.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.rdbCSV.AutoSize = True
        Me.rdbCSV.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbCSV.Location = New System.Drawing.Point(250, 238)
        Me.rdbCSV.Name = "rdbCSV"
        Me.rdbCSV.Size = New System.Drawing.Size(199, 20)
        Me.rdbCSV.TabIndex = 120
        Me.rdbCSV.Text = "CSV (Deliminted Text File)"
        Me.rdbCSV.UseVisualStyleBackColor = True
        Me.rdbCSV.UseWaitCursor = True
        Me.rdbCSV.Visible = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.Label1.Location = New System.Drawing.Point(2, 221)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 16)
        Me.Label1.TabIndex = 121
        Me.Label1.Text = "Output Format:"
        Me.Label1.UseWaitCursor = True
        '
        'rdbDBF
        '
        Me.rdbDBF.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.rdbDBF.AutoSize = True
        Me.rdbDBF.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbDBF.Location = New System.Drawing.Point(5, 238)
        Me.rdbDBF.Name = "rdbDBF"
        Me.rdbDBF.Size = New System.Drawing.Size(107, 20)
        Me.rdbDBF.TabIndex = 122
        Me.rdbDBF.Text = "dBase (DBF)"
        Me.rdbDBF.UseVisualStyleBackColor = True
        Me.rdbDBF.UseWaitCursor = True
        '
        'rdbTable
        '
        Me.rdbTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.rdbTable.AutoSize = True
        Me.rdbTable.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.rdbTable.Location = New System.Drawing.Point(121, 238)
        Me.rdbTable.Name = "rdbTable"
        Me.rdbTable.Size = New System.Drawing.Size(120, 20)
        Me.rdbTable.TabIndex = 123
        Me.rdbTable.Text = "Envision Table"
        Me.rdbTable.UseVisualStyleBackColor = True
        Me.rdbTable.UseWaitCursor = True
        '
        'grpEnvisionTable
        '
        Me.grpEnvisionTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpEnvisionTable.Controls.Add(Me.tbxExportTable)
        Me.grpEnvisionTable.Controls.Add(Me.lblExportTable)
        Me.grpEnvisionTable.Location = New System.Drawing.Point(5, 263)
        Me.grpEnvisionTable.Name = "grpEnvisionTable"
        Me.grpEnvisionTable.Size = New System.Drawing.Size(518, 76)
        Me.grpEnvisionTable.TabIndex = 120
        Me.grpEnvisionTable.TabStop = False
        Me.grpEnvisionTable.Text = "Envision Geodatabase Table"
        Me.grpEnvisionTable.UseWaitCursor = True
        Me.grpEnvisionTable.Visible = False
        '
        'frmOnetoManyExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 393)
        Me.Controls.Add(Me.rdbTable)
        Me.Controls.Add(Me.rdbDBF)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.grpDBF)
        Me.Controls.Add(Me.rdbCSV)
        Me.Controls.Add(Me.grpCSV)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnUncheckAll)
        Me.Controls.Add(Me.btnCheckAll)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.lbl1toManyFields)
        Me.Controls.Add(Me.cmb1toManyFields)
        Me.Controls.Add(Me.lblLstFields)
        Me.Controls.Add(Me.lstFields)
        Me.Controls.Add(Me.grpEnvisionTable)
        Me.Name = "frmOnetoManyExport"
        Me.ShowInTaskbar = False
        Me.Text = "One to Many Record Export"
        Me.TopMost = True
        Me.UseWaitCursor = True
        Me.grpCSV.ResumeLayout(False)
        Me.grpCSV.PerformLayout()
        Me.grpDBF.ResumeLayout(False)
        Me.grpDBF.PerformLayout()
        Me.grpEnvisionTable.ResumeLayout(False)
        Me.grpEnvisionTable.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblLstFields As System.Windows.Forms.Label
    Friend WithEvents lstFields As System.Windows.Forms.CheckedListBox
    Friend WithEvents lbl1toManyFields As System.Windows.Forms.Label
    Friend WithEvents cmb1toManyFields As System.Windows.Forms.ComboBox
    Friend WithEvents tbxCSVFile As System.Windows.Forms.TextBox
    Friend WithEvents btnSaveCSVFile As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbxExportTable As System.Windows.Forms.TextBox
    Friend WithEvents lblExportTable As System.Windows.Forms.Label
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents btnCheckAll As System.Windows.Forms.Button
    Friend WithEvents btnUncheckAll As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents grpCSV As System.Windows.Forms.GroupBox
    Friend WithEvents grpDBF As System.Windows.Forms.GroupBox
    Friend WithEvents rdbCSV As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rdbDBF As System.Windows.Forms.RadioButton
    Friend WithEvents rdbTable As System.Windows.Forms.RadioButton
    Friend WithEvents tbxDBFFile As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnDBFSaveFile As System.Windows.Forms.Button
    Friend WithEvents grpEnvisionTable As System.Windows.Forms.GroupBox
End Class
