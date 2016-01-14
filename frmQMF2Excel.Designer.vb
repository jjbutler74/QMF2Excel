<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQMF2Excel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQMF2Excel))
        Me.btnSubmit = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.btnExcelFolder = New System.Windows.Forms.Button
        Me.txtExcelFolder = New System.Windows.Forms.TextBox
        Me.lbl3 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.lbl2 = New System.Windows.Forms.Label
        Me.txtMainframeFile = New System.Windows.Forms.TextBox
        Me.txtExcelFile = New System.Windows.Forms.TextBox
        Me.lbl4 = New System.Windows.Forms.Label
        Me.lstLog = New System.Windows.Forms.ListBox
        Me.chkNumeric = New System.Windows.Forms.CheckBox
        Me.chkRememberScreen = New System.Windows.Forms.CheckBox
        Me.cmbMode = New System.Windows.Forms.ComboBox
        Me.lbl1 = New System.Windows.Forms.Label
        Me.txtBatchFile = New System.Windows.Forms.TextBox
        Me.btnBatchFile = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.chkOverwrite = New System.Windows.Forms.CheckBox
        Me.lbl5 = New System.Windows.Forms.Label
        Me.txtMainframeHost = New System.Windows.Forms.TextBox
        Me.lbl6 = New System.Windows.Forms.Label
        Me.txtMainframeUserId = New System.Windows.Forms.TextBox
        Me.lbl7 = New System.Windows.Forms.Label
        Me.txtMainframePassword = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnSubmit
        '
        Me.btnSubmit.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSubmit.Location = New System.Drawing.Point(247, 438)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(97, 23)
        Me.btnSubmit.TabIndex = 15
        Me.btnSubmit.Text = "&Submit"
        Me.btnSubmit.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(397, 438)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(95, 23)
        Me.btnExit.TabIndex = 16
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnExcelFolder
        '
        Me.btnExcelFolder.Location = New System.Drawing.Point(655, 168)
        Me.btnExcelFolder.Name = "btnExcelFolder"
        Me.btnExcelFolder.Size = New System.Drawing.Size(48, 23)
        Me.btnExcelFolder.TabIndex = 9
        Me.btnExcelFolder.Text = "Select"
        Me.btnExcelFolder.UseVisualStyleBackColor = True
        '
        'txtExcelFolder
        '
        Me.txtExcelFolder.BackColor = System.Drawing.Color.White
        Me.txtExcelFolder.Location = New System.Drawing.Point(121, 170)
        Me.txtExcelFolder.Name = "txtExcelFolder"
        Me.txtExcelFolder.ReadOnly = True
        Me.txtExcelFolder.Size = New System.Drawing.Size(528, 20)
        Me.txtExcelFolder.TabIndex = 8
        Me.txtExcelFolder.TabStop = False
        '
        'lbl3
        '
        Me.lbl3.AutoSize = True
        Me.lbl3.Location = New System.Drawing.Point(40, 170)
        Me.lbl3.Name = "lbl3"
        Me.lbl3.Size = New System.Drawing.Size(65, 13)
        Me.lbl3.TabIndex = 4
        Me.lbl3.Text = "Excel Folder"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.Location = New System.Drawing.Point(40, 137)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(75, 13)
        Me.lbl2.TabIndex = 7
        Me.lbl2.Text = "Mainframe File"
        '
        'txtMainframeFile
        '
        Me.txtMainframeFile.BackColor = System.Drawing.Color.White
        Me.txtMainframeFile.Location = New System.Drawing.Point(121, 134)
        Me.txtMainframeFile.MaxLength = 500
        Me.txtMainframeFile.Name = "txtMainframeFile"
        Me.txtMainframeFile.Size = New System.Drawing.Size(196, 20)
        Me.txtMainframeFile.TabIndex = 7
        '
        'txtExcelFile
        '
        Me.txtExcelFile.Location = New System.Drawing.Point(121, 206)
        Me.txtExcelFile.MaxLength = 500
        Me.txtExcelFile.Name = "txtExcelFile"
        Me.txtExcelFile.Size = New System.Drawing.Size(196, 20)
        Me.txtExcelFile.TabIndex = 10
        '
        'lbl4
        '
        Me.lbl4.AutoSize = True
        Me.lbl4.Location = New System.Drawing.Point(40, 209)
        Me.lbl4.Name = "lbl4"
        Me.lbl4.Size = New System.Drawing.Size(52, 13)
        Me.lbl4.TabIndex = 9
        Me.lbl4.Text = "Excel File"
        '
        'lstLog
        '
        Me.lstLog.FormattingEnabled = True
        Me.lstLog.Location = New System.Drawing.Point(43, 276)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.Size = New System.Drawing.Size(660, 147)
        Me.lstLog.TabIndex = 14
        Me.lstLog.TabStop = False
        '
        'chkNumeric
        '
        Me.chkNumeric.AutoSize = True
        Me.chkNumeric.Checked = True
        Me.chkNumeric.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkNumeric.Location = New System.Drawing.Point(38, 243)
        Me.chkNumeric.Name = "chkNumeric"
        Me.chkNumeric.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkNumeric.Size = New System.Drawing.Size(100, 17)
        Me.chkNumeric.TabIndex = 11
        Me.chkNumeric.Text = "Auto Formatting"
        Me.chkNumeric.UseVisualStyleBackColor = True
        '
        'chkRememberScreen
        '
        Me.chkRememberScreen.AutoSize = True
        Me.chkRememberScreen.Checked = True
        Me.chkRememberScreen.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRememberScreen.Location = New System.Drawing.Point(553, 243)
        Me.chkRememberScreen.Name = "chkRememberScreen"
        Me.chkRememberScreen.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkRememberScreen.Size = New System.Drawing.Size(149, 17)
        Me.chkRememberScreen.TabIndex = 13
        Me.chkRememberScreen.Text = "Remember Screen Values"
        Me.chkRememberScreen.UseVisualStyleBackColor = True
        '
        'cmbMode
        '
        Me.cmbMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMode.FormattingEnabled = True
        Me.cmbMode.Location = New System.Drawing.Point(121, 61)
        Me.cmbMode.Name = "cmbMode"
        Me.cmbMode.Size = New System.Drawing.Size(121, 21)
        Me.cmbMode.TabIndex = 1
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(40, 101)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(54, 13)
        Me.lbl1.TabIndex = 14
        Me.lbl1.Text = "Batch File"
        '
        'txtBatchFile
        '
        Me.txtBatchFile.BackColor = System.Drawing.Color.White
        Me.txtBatchFile.Location = New System.Drawing.Point(121, 98)
        Me.txtBatchFile.Name = "txtBatchFile"
        Me.txtBatchFile.ReadOnly = True
        Me.txtBatchFile.Size = New System.Drawing.Size(528, 20)
        Me.txtBatchFile.TabIndex = 2
        Me.txtBatchFile.TabStop = False
        '
        'btnBatchFile
        '
        Me.btnBatchFile.Location = New System.Drawing.Point(655, 96)
        Me.btnBatchFile.Name = "btnBatchFile"
        Me.btnBatchFile.Size = New System.Drawing.Size(48, 23)
        Me.btnBatchFile.TabIndex = 3
        Me.btnBatchFile.Text = "Select"
        Me.btnBatchFile.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(40, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Mode"
        '
        'chkOverwrite
        '
        Me.chkOverwrite.AutoSize = True
        Me.chkOverwrite.Checked = True
        Me.chkOverwrite.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOverwrite.Location = New System.Drawing.Point(301, 243)
        Me.chkOverwrite.Name = "chkOverwrite"
        Me.chkOverwrite.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOverwrite.Size = New System.Drawing.Size(134, 17)
        Me.chkOverwrite.TabIndex = 12
        Me.chkOverwrite.Text = "Overwrite Existing Files"
        Me.chkOverwrite.UseVisualStyleBackColor = True
        '
        'lbl5
        '
        Me.lbl5.AutoSize = True
        Me.lbl5.Location = New System.Drawing.Point(40, 27)
        Me.lbl5.Name = "lbl5"
        Me.lbl5.Size = New System.Drawing.Size(81, 13)
        Me.lbl5.TabIndex = 17
        Me.lbl5.Text = "Mainframe Host"
        '
        'txtMainframeHost
        '
        Me.txtMainframeHost.BackColor = System.Drawing.Color.White
        Me.txtMainframeHost.Location = New System.Drawing.Point(121, 23)
        Me.txtMainframeHost.MaxLength = 100
        Me.txtMainframeHost.Name = "txtMainframeHost"
        Me.txtMainframeHost.Size = New System.Drawing.Size(103, 20)
        Me.txtMainframeHost.TabIndex = 17
        '
        'lbl6
        '
        Me.lbl6.AutoSize = True
        Me.lbl6.Location = New System.Drawing.Point(266, 26)
        Me.lbl6.Name = "lbl6"
        Me.lbl6.Size = New System.Drawing.Size(93, 13)
        Me.lbl6.TabIndex = 19
        Me.lbl6.Text = "Mainframe User Id"
        '
        'txtMainframeUserId
        '
        Me.txtMainframeUserId.BackColor = System.Drawing.Color.White
        Me.txtMainframeUserId.Location = New System.Drawing.Point(365, 23)
        Me.txtMainframeUserId.MaxLength = 25
        Me.txtMainframeUserId.Name = "txtMainframeUserId"
        Me.txtMainframeUserId.Size = New System.Drawing.Size(103, 20)
        Me.txtMainframeUserId.TabIndex = 18
        '
        'lbl7
        '
        Me.lbl7.AutoSize = True
        Me.lbl7.Location = New System.Drawing.Point(489, 27)
        Me.lbl7.Name = "lbl7"
        Me.lbl7.Size = New System.Drawing.Size(105, 13)
        Me.lbl7.TabIndex = 21
        Me.lbl7.Text = "Mainframe Password"
        '
        'txtMainframePassword
        '
        Me.txtMainframePassword.BackColor = System.Drawing.Color.White
        Me.txtMainframePassword.Location = New System.Drawing.Point(600, 23)
        Me.txtMainframePassword.MaxLength = 25
        Me.txtMainframePassword.Name = "txtMainframePassword"
        Me.txtMainframePassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtMainframePassword.Size = New System.Drawing.Size(103, 20)
        Me.txtMainframePassword.TabIndex = 19
        '
        'frmQMF2Excel
        '
        Me.AcceptButton = Me.btnSubmit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(739, 482)
        Me.Controls.Add(Me.lbl7)
        Me.Controls.Add(Me.txtMainframePassword)
        Me.Controls.Add(Me.lbl6)
        Me.Controls.Add(Me.txtMainframeUserId)
        Me.Controls.Add(Me.lbl5)
        Me.Controls.Add(Me.txtMainframeHost)
        Me.Controls.Add(Me.chkOverwrite)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.txtBatchFile)
        Me.Controls.Add(Me.btnBatchFile)
        Me.Controls.Add(Me.cmbMode)
        Me.Controls.Add(Me.chkRememberScreen)
        Me.Controls.Add(Me.chkNumeric)
        Me.Controls.Add(Me.lstLog)
        Me.Controls.Add(Me.lbl4)
        Me.Controls.Add(Me.txtExcelFile)
        Me.Controls.Add(Me.lbl2)
        Me.Controls.Add(Me.txtMainframeFile)
        Me.Controls.Add(Me.lbl3)
        Me.Controls.Add(Me.txtExcelFolder)
        Me.Controls.Add(Me.btnExcelFolder)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSubmit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmQMF2Excel"
        Me.Text = "QMF2Excel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSubmit As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnExcelFolder As System.Windows.Forms.Button
    Friend WithEvents txtExcelFolder As System.Windows.Forms.TextBox
    Friend WithEvents lbl3 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents txtMainframeFile As System.Windows.Forms.TextBox
    Friend WithEvents txtExcelFile As System.Windows.Forms.TextBox
    Friend WithEvents lbl4 As System.Windows.Forms.Label
    Friend WithEvents lstLog As System.Windows.Forms.ListBox
    Friend WithEvents chkNumeric As System.Windows.Forms.CheckBox
    Friend WithEvents chkRememberScreen As System.Windows.Forms.CheckBox
    Friend WithEvents cmbMode As System.Windows.Forms.ComboBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents txtBatchFile As System.Windows.Forms.TextBox
    Friend WithEvents btnBatchFile As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents chkOverwrite As System.Windows.Forms.CheckBox
    Friend WithEvents lbl5 As System.Windows.Forms.Label
    Friend WithEvents txtMainframeHost As System.Windows.Forms.TextBox
    Friend WithEvents lbl6 As System.Windows.Forms.Label
    Friend WithEvents txtMainframeUserId As System.Windows.Forms.TextBox
    Friend WithEvents lbl7 As System.Windows.Forms.Label
    Friend WithEvents txtMainframePassword As System.Windows.Forms.TextBox

End Class
