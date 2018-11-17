<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRestore
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRestore))
        Me.dgMapping = New System.Windows.Forms.DataGridView()
        Me.Source = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destination = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.OpenSourceFile = New System.Windows.Forms.OpenFileDialog()
        Me.btnSource = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnListTables = New System.Windows.Forms.Button()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.lblAppToken = New System.Windows.Forms.Label()
        Me.txtAppToken = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.ckbDetectProxy = New System.Windows.Forms.CheckBox()
        Me.chkBxHeaders = New System.Windows.Forms.CheckBox()
        Me.dtPicker = New System.Windows.Forms.DateTimePicker()
        Me.dgCriteria = New System.Windows.Forms.DataGridView()
        Me.cmbCriteria = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.cmbOperator = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.lblMode = New System.Windows.Forms.Label()
        Me.cmbPassword = New System.Windows.Forms.ComboBox()
        Me.cmbBulkorSingle = New System.Windows.Forms.ComboBox()
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgCriteria, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgMapping
        '
        Me.dgMapping.AllowUserToAddRows = False
        Me.dgMapping.AllowUserToDeleteRows = False
        Me.dgMapping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgMapping.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Source, Me.Destination})
        Me.dgMapping.Location = New System.Drawing.Point(662, 230)
        Me.dgMapping.Name = "dgMapping"
        Me.dgMapping.Size = New System.Drawing.Size(445, 682)
        Me.dgMapping.TabIndex = 0
        '
        'Source
        '
        Me.Source.HeaderText = "Source"
        Me.Source.Name = "Source"
        Me.Source.ReadOnly = True
        Me.Source.Width = 200
        '
        'Destination
        '
        Me.Destination.HeaderText = "Destination"
        Me.Destination.Name = "Destination"
        Me.Destination.Width = 200
        '
        'OpenSourceFile
        '
        Me.OpenSourceFile.Filter = "Comma Separated Values (*.csv) | *.csv"
        '
        'btnSource
        '
        Me.btnSource.Location = New System.Drawing.Point(397, 142)
        Me.btnSource.Name = "btnSource"
        Me.btnSource.Size = New System.Drawing.Size(126, 27)
        Me.btnSource.TabIndex = 1
        Me.btnSource.Text = "Choose File to Import..."
        Me.btnSource.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(921, 176)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(185, 27)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Import from the CSV backup file"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnListTables
        '
        Me.btnListTables.Location = New System.Drawing.Point(12, 178)
        Me.btnListTables.Name = "btnListTables"
        Me.btnListTables.Size = New System.Drawing.Size(126, 23)
        Me.btnListTables.TabIndex = 4
        Me.btnListTables.Text = "Choose Table..."
        Me.btnListTables.UseVisualStyleBackColor = True
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(12, 116)
        Me.pb.Maximum = 1000
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(511, 23)
        Me.pb.TabIndex = 33
        Me.pb.Visible = False
        '
        'lblAppToken
        '
        Me.lblAppToken.AutoSize = True
        Me.lblAppToken.Location = New System.Drawing.Point(15, 64)
        Me.lblAppToken.Name = "lblAppToken"
        Me.lblAppToken.Size = New System.Drawing.Size(148, 13)
        Me.lblAppToken.TabIndex = 30
        Me.lblAppToken.Text = "QuickBase Application Token"
        Me.lblAppToken.Visible = False
        '
        'txtAppToken
        '
        Me.txtAppToken.Location = New System.Drawing.Point(12, 83)
        Me.txtAppToken.Name = "txtAppToken"
        Me.txtAppToken.Size = New System.Drawing.Size(258, 20)
        Me.txtAppToken.TabIndex = 29
        Me.txtAppToken.Visible = False
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(460, 19)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(93, 13)
        Me.lblServer.TabIndex = 28
        Me.lblServer.Text = "QuickBase Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(457, 38)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(237, 20)
        Me.txtServer.TabIndex = 27
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(243, 37)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(202, 20)
        Me.txtPassword.TabIndex = 25
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(15, 19)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(110, 13)
        Me.lblUsername.TabIndex = 24
        Me.lblUsername.Text = "QuickBase Username"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(12, 37)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(213, 20)
        Me.txtUsername.TabIndex = 23
        '
        'lblTable
        '
        Me.lblTable.AutoSize = True
        Me.lblTable.Location = New System.Drawing.Point(147, 183)
        Me.lblTable.Name = "lblTable"
        Me.lblTable.Size = New System.Drawing.Size(0, 13)
        Me.lblTable.TabIndex = 34
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(530, 152)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(0, 13)
        Me.lblFile.TabIndex = 35
        '
        'ckbDetectProxy
        '
        Me.ckbDetectProxy.AutoSize = True
        Me.ckbDetectProxy.Location = New System.Drawing.Point(727, 41)
        Me.ckbDetectProxy.Name = "ckbDetectProxy"
        Me.ckbDetectProxy.Size = New System.Drawing.Size(188, 17)
        Me.ckbDetectProxy.TabIndex = 36
        Me.ckbDetectProxy.Text = "Automatically detect proxy settings"
        Me.ckbDetectProxy.UseVisualStyleBackColor = True
        '
        'chkBxHeaders
        '
        Me.chkBxHeaders.AutoSize = True
        Me.chkBxHeaders.Location = New System.Drawing.Point(969, 148)
        Me.chkBxHeaders.Name = "chkBxHeaders"
        Me.chkBxHeaders.Size = New System.Drawing.Size(137, 17)
        Me.chkBxHeaders.TabIndex = 38
        Me.chkBxHeaders.Text = "First row has field labels"
        Me.chkBxHeaders.UseVisualStyleBackColor = True
        '
        'dtPicker
        '
        Me.dtPicker.Location = New System.Drawing.Point(286, 83)
        Me.dtPicker.Name = "dtPicker"
        Me.dtPicker.Size = New System.Drawing.Size(200, 20)
        Me.dtPicker.TabIndex = 39
        Me.dtPicker.Visible = False
        '
        'dgCriteria
        '
        Me.dgCriteria.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCriteria.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.cmbCriteria, Me.cmbOperator, Me.DataGridViewTextBoxColumn2})
        Me.dgCriteria.Location = New System.Drawing.Point(12, 230)
        Me.dgCriteria.Name = "dgCriteria"
        Me.dgCriteria.Size = New System.Drawing.Size(644, 682)
        Me.dgCriteria.TabIndex = 40
        '
        'cmbCriteria
        '
        Me.cmbCriteria.HeaderText = "Source"
        Me.cmbCriteria.Name = "cmbCriteria"
        Me.cmbCriteria.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.cmbCriteria.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.cmbCriteria.Width = 200
        '
        'cmbOperator
        '
        Me.cmbOperator.HeaderText = "Import Condition"
        Me.cmbOperator.Name = "cmbOperator"
        Me.cmbOperator.Width = 200
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Criteria"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 200
        '
        'btnPreview
        '
        Me.btnPreview.Location = New System.Drawing.Point(694, 176)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(221, 27)
        Me.btnPreview.TabIndex = 41
        Me.btnPreview.Text = "Preview which rows will be inported"
        Me.btnPreview.UseVisualStyleBackColor = True
        Me.btnPreview.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(587, 116)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 42
        '
        'lblMode
        '
        Me.lblMode.AutoSize = True
        Me.lblMode.Location = New System.Drawing.Point(589, 97)
        Me.lblMode.Name = "lblMode"
        Me.lblMode.Size = New System.Drawing.Size(0, 13)
        Me.lblMode.TabIndex = 43
        '
        'cmbPassword
        '
        Me.cmbPassword.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPassword.FormattingEnabled = True
        Me.cmbPassword.Items.AddRange(New Object() {"Please choose...", "QuickBase Password", "QuickBase User Token"})
        Me.cmbPassword.Location = New System.Drawing.Point(242, 12)
        Me.cmbPassword.Name = "cmbPassword"
        Me.cmbPassword.Size = New System.Drawing.Size(203, 21)
        Me.cmbPassword.TabIndex = 44
        '
        'cmbBulkorSingle
        '
        Me.cmbBulkorSingle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBulkorSingle.FormattingEnabled = True
        Me.cmbBulkorSingle.Items.AddRange(New Object() {"Please choose", "Import a single CSV file with control of column mapping and record selection", "Bulk import all CSV files in a directory with no mapping or selection control"})
        Me.cmbBulkorSingle.Location = New System.Drawing.Point(12, 146)
        Me.cmbBulkorSingle.Name = "cmbBulkorSingle"
        Me.cmbBulkorSingle.Size = New System.Drawing.Size(375, 21)
        Me.cmbBulkorSingle.TabIndex = 45
        Me.cmbBulkorSingle.Visible = False
        '
        'frmRestore
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1128, 924)
        Me.Controls.Add(Me.cmbBulkorSingle)
        Me.Controls.Add(Me.cmbPassword)
        Me.Controls.Add(Me.lblMode)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.dgCriteria)
        Me.Controls.Add(Me.dtPicker)
        Me.Controls.Add(Me.chkBxHeaders)
        Me.Controls.Add(Me.ckbDetectProxy)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblTable)
        Me.Controls.Add(Me.pb)
        Me.Controls.Add(Me.lblAppToken)
        Me.Controls.Add(Me.txtAppToken)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.btnListTables)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnSource)
        Me.Controls.Add(Me.dgMapping)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRestore"
        Me.Text = "QuNect Restore"
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgCriteria, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgMapping As System.Windows.Forms.DataGridView
    Friend WithEvents OpenSourceFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnSource As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnListTables As System.Windows.Forms.Button
    Friend WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents lblAppToken As System.Windows.Forms.Label
    Friend WithEvents txtAppToken As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents ckbDetectProxy As System.Windows.Forms.CheckBox
    Friend WithEvents chkBxHeaders As System.Windows.Forms.CheckBox
    Friend WithEvents dtPicker As DateTimePicker
    Friend WithEvents dgCriteria As DataGridView
    Friend WithEvents Source As DataGridViewTextBoxColumn
    Friend WithEvents Destination As DataGridViewComboBoxColumn
    Friend WithEvents cmbCriteria As DataGridViewComboBoxColumn
    Friend WithEvents cmbOperator As DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents btnPreview As Button
    Friend WithEvents lblProgress As Label
    Friend WithEvents lblMode As Label
    Friend WithEvents cmbPassword As ComboBox
    Friend WithEvents cmbBulkorSingle As ComboBox
    Friend WithEvents FolderBrowserDialog As FolderBrowserDialog
End Class
