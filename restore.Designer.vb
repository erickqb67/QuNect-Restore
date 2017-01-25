﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.dgMapping = New System.Windows.Forms.DataGridView()
        Me.OpenSourceFile = New System.Windows.Forms.OpenFileDialog()
        Me.btnSource = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnListTables = New System.Windows.Forms.Button()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.lblAppToken = New System.Windows.Forms.Label()
        Me.txtAppToken = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblTable = New System.Windows.Forms.Label()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.ckbDetectProxy = New System.Windows.Forms.CheckBox()
        Me.btnListFields = New System.Windows.Forms.Button()
        Me.chkBxHeaders = New System.Windows.Forms.CheckBox()
        Me.Source = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.op = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.criteria = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destination = New System.Windows.Forms.DataGridViewComboBoxColumn()
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgMapping
        '
        Me.dgMapping.AllowUserToAddRows = False
        Me.dgMapping.AllowUserToDeleteRows = False
        Me.dgMapping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgMapping.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Source, Me.op, Me.criteria, Me.Destination})
        Me.dgMapping.Location = New System.Drawing.Point(12, 230)
        Me.dgMapping.Name = "dgMapping"
        Me.dgMapping.Size = New System.Drawing.Size(871, 672)
        Me.dgMapping.TabIndex = 0
        '
        'OpenSourceFile
        '
        Me.OpenSourceFile.Filter = "Comma Separated Values (*.csv) | *.csv"
        '
        'btnSource
        '
        Me.btnSource.Location = New System.Drawing.Point(12, 145)
        Me.btnSource.Name = "btnSource"
        Me.btnSource.Size = New System.Drawing.Size(126, 27)
        Me.btnSource.TabIndex = 1
        Me.btnSource.Text = "Choose File to Import..."
        Me.btnSource.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(806, 174)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(77, 27)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Import"
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
        '
        'txtAppToken
        '
        Me.txtAppToken.Location = New System.Drawing.Point(12, 83)
        Me.txtAppToken.Name = "txtAppToken"
        Me.txtAppToken.Size = New System.Drawing.Size(258, 20)
        Me.txtAppToken.TabIndex = 29
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(289, 11)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(93, 13)
        Me.lblServer.TabIndex = 28
        Me.lblServer.Text = "QuickBase Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(286, 30)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(237, 20)
        Me.txtServer.TabIndex = 27
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(153, 11)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(108, 13)
        Me.lblPassword.TabIndex = 26
        Me.lblPassword.Text = "QuickBase Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(150, 30)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(120, 20)
        Me.txtPassword.TabIndex = 25
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(15, 11)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(110, 13)
        Me.lblUsername.TabIndex = 24
        Me.lblUsername.Text = "QuickBase Username"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(12, 30)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(120, 20)
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
        Me.lblFile.Location = New System.Drawing.Point(147, 152)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(0, 13)
        Me.lblFile.TabIndex = 35
        '
        'ckbDetectProxy
        '
        Me.ckbDetectProxy.AutoSize = True
        Me.ckbDetectProxy.Location = New System.Drawing.Point(556, 33)
        Me.ckbDetectProxy.Name = "ckbDetectProxy"
        Me.ckbDetectProxy.Size = New System.Drawing.Size(188, 17)
        Me.ckbDetectProxy.TabIndex = 36
        Me.ckbDetectProxy.Text = "Automatically detect proxy settings"
        Me.ckbDetectProxy.UseVisualStyleBackColor = True
        '
        'btnListFields
        '
        Me.btnListFields.Location = New System.Drawing.Point(718, 174)
        Me.btnListFields.Name = "btnListFields"
        Me.btnListFields.Size = New System.Drawing.Size(82, 27)
        Me.btnListFields.TabIndex = 37
        Me.btnListFields.Text = "List Fields"
        Me.btnListFields.UseVisualStyleBackColor = True
        '
        'chkBxHeaders
        '
        Me.chkBxHeaders.AutoSize = True
        Me.chkBxHeaders.Location = New System.Drawing.Point(12, 207)
        Me.chkBxHeaders.Name = "chkBxHeaders"
        Me.chkBxHeaders.Size = New System.Drawing.Size(137, 17)
        Me.chkBxHeaders.TabIndex = 38
        Me.chkBxHeaders.Text = "First row has field labels"
        Me.chkBxHeaders.UseVisualStyleBackColor = True
        '
        'Source
        '
        Me.Source.HeaderText = "Source"
        Me.Source.Name = "Source"
        Me.Source.ReadOnly = True
        Me.Source.Width = 200
        '
        'op
        '
        Me.op.HeaderText = "Import Condition"
        Me.op.Name = "op"
        Me.op.Width = 200
        '
        'criteria
        '
        Me.criteria.HeaderText = "Criteria"
        Me.criteria.Name = "criteria"
        Me.criteria.Width = 200
        '
        'Destination
        '
        Me.Destination.HeaderText = "Destination"
        Me.Destination.Name = "Destination"
        Me.Destination.Width = 200
        '
        'frmRestore
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(895, 924)
        Me.Controls.Add(Me.chkBxHeaders)
        Me.Controls.Add(Me.btnListFields)
        Me.Controls.Add(Me.ckbDetectProxy)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblTable)
        Me.Controls.Add(Me.pb)
        Me.Controls.Add(Me.lblAppToken)
        Me.Controls.Add(Me.txtAppToken)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.btnListTables)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnSource)
        Me.Controls.Add(Me.dgMapping)
        Me.Name = "frmRestore"
        Me.Text = "QuNect Restore"
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblTable As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents ckbDetectProxy As System.Windows.Forms.CheckBox
    Friend WithEvents btnListFields As System.Windows.Forms.Button
    Friend WithEvents chkBxHeaders As System.Windows.Forms.CheckBox
    Friend WithEvents Source As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents op As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents criteria As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Destination As System.Windows.Forms.DataGridViewComboBoxColumn

End Class