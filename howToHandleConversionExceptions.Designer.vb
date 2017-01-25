<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmErr
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
        Me.rdbSkipRecords = New System.Windows.Forms.RadioButton()
        Me.rdbBlankFields = New System.Windows.Forms.RadioButton()
        Me.rdbCancel = New System.Windows.Forms.RadioButton()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.TabControlErrors = New System.Windows.Forms.TabControl()
        Me.TabConversions = New System.Windows.Forms.TabPage()
        Me.lblConversions = New System.Windows.Forms.Label()
        Me.TabRequired = New System.Windows.Forms.TabPage()
        Me.lblRequired = New System.Windows.Forms.Label()
        Me.TabUnique = New System.Windows.Forms.TabPage()
        Me.lblUnique = New System.Windows.Forms.Label()
        Me.TabMalformed = New System.Windows.Forms.TabPage()
        Me.lblMalformed = New System.Windows.Forms.Label()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.TabMissing = New System.Windows.Forms.TabPage()
        Me.lblMissing = New System.Windows.Forms.Label()
        Me.TabControlErrors.SuspendLayout()
        Me.TabConversions.SuspendLayout()
        Me.TabRequired.SuspendLayout()
        Me.TabUnique.SuspendLayout()
        Me.TabMalformed.SuspendLayout()
        Me.pnlButtons.SuspendLayout()
        Me.TabMissing.SuspendLayout()
        Me.SuspendLayout()
        '
        'rdbSkipRecords
        '
        Me.rdbSkipRecords.AutoSize = True
        Me.rdbSkipRecords.Location = New System.Drawing.Point(27, 0)
        Me.rdbSkipRecords.Name = "rdbSkipRecords"
        Me.rdbSkipRecords.Size = New System.Drawing.Size(135, 17)
        Me.rdbSkipRecords.TabIndex = 0
        Me.rdbSkipRecords.Text = "Skip records with errors"
        Me.rdbSkipRecords.UseVisualStyleBackColor = True
        '
        'rdbBlankFields
        '
        Me.rdbBlankFields.AutoSize = True
        Me.rdbBlankFields.Location = New System.Drawing.Point(27, 23)
        Me.rdbBlankFields.Name = "rdbBlankFields"
        Me.rdbBlankFields.Size = New System.Drawing.Size(371, 17)
        Me.rdbBlankFields.TabIndex = 1
        Me.rdbBlankFields.Text = "Blank out fields with conversion errors (may erase data in existing records)"
        Me.rdbBlankFields.UseVisualStyleBackColor = True
        '
        'rdbCancel
        '
        Me.rdbCancel.AutoSize = True
        Me.rdbCancel.Checked = True
        Me.rdbCancel.Location = New System.Drawing.Point(27, 46)
        Me.rdbCancel.Name = "rdbCancel"
        Me.rdbCancel.Size = New System.Drawing.Size(136, 17)
        Me.rdbCancel.TabIndex = 2
        Me.rdbCancel.TabStop = True
        Me.rdbCancel.Text = "Cancel the entire import"
        Me.rdbCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(345, 54)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(70, 20)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TabControlErrors
        '
        Me.TabControlErrors.Controls.Add(Me.TabMissing)
        Me.TabControlErrors.Controls.Add(Me.TabConversions)
        Me.TabControlErrors.Controls.Add(Me.TabRequired)
        Me.TabControlErrors.Controls.Add(Me.TabUnique)
        Me.TabControlErrors.Controls.Add(Me.TabMalformed)
        Me.TabControlErrors.Location = New System.Drawing.Point(12, 12)
        Me.TabControlErrors.Name = "TabControlErrors"
        Me.TabControlErrors.SelectedIndex = 0
        Me.TabControlErrors.Size = New System.Drawing.Size(435, 356)
        Me.TabControlErrors.TabIndex = 5
        '
        'TabConversions
        '
        Me.TabConversions.Controls.Add(Me.lblConversions)
        Me.TabConversions.Location = New System.Drawing.Point(4, 22)
        Me.TabConversions.Name = "TabConversions"
        Me.TabConversions.Padding = New System.Windows.Forms.Padding(3)
        Me.TabConversions.Size = New System.Drawing.Size(427, 330)
        Me.TabConversions.TabIndex = 0
        Me.TabConversions.Text = "Conversions"
        Me.TabConversions.UseVisualStyleBackColor = True
        '
        'lblConversions
        '
        Me.lblConversions.AutoSize = True
        Me.lblConversions.Location = New System.Drawing.Point(9, 9)
        Me.lblConversions.Name = "lblConversions"
        Me.lblConversions.Size = New System.Drawing.Size(0, 13)
        Me.lblConversions.TabIndex = 0
        '
        'TabRequired
        '
        Me.TabRequired.Controls.Add(Me.lblRequired)
        Me.TabRequired.Location = New System.Drawing.Point(4, 22)
        Me.TabRequired.Name = "TabRequired"
        Me.TabRequired.Padding = New System.Windows.Forms.Padding(3)
        Me.TabRequired.Size = New System.Drawing.Size(427, 330)
        Me.TabRequired.TabIndex = 1
        Me.TabRequired.Text = "Required"
        Me.TabRequired.UseVisualStyleBackColor = True
        '
        'lblRequired
        '
        Me.lblRequired.AutoSize = True
        Me.lblRequired.Location = New System.Drawing.Point(7, 7)
        Me.lblRequired.Name = "lblRequired"
        Me.lblRequired.Size = New System.Drawing.Size(0, 13)
        Me.lblRequired.TabIndex = 0
        '
        'TabUnique
        '
        Me.TabUnique.Controls.Add(Me.lblUnique)
        Me.TabUnique.Location = New System.Drawing.Point(4, 22)
        Me.TabUnique.Name = "TabUnique"
        Me.TabUnique.Size = New System.Drawing.Size(427, 330)
        Me.TabUnique.TabIndex = 2
        Me.TabUnique.Text = "Unique"
        Me.TabUnique.UseVisualStyleBackColor = True
        '
        'lblUnique
        '
        Me.lblUnique.AutoSize = True
        Me.lblUnique.Location = New System.Drawing.Point(10, 9)
        Me.lblUnique.Name = "lblUnique"
        Me.lblUnique.Size = New System.Drawing.Size(0, 13)
        Me.lblUnique.TabIndex = 0
        '
        'TabMalformed
        '
        Me.TabMalformed.Controls.Add(Me.lblMalformed)
        Me.TabMalformed.Location = New System.Drawing.Point(4, 22)
        Me.TabMalformed.Name = "TabMalformed"
        Me.TabMalformed.Size = New System.Drawing.Size(427, 330)
        Me.TabMalformed.TabIndex = 3
        Me.TabMalformed.Text = "Malformed lines"
        Me.TabMalformed.UseVisualStyleBackColor = True
        '
        'lblMalformed
        '
        Me.lblMalformed.AutoSize = True
        Me.lblMalformed.Location = New System.Drawing.Point(8, 8)
        Me.lblMalformed.Name = "lblMalformed"
        Me.lblMalformed.Size = New System.Drawing.Size(0, 13)
        Me.lblMalformed.TabIndex = 0
        '
        'pnlButtons
        '
        Me.pnlButtons.Controls.Add(Me.btnOK)
        Me.pnlButtons.Controls.Add(Me.rdbCancel)
        Me.pnlButtons.Controls.Add(Me.rdbBlankFields)
        Me.pnlButtons.Controls.Add(Me.rdbSkipRecords)
        Me.pnlButtons.Location = New System.Drawing.Point(12, 374)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(422, 86)
        Me.pnlButtons.TabIndex = 6
        '
        'TabMissing
        '
        Me.TabMissing.Controls.Add(Me.lblMissing)
        Me.TabMissing.Location = New System.Drawing.Point(4, 22)
        Me.TabMissing.Name = "TabMissing"
        Me.TabMissing.Size = New System.Drawing.Size(427, 330)
        Me.TabMissing.TabIndex = 4
        Me.TabMissing.Text = "Missing Record ID#s"
        Me.TabMissing.UseVisualStyleBackColor = True
        '
        'lblMissing
        '
        Me.lblMissing.AutoSize = True
        Me.lblMissing.Location = New System.Drawing.Point(3, 7)
        Me.lblMissing.Name = "lblMissing"
        Me.lblMissing.Size = New System.Drawing.Size(0, 13)
        Me.lblMissing.TabIndex = 0
        '
        'frmErr
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 466)
        Me.Controls.Add(Me.pnlButtons)
        Me.Controls.Add(Me.TabControlErrors)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmErr"
        Me.ShowInTaskbar = False
        Me.Text = "Please Indicate How to Handle Import Errors"
        Me.TopMost = True
        Me.TabControlErrors.ResumeLayout(False)
        Me.TabConversions.ResumeLayout(False)
        Me.TabConversions.PerformLayout()
        Me.TabRequired.ResumeLayout(False)
        Me.TabRequired.PerformLayout()
        Me.TabUnique.ResumeLayout(False)
        Me.TabUnique.PerformLayout()
        Me.TabMalformed.ResumeLayout(False)
        Me.TabMalformed.PerformLayout()
        Me.pnlButtons.ResumeLayout(False)
        Me.pnlButtons.PerformLayout()
        Me.TabMissing.ResumeLayout(False)
        Me.TabMissing.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rdbSkipRecords As System.Windows.Forms.RadioButton
    Friend WithEvents rdbBlankFields As System.Windows.Forms.RadioButton
    Friend WithEvents rdbCancel As System.Windows.Forms.RadioButton
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TabControlErrors As System.Windows.Forms.TabControl
    Friend WithEvents TabConversions As System.Windows.Forms.TabPage
    Friend WithEvents TabRequired As System.Windows.Forms.TabPage
    Friend WithEvents TabUnique As System.Windows.Forms.TabPage
    Friend WithEvents TabMalformed As System.Windows.Forms.TabPage
    Friend WithEvents lblConversions As System.Windows.Forms.Label
    Friend WithEvents lblRequired As System.Windows.Forms.Label
    Friend WithEvents lblUnique As System.Windows.Forms.Label
    Friend WithEvents lblMalformed As System.Windows.Forms.Label
    Friend WithEvents pnlButtons As System.Windows.Forms.Panel
    Friend WithEvents TabMissing As System.Windows.Forms.TabPage
    Friend WithEvents lblMissing As System.Windows.Forms.Label
End Class
