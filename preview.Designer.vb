<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPreview
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
        Me.dgPreview = New System.Windows.Forms.DataGridView()
        Me.btnOK = New System.Windows.Forms.Button()
        CType(Me.dgPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgPreview
        '
        Me.dgPreview.AllowUserToAddRows = False
        Me.dgPreview.AllowUserToDeleteRows = False
        Me.dgPreview.AllowUserToOrderColumns = True
        Me.dgPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgPreview.Location = New System.Drawing.Point(12, 12)
        Me.dgPreview.Name = "dgPreview"
        Me.dgPreview.ReadOnly = True
        Me.dgPreview.Size = New System.Drawing.Size(1059, 769)
        Me.dgPreview.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(515, 783)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(70, 20)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1083, 815)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgPreview)
        Me.Name = "frmPreview"
        Me.Text = "Preview of import to QuickBase"
        CType(Me.dgPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgPreview As DataGridView
    Friend WithEvents btnOK As Button
End Class
