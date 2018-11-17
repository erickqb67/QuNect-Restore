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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPreview))
        Me.btnOK = New System.Windows.Forms.Button()
        Me.lblPreview = New System.Windows.Forms.Label()
        Me.dgPreview = New System.Windows.Forms.DataGridView()
        CType(Me.dgPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnOK.Location = New System.Drawing.Point(0, 645)
        Me.btnOK.MaximumSize = New System.Drawing.Size(100, 20)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(100, 20)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'lblPreview
        '
        Me.lblPreview.AutoSize = True
        Me.lblPreview.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblPreview.Location = New System.Drawing.Point(0, 632)
        Me.lblPreview.Name = "lblPreview"
        Me.lblPreview.Size = New System.Drawing.Size(0, 13)
        Me.lblPreview.TabIndex = 5
        '
        'dgPreview
        '
        Me.dgPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgPreview.Location = New System.Drawing.Point(0, 0)
        Me.dgPreview.Name = "dgPreview"
        Me.dgPreview.Size = New System.Drawing.Size(1075, 632)
        Me.dgPreview.TabIndex = 6
        '
        'frmPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1075, 665)
        Me.Controls.Add(Me.dgPreview)
        Me.Controls.Add(Me.lblPreview)
        Me.Controls.Add(Me.btnOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPreview"
        Me.Text = "Preview of import to QuickBase"
        CType(Me.dgPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As Button
    Friend WithEvents lblPreview As Label
    Friend WithEvents dgPreview As DataGridView
End Class
