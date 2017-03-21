<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTableChooser
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTableChooser))
        Me.tvAppsTables = New System.Windows.Forms.TreeView()
        Me.btnDone = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tvAppsTables
        '
        Me.tvAppsTables.Location = New System.Drawing.Point(12, 12)
        Me.tvAppsTables.Name = "tvAppsTables"
        Me.tvAppsTables.Size = New System.Drawing.Size(449, 365)
        Me.tvAppsTables.TabIndex = 9
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(205, 389)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(61, 25)
        Me.btnDone.TabIndex = 10
        Me.btnDone.Text = "Done"
        Me.btnDone.UseVisualStyleBackColor = True
        '
        'frmTableChooser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(475, 426)
        Me.Controls.Add(Me.btnDone)
        Me.Controls.Add(Me.tvAppsTables)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTableChooser"
        Me.Text = "Choose a QuickBase Table"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tvAppsTables As System.Windows.Forms.TreeView
    Friend WithEvents btnDone As System.Windows.Forms.Button
End Class
