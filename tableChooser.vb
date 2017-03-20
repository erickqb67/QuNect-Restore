Public Class frmTableChooser

    Private Sub tvAppsTables_DoubleClick(sender As Object, e As EventArgs) Handles tvAppsTables.DoubleClick
        If tvAppsTables.SelectedNode Is Nothing Then
            Me.Hide()
            Exit Sub
        End If
        If tvAppsTables.SelectedNode.Level <> 1 Then
            Exit Sub
        End If
        frmRestore.lblTable.Text = tvAppsTables.SelectedNode.FullPath()
        hideButtons()
        Me.Hide()
    End Sub
    Private Sub tvAppsTables_Click(sender As Object, e As EventArgs) Handles tvAppsTables.Click
        If tvAppsTables.SelectedNode Is Nothing Then
            Exit Sub
        End If
        If tvAppsTables.SelectedNode.Level <> 1 Then
            Exit Sub
        End If
        frmRestore.lblTable.Text = tvAppsTables.SelectedNode.FullPath()
        hideButtons()
    End Sub

    Private Sub btnDone_Click(sender As Object, e As EventArgs) Handles btnDone.Click
        If tvAppsTables.SelectedNode Is Nothing Then
            Me.Hide()
            Exit Sub
        End If

        If tvAppsTables.SelectedNode.Level <> 1 Then
            frmRestore.lblTable.Text = ""
        Else
            frmRestore.lblTable.Text = tvAppsTables.SelectedNode.FullPath()
        End If
        hideButtons()
        Me.Hide()
    End Sub
    Private Sub hideButtons()
        frmRestore.btnPreview.Visible = False
        frmRestore.btnImport.Visible = False
        frmRestore.dgCriteria.Visible = False
        frmRestore.dgMapping.Visible = False
    End Sub
End Class