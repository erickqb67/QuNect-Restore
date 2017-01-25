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
        Me.Hide()
    End Sub
End Class