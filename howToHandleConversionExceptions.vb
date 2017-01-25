Public Class frmErr



    Private Sub frmHowToHandleExceptions_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        TabControlErrors.Width = Me.Width - (4 * TabControlErrors.Left)
        TabControlErrors.Height = Me.Height - (4 * TabControlErrors.Top) - pnlButtons.Height
        pnlButtons.Top = Me.Height - pnlButtons.Height - (3 * TabControlErrors.Top)
    End Sub


   
End Class