Public Class Log

    Private Sub LogBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogBox.TextChanged

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Refresh()
    End Sub

    Private Sub Log_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class