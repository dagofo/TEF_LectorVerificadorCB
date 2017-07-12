Public Class VerifOK

    Private Sub VerifOK_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Me.SimpleButton1.Focus()
    End Sub

    Private Sub VefifOK_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.SimpleButton1.Focus()
    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub
End Class