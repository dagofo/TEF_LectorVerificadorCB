Public Class VerifEtiquetas

    Public Resp2 As MsgBoxResult = MsgBoxResult.No

    Private Sub VerifEtiquetas_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Me.SimpleButton2.Focus()
    End Sub

    Private Sub VerifEtiquetas_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.SimpleButton2.Focus()
    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Resp2 = MsgBoxResult.No
        Me.Close()
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Resp2 = MsgBoxResult.Yes
        Me.Close()
    End Sub
End Class