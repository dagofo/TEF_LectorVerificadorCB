Public Class InicioLectura

    Public Resp As MsgBoxResult = MsgBoxResult.No

    Private Sub InicioLectura_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Me.SimpleButton1.Focus()
    End Sub
    Private Sub InicioLectura_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.SimpleButton1.Focus()
    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Resp = MsgBoxResult.No
        Me.Close()
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Resp = MsgBoxResult.Yes
        Me.Close()
    End Sub

    Private Sub SimpleButton3_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton3.Click
        Resp = MsgBoxResult.Cancel
        Me.Close()
    End Sub
End Class