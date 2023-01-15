Public Class Form3

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click
        Form1.Show()
    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Form2.Show()
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.IsMdiContainer = True
        Me.CenterToScreen()
    End Sub

End Class