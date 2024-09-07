Public Class Form2

    '  Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '     ColorDialog1.Color = PictureBox1.BackColor
    '     ColorDialog1.ShowDialog()
    '     PictureBox1.BackColor = ColorDialog1.Color
    '  End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        My.Computer.Registry.CurrentUser.CreateSubKey("ST's Dark Screen Saver")
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "FadeTime", TrackBar1.Value)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "Opac", TrackBar2.Value)

        End

    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        On Error GoTo 50

        TrackBar1.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "FadeTime", 300)
        TrackBar2.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "Opac", 900)

50:     Exit Sub


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        End

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class