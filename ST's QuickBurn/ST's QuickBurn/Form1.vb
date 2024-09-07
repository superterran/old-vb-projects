

Public Class Form1

    Public BurnDrive As Char
    Public NeroPath As String
    Public ImagePath As String


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load



       
        With SysTray
            .BalloonTipTitle = "ST's QuickBurn Alpha"
            .BalloonTipText = My.Computer.FileSystem.CurrentDirectory
            .BalloonTipIcon = ToolTipIcon.Info
            .ShowBalloonTip(5000)
        End With






        Call OpenImageFile()



    



        Me.Visible = False


    End Sub

    Public Sub OpenImageFile()

        With Open

            .Title = "Select Image file to Burn to CD/DVD"
            .ShowDialog()

            ImagePath = .FileName

        End With

        With SysTray

            .BalloonTipTitle = "ST's QuickBurn Alpha"
            .BalloonTipText = ImagePath
            .BalloonTipIcon = ToolTipIcon.Info
            .ShowBalloonTip(5000)

        End With

    End Sub

    Public Sub EndApp()

        SysTray.Dispose()
        End

    End Sub

    Private Sub SysTray_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SysTray.Click
        Call EndApp()

    End Sub

    Private Sub Form1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        Me.Visible = False

    End Sub
End Class

