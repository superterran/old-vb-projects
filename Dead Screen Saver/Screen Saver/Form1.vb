Public Class Form1


    Dim darkness As Integer
    Dim recount As Integer
    Dim args As String
    Dim trip As Boolean

    Dim was As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        args = UCase$(Trim$(Command$))

        Select Case Mid$(args, 1, 2)
            Case "/C"   ' Display configuration dialog.
                Form2.Show()
                trip = True
                Exit Sub

            Case "", "/S"   ' Run as a screen saver.

                GoTo 10
            Case "/P"       ' Run in preview mode.

                End

            Case Else       ' This shouldn't happen.

        End Select


10:     Windows.Forms.Cursor.Hide()
        trip = False

        Me.BringToFront()



        Me.Opacity = 0

        was = System.Windows.Forms.Control.MousePosition.X


        recount = 300
        darkness = 95
        Me.Visible = False
        Me.Opacity = 0

        On Error GoTo 50

        recount = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "FadeTime", 300)
        darkness = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ST's Dark Screen Saver", "Opac", 95)


50:




        Me.Visible = True

        Do Until Me.Opacity > (darkness / 100)

            Me.Opacity = Me.Opacity + (recount / 10000)

        Loop

        Exit Sub



    End Sub

 
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove

        If trip = True Then Exit Sub


        If was = MousePosition.X Then

        Else

            Do Until Me.Opacity = 0
                Me.Opacity = Me.Opacity - (recount / 10000)

            Loop

            End

        End If




    End Sub

  
End Class

