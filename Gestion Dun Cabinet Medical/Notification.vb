Public Class Notification
    Dim i As Integer = 0
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
            Timer2.Start()
    End Sub

    Private Sub Notification_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Timerx.Start()
            Me.Top = Screen.PrimaryScreen.Bounds.Height - Me.Height - (Screen.PrimaryScreen.Bounds.Height - Screen.PrimaryScreen.WorkingArea.Height)
            'Me.Top = Screen.PrimaryScreen.Bounds.Height - Me.Height - 30
            Me.Left = Screen.PrimaryScreen.Bounds.Width - Me.Width
        Catch
        End Try
    End Sub

    Private Sub Timerx_Tick(sender As Object, e As EventArgs) Handles Timerx.Tick
        Try
            i = i + 1
            If i = 15 Then
                Timer2.Start()
            End If
        Catch
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            If Me.Opacity = 1 Then
                Timer1.Enabled = False
            Else
                Me.Opacity = Me.Opacity + 0.05
            End If
        Catch
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            If Me.Opacity = 0 Then
                Timer2.Enabled = False
                Me.Close()
            Else
                Me.Opacity = Me.Opacity - 0.05
            End If
        Catch
        End Try
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub
End Class