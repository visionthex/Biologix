Public Class frmAnimation
    Dim UpKey As Boolean
    Dim DownKey As Boolean
    Dim Leftkey As Boolean
    Dim Rightkey As Boolean
    Dim Score As Integer = 0

    Private Sub Form2_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.W Then
            UpKey = True
        End If

        If e.KeyCode = Keys.S Then
            Downkey = True
        End If

        If e.KeyCode = Keys.A Then
            Leftkey = True
        End If

        If e.KeyCode = Keys.D Then
            Rightkey = True
        End If
    End Sub

    Private Sub Form2_KeyUp(Sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.W Then
            UpKey = False
        End If

        If e.KeyCode = Keys.S Then
            Downkey = False
        End If

        If e.KeyCode = Keys.A Then
            Leftkey = False
        End If

        If e.KeyCode = Keys.D Then
            Rightkey = False
        End If
    
        'Collision Detection
        If PictureBox1.Bounds.IntersectsWith(PictureBox2.Bounds) Then
            PictureBox2.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox3.Bounds) Then
            PictureBox3.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox4.Bounds) Then
            PictureBox4.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox5.Bounds) Then
            PictureBox5.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox6.Bounds) Then
            PictureBox6.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox7.Bounds) Then
            PictureBox7.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox8.Bounds) Then
            PictureBox8.Visible = False
        End If

        If PictureBox1.Bounds.IntersectsWith(PictureBox9.Bounds) Then
            PictureBox9.Visible = False
        End If
    End Sub

    Private Sub Movement_Tick(sender As Object, e As EventArgs) Handles Movement.Movement_Tick
        If UpKey = True Then
            PictureBox1.Top -= 3
            'Keeps Picture from going past top border
            If PictureBox1.Top <= 0 Then
                PictureBox1.Top = 0
            End If
        End If

        If DownKey = True Then
            PictureBox1.Top += 3
            'Keeps Picture from going past top border
            If PictureBox1.Top >= 600 Then
                PictureBox1.Top = 600
            End If
        End If

        If LeftKey = True Then
            PictureBox1.Top -= 3
            'Keeps Picture from going past top border
            If PictureBox1.Top <= 0 Then
                PictureBox1.Top = 0
            End If
        End If

        If RightKey = True Then
            PictureBox1.Top += 3
            'Keeps Picture from going past top border
            If PictureBox1.Top >= 1110 Then
                PictureBox1.Top = 1110
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If PictureBox2.Visible = False Then
            Score += 1
            Label1.Text - CStr(Score)
        End If

        If PictureBox9.Visible = False Then
            Timer1.Enabled = False
            Label1.Text = "YOU WIN!" : Label1.ForeColor = Color.Red
        End If

        If Score >= 120 Then
            Label1.Text = "YOU LOSE!" : Label1.ForeColor = Color.Yellow
        End If 
    End Sub
End Class



    


