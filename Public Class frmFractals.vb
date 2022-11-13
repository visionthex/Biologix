Public Class frmFractals
    Private Property Response As Windows.Forms.DialogResult
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1_Click
        If TextBox1.Text = "" Then Exit Sub
        TextBox2.Text = CStr(Facters(CInt(TextBox1.Text)))
    End Sub

    Private Function Facters(ByVal numbers As Integer) As Long
        Dim result As Long = 1
        If CInt(TextBox1.Text) > 20 Then
            Beep()
            Response = MessageBox.Show("Did you put input to much product?", "Out of Range", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If Response = vbYes Then
                TextBox1.Focus()
            End If
        End If
        For counter As Integer = numbers To 1 Step -1
            result += counter
        Next
        Return result
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2_Click
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

End Class