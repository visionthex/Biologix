Public Class frmLoops
    Dim NumToMultiply As Integer
    Dim NumToDivion As Integer
    Dim NumToAddition As Integer
    Dim NumToSubtraction As Integer
    Dim NumToInteger As Integer
    Dim NumToMOD As Integer
    Dim NumToExponent As Integer
    Dim NumToPercent As Integer

    'Multiply
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Do While (NumToMultiply <= 100)
            ListBox1.Items.Add(NumToMultiply & " * " & Val(TextBox1.Text) & " = " NumToMultiply * Val(TextBox1.Text))
            NumToMultiply += 1
        Loop
    End Sub

    'Division
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Do While (NumToDivion <= 100)
            ListBox1.Items.Add(NumToDivion & " / " & Val(TextBox1.Text) & " = " NumToDivion / Val(TextBox1.Text))
            NumToDivion += 1
        Loop
    End Sub

    'Addition
        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Do While (NumToAddition <= 100)
            ListBox1.Items.Add(NumToAddition & " + " & Val(TextBox1.Text) & " = " NumToAddition + Val(TextBox1.Text))
            NumToAddition += 1
        Loop
    End Sub

    'Subtraction
        Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Do While (NumToSubtraction <= 100)
            ListBox1.Items.Add(NumToSubtraction & " - " & Val(TextBox1.Text) & " = " NumToSubtraction - Val(TextBox1.Text))
            NumToSubtraction += 1
        Loop
    End Sub

    'Integer
        Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Do While (NumToInteger <= 100)
            ListBox1.Items.Add(NumToInteger & " \ " & Val(TextBox1.Text) & " = " NumToInteger \ Val(TextBox1.Text))
            NumToInteger += 1
        Loop
    End Sub

    'MOD
        Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Do While (NumToMOD <= 100)
            ListBox1.Items.Add(NumToMOD & " MOD " & Val(TextBox1.Text) & " = " NumToMOD Mod Val(TextBox1.Text))
            NumToMod += 1
        Loop
    End Sub

    'Exponent
        Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Do While (NumToExponent <= 100)
            ListBox1.Items.Add(NumToExponent & " ^ " & Val(TextBox1.Text) & " = " NumToExponent ^ Val(TextBox1.Text))
            NumToExponent += 1
        Loop
    End Sub

    'Percentage
        Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Do While (NumToPercent <= 100)
            ListBox1.Items.Add(NumToPercent & " * " & Val(TextBox1.Text) & " / " & "" & "100" & "" & " = " & NumToPercent * Val(CDbl(TextBox1.Text) / 100))
            NumToPercent += 1
        Loop
    End Sub

    'Clear Button | To Clear Calculator ListBox
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ListBox1.Items.Clear()
        ListBox1.Refresh()
        TextBox1.Text = ""
        NumToMultiply = 0
        NumToDivion = 0
        NumToAddition = 0
        NumToSubtraction = 0
        NumToInteger = 0
        NumToMOD = 0
        NumToExponent = 0
        NumToPercent = 0
    End Sub

    'Exit Program
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.Close()
    End Sub
End Class
