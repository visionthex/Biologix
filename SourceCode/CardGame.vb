Public Class frmArray
    Dim Rank(13) As String
    Dim Suit(4) As String
    Dim Cards(4,13) As String

    Private Sub frmArray_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Rank(1) = "2"
        Rank(2) = "3"
        Rank(3) = "4"
        Rank(4) = "5"
        Rank(5) = "6"
        Rank(6) = "7"
        Rank(7) = "8"
        Rank(8) = "9"
        Rank(9) = "10"
        Rank(10) = "Jack"
        Rank(11) = "Queen"
        Rank(12) = "King"
        Rank(13) = "Ace"

        Suit(1) = "Diamonds"
        Suit(2) = "Spades"
        Suit(3) = "Hearts"
        Suit(4) = "Clubs"
    End Sub

    Private Sub tssArray_Click(sender As Object, e As EventArgs) Handles tssArray.Click
        LoadCards()
        PrintCards()
    End Sub

    Private Sub LoadCards()
        Dim x As Integer
        Dim y As Integer

        For x = 1 To 4
            For y = 1 To 13
                Cards(x, y) = Suit(x) + Rank(y)
            Next
        Next
    End Sub

    Private Sub ToolStripSplitButton1_Click(sender As Object, e As EventArgs) Handles tssClear.Click
        ListBox1.Items.Clear()
        MoveRight(4) = False
        MoveRight(5) = False

        For Me.x - 1 To 5
            Speed(x) = x
        Next
    End Sub
End Class
