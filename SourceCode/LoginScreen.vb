Imports System.IO
Public Class frmLogin
  Dim Attempt As Integer = 1
  ' TODO: Insert code to perform custom authentication using a provided username and password
  '(See https://go.microsoft.com/fwlink/?LinkId=35339).
  'The custom principal can then be attached to the current thread's principal as follows:
  '       My.User.CurrentPrincipal = CustomPrincipal
  'where CustomPrincipal is the IPrincipal implementation used to perform authentication.
  'Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
  'such as the username, display name, etc.
  
  Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
    Dim inputFile As StreamReader = File.OpenText("sys.dll")
    Dim LinePW As String
    
    LinePW = inputFile.ReadLine()
    inputFile.Close()
    
    If txtUser.Text = "Admin" And TxtPassWord.Text = LinePW Then
      MsgBox("Welcome, " & txtUser.Text)
      frmInvoice.Show()
      Me.Close()
    ElseIf Attempt = 3 Then
      MsgBox("Program is now closing, Maxium number of " & Attempt & " attampts!")
      Me.Close()
    Else
      MsgBox("Username and password incorrect, Please re-enter, you currently have reached attempt " & Attempt & " of 3.")
        Attempt = Attempt + 1
        txtUser.Text = ""
        TxtPassWord.Text = ""
        txtUser.Focus()
      End If
    End Sub
 
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
      Me.Close()
    End Sub
 
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click 'Back Door! into Invoice bypassing Login
      frmInvoice.Show()
      Me.Close()
    End Sub
 
End Class
