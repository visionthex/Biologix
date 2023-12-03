Import System.IO
Imports System.Drawing.Printing

Public Class frmMemoPad
    Dim lstLinesToPrint As New List(Of String)

#Region "Form Load"
    Private Sub frmMemoPad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
    End Sub
#End Region

#Region "Menu Strip"
    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub

    Private Sub mnuFormatFont_Click(sender As Object, e As EventArgs) Handles mnuFormatFont.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub mnuFormatColor_Click(sender As Object, e As EventArgs) Handles mnuFormatColor.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub mnuEditRedo_Click(sender As Object, e As EventArgs) Handles mnuEditRedo.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub mnuEditUndo_Click(sender As Object, e As EventArgs) Handles mnuEditUndo.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub mnuEditCut_Click(sender As Object, e As EventArgs) Handles mnuEditCut.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub mnuEditCopy_Click(sender As Object, e As EventArgs) Handles mnuEditCopy.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub mnuEditPaste_Click(sender As Object, e As EventArgs) Handles mnuEditPaste.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub mnuEditSelectAll_Click(sender As Object, e As EventArgs) Handles mnuEditSelectAll.Click
        RichTextBox1.SelectAll()
    End Sub
#End Region

#Region "Tool Menu Strip"
    Private Sub tsmUndo_Click(sender As Object, e As EventArgs) Handles tsmUndo.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub tsmRedo_Click(sender As Object, e As EventArgs) Handles tsmRedo.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub tsmCut_Click(sender As Object, e As EventArgs) Handles tsmCut.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub tsmCopy_Click(sender As Object, e As EventArgs) Handles tsmCopy.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub tsmPaste_Click(sender As Object, e As EventArgs) Handles tsmPaste.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub
#End Region

#Region "Save"
    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click
        Dim outputFileName As System.IO.StreamWriter
        SaveFileDialog1.Filter() = "Rich Text File (*.rtf)|*.rtf|Text File (*.txt)|*.txt|All Files (*.*)|*.*"
        Try 
            SaveFileDialog1.ShowDialog() 'Add Dialog boxes from toolbox to component tray on bottom of form
            outputFileName = File.CreateText(SaveFileDialog1.FileName)
            outputFileName.Write(RichTextBox1.rtf)
            outputFileName.Close()
        Catch
            MessageBox.Show("Sorry, The file cannot be created.", "File Create Error!")
        End Try
    End Sub
#End Region

#Region "Open"
    Private Sub mnuFileOpen_Click(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        Dim fileName = String.Empty
        Dim inputFile As StreamReader
        OpenFileDialog1.Filter = "Rich Text File (*.rtf)|*.rtf|All Files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        fileName = OpenFileDialog1.FileName
        inputFile = File.OpenText(fileName)
        RichTextBox1.rtf = inputFile.ReadtoEnd
    End Sub
#End Region

#Region "Print"
    Private Sub mnuFilePrint_Click(sender As Object, e As EventArgs) Handles mnuFilePrint.Click
        PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
        If PrintDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

            Dim PageSetup As New PageSettings
            With PageSetup
                .Margins.Left = 50
                .Margins.Right = 50
                .Margins.Top = 50
                .Margins.Bottom = 50
                .Landscape = False
            End With
            
            PrintDocument1.DefaultPageSettings = PageSetup
        End If
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_BeginPrint(sender As Object, e As Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        Dim fntText As Font = RichTextBox1.Font
        Dim txtWidth As Integer = PrintDocument1.DefaultPageSettings.PaperSize.Width - PrintDocument1.DefaultPageSettings.Margins.Left - PrintDocument1.DefaultPageSettings.Margins.Right
        Dim stringSize As New SizeF
        Dim g = Me.CreateGraphics
        listLinesToPrint.Clear()
        For inCounter = 0 To RichTextBox1.Font - 1
            stringSize = g.MeasureString(RichTextBox1.Lines(intCounter), fntText)
            If stringSize.Width < txtWidth Then
                lstLinesToPrint.Add(RichTextBox1.Lines(intCounter))
            Else
                Dim leftMargin As Integer = PrintDocument1.DefaultPageSettings.Margins.Left
                Dim TopMargin As Integer = PrintDocument1.DefaultPageSettings.Margins.Top
                Dim sfBuffer As SizeF = g.MeasureString("M", fntText)
                Dim layOutRec As New Rectangle(leftMargin, TopMargin, CInt(txtWidth - sfBuffer.Width), fntText.Height)
                Dim string_format As New Drawing.StringFormat
                string_format.Trimming = StringTrimming.Word
                Dim CharactersFitted As Integer = 0
                Dim LinesFilled As Integer = 0

                For intFittedChar = 0 To RichTextBox1.Lines(intCounter).Length - 1
                    g.MeasureString(RichTextBox1.Lines(intCounter).Substring(intFittedChar), fntText, New SizF(layOutRec.Width, layOutRec.Height), string_format, CharactersFitted, LinesFilled)
                    lstLinesToPrint.Add(RichTextBox1.Lines(intCounter).SubString(intFittedChar, CharactersFitted))
                    intFittedChar += CharactersFitted - 1
                Next
            End If
        Next
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static intStart As Integer
        Dim fntText As Font = RichTextBox1.Font
        Dim txtHeight As Integer
        Dim LeftMargin As Integer = PrintDocument1.DefaultPageSettings.Margins.Left
        Dim TopMargin As Integer = PrintDocument1.DefaultPageSettings.Margins.Top
        txtHeight = PrintDocument1.DefaultPageSettings.PaperSize.Height - PrintDocument1.DefaultPageSettings.Margins.Top - PrintDocument1.DefaultPageSettings.Margins.Bottom
        Dim LinesPerPage As Integer = CInt(Math.Round(txtHeight / fntText.Height + 0.025))
        e.Graphics.DrawRectangles(Pens.Red, e.MarginBounds)
        Dim intLineNumber As Integer

        For intCounter = intStart To RichTextBox1.Lines.Count - 1
            e.Graphics.DrawString(RichTextBox1.Lines(intCounter), fntText, Brushes.Black, LeftMargin, fntText.Height * intLineNumber + TopMargin)
            intLineNumber += 1
            If intLineNumber > LinesPerPage - 1 Then
                intStart = intCounter
                e.HasMorePages = True
                Exit For
            End If
        Next
    End Sub
#End Region

#Region "Tool Bar Menu"
    Private Sub tsbBold_Click(sender As Object, e As EventArgs) Handles tsbBold.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Bold)
    End Sub

    Private Sub tsbItalics_Click(sender As Object, e As EventArgs) Handles tsbItalics.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Italics)
    End Sub

    Private Sub tsbUnderline_Click(sender As Object, e As EventArgs) Handles tsbUnderline.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Underline)
    End Sub

    Private Sub tsbStrikeOut_Click(sender As Object, e As EventArgs) Handles tsbStrikeOut.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.StrikeOut)
    End Sub

    Private Sub tsbFont_Click(sender As Object, e As EventArgs) Handles tsbFont.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub tsbColor_Click(sender As Object, e As EventArgs) Handles tsbColor.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub tsbUpper_Click(sender As Object, e As EventArgs) Handles tsbUpper.Click
        RichTextBox1.SelectedText = RichTextBox1.SelectedText.ToUpper
    End Sub

    Private Sub tsbLower_Click(sender As Object, e As EventArgs) Handles tsbLower.Click
        RichTextBox1.SelectedText = RichTextBox1.SelectedText.ToLower
    End Sub

    Private Sub tsbLeftJustified_Click(sender As Object, e As EventArgs) Handles tsbLeftJustified.Click
        RichTextBox1.SelectedAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub tsbCenter_Click(sender As Object, e As EventArgs) Handles tsbCenter.Click
        RichTextBox1.SelectedAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub tsbRightJustified_Click(sender As Object, e As EventArgs) Handles tsbRightJustified.Click
        RichTextBox1.SelectedAlignment = HorizontalAlignment.Right
    End Sub

    Private Sub tsbCut_Click(sender As Object, e As EventArgs) Handles tsbCut.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub tsbCopy_Click(sender As Object, e As EventArgs) Handles tsbCopy.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub tsbPaste_Click(sender As Object, e As EventArgs) Handles tsbPaste.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub tsbDecreaseIndent_Click(sender As Object, e As EventArgs) Handles tsbDecreaseIndent.Click
        RichTextBox1.SelectionIndent = 0
    End Sub

    Private Sub tsbIncreaseIndent_Click(sender As Object, e As EventArgs) Handles tsbIncreaseIndent.Click
        RichTextBox1.SelectionIndent = 20
    End Sub

    Private Sub tsbBullets_Click(sender As Object, e As EventArgs) Handles tsbBullets.Click
        RichTextBox1.BulletIndent = 10
        RichTextBox1.SelectionBullet = True
    End Sub
#End Region

End Class
