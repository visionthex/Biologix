'Charles Sanders
'Invoice Biologix
'Project Created 1/10/15
'To Do: Add If/Then Decisions

Option Strict On
Import System.IO

#Region "Code"
    Public Class frmInvoice
#End Region

#Region "Dims"
    Dim Count As Integer
    Dim colFormItems As New Collection()
    Dim colFormTexts As New Collection()
    Dim ctrlFormItems As Control 
        'Used to iterate the Form Items Collection
#End Region

#Region "Buttons"
    Private Sub btnDisclaimer_Click(sender As Object, e As EventArgs) Handles btnDisclaimer.Click
        lblDisclaimer.Visible = True
    End Sub

    Private Sub lblPartPrice_Click(sender As Object, e As EventArgs) Handles lblPartPrice.Click
            'Display the information about the Company.
        Messagebox.Show("A cash discount is a reduction in net price offered by a vender to a merchant, in order to encourage early payment of an invoice.", "Company Discount",
        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtPartPrice.Text = "" Or txtPartQty.Text = "" Or txtPartNum.Text = "" Or cboPartDescription.Text = "" Or lblCalcTotalPrice.Text = "" Then 
            Exit Sub
        End If

        If CDbl(txtPartNum.Text) = arrPartNumber(arrCount) Then
            MessageBox.Show("Record Already Saved", "Probably Duplicate")
                'Add above line so if Save is clicked twice, record is not saved twice
            Exit Sub
        End If

        If cboPartDescription.Text = arrDescription(arrCount) Then
            MessageBox.Show("Record Already Saved", "Probably Duplicate")
                'Add above line so if Save is clicked twice, record is not saved twice
            Exit Sub
        End If

        tssDisplay.Text = "File is being saved... "
        Timer1.Enabled = True

        InvCount += 1
        lblInvoiceCount.Text = CStr(InvCount)
        Grd = CDec(Grd + CDbl(lblGrandTotalCalc.Text))
        lblPrice.Text = FormatCurrency(Grd, 2)

            'Array
        arrCount += 1
        arrPartNumber(arrCount) = CInt(txtPartNum.Text)
        arrDescription(arrCount) = cboPartDescription.Text
        arrPartQty(arrCount) = CInt(txtPartQty.Text)
        arrPartPrice(arrCount) = CDec(txtPartPrice.Text)
        arrPartTotal(arrCout) = CDec(lblCalcTotalPrice.Text)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        tspProgress.Value = tspProgress.Value + 5
        If tspProgress.Value = 100 Then
            Timer1.Enabled = False
            tspProgress.Value = 0
            tssDisplay.Text = "Ready"
        End If
    End Sub

    Private Sub btnDisplay_Click(sender As Object, e As EventArgs) Handles btnDisplay.Click
        If lblPrice.Text = "" Then Exit Sub
        lstInvoice.Items.Clear()

        lst Invoice.Items.Add("Inv" & vbTab & "Part" & vbTab & "Desc" & vbTab & "Qty" & vbTab & "Price" & vbTab & "Total")

        For x = 1 To arrCount
            lstInvoice.Items.Add(x & vbTab & arrPartNumber(x) & vbTab & arrDescription(x) & vbTab & arrPartQty(x) & vbTab & FormatCurrency(arrPartPrice(x)) & vbTab & 
            FormatCurrency(arrPartTotal(x)))
            Next
            arrGrd = CDec(lblPrice.Text)
            lstInvoice.Items.Add(vbCrLf)
            lstInvoice.Itmes.Add("Grand Total:" & " " & FormatCurrency(arrGrd))
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
            'Clear the text boxes
        txtPartNum.Clear()
        txtPartPrice.Clear()
        txtPartQty.Clear()
        lblCalcTotalPrice.Text = String.Empty
        lblGrandTotalCalc.Text = String.Empty
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
#End Region

#Region "Form Load"
    Public Sub frmInvoice_Load(sender As Object, e As EventArgs) Handles Me.Load
        webBrowser.Navigate(tssURL.Text)
        DateTimePicker1.Text = CStr(Now())
        DateTimePicker1.CustomFormat = "MM-dd-yyyy"
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        tss.Date.Text = Format(Now(), "MM-dd-yyyy")

            'Collection of Labels
        colFormItems.Add(lblPurchaseDate)
        colFormItems.Add(lblPartDescription)
        colFormItems.Add(lblGrandTotal)
        colFormItems.Add(lblPartNum)
        colFormItems.Add(lblPartQty)
        colFormItems.Add(lblPartPrice)
        colFormItems.Add(lblTotalPrice)
        colFormItems.Add(lblTaxExept)
            'End Label Collection

            'Collection of Texts
        colFormTexts.Add(txtPartNum)
        colFormTexts.Add(DateTimePicker1)
        colFormTexts.Add(cboPartDescription)
        colFormTexts.Add(lstInvoice)
        colFormTexts.Add(txtPartQty)
        colFormTexts.Add(txtPartPrice)
        colFormTexts.Add(lblCalcTotalPrice)
        colFormTexts.Add(lblTax)
        colFormTexts.Add(lblGrandTotalCalc)
            'End Text Collection
    End Sub
#End Region

#Region "Text Boxes"
    Private Property Response As Windows.Forms.DialogResult

        Private Sub txtPartQty_LostFocus(sender As Object, e As EventArgs) Handles txtPartQty.LostFocus
            If txtPartQty.Text = "" Then Exit Sub
            If Integer.TryParse(txtPartQty.Text, intQty) Then
                Qty = CInt(CDec(txtPartQty.Text))
                Dim Response As DialogResult
                If CInt(txtPartQty.Text) > 500 Then
                    Beep()
                    Response = MessageBox.Show("Did you put input too much product?", "Out of Range", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If Response = vbYes Then   
                            txtPartQty.Focus()
                            Exit Sub
                        End If
                End If
            Else
                MessageBox.Show("You must enter a whole number", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                txtPartQty.Focus()
                Exit Sub
            End If
            Call Calcs()
        End Sub  

        Private Sub txtPartPrice_LostFocus(sander As Object, e As EventArgs) Handles txtPartPrice.LostFocus
            If txtPartPrice.Text = "" Then Exit Sub
                txtPartPrice.Text = FormatCurrency(txtPartPrice.Text, 2)
            If IsNumeric(txtPartPrice.Text) Then
                Price = CDec(txtPartPrice.Text)
            Else
                MessageBox.Show("Enter only numbers for Price!", "Data Entry Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                txtPartPrice.Focus()
                Exit Sub
            End If 
            Call Calcs()
        End Sub

        Private Sub label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
            Process.Start("http://www.animavix.com/php")
        End Sub
#End Region   

#Region "Calculations"
    Private Sub ckbTaxExempt_CheckedChanged(sender As Object, e As EventArgs) Handles ckbTaxExempt.CheckedChanged
        If ckbTaxExempt.Checked = True Then
            TaxRate = CDec(0.0)
            lblTax.Text = CStr(TaxRate)
        Else
            TaxRate = CDec(0.075)
            lblTax.Text = CStr(TaxRate)
        End If
        Call Calcs()
    End Sub

    Private Sub Calcs()
        If txtPartPrice.Text = "" Or txtPartQty.Text = "" Then Exit Sub
            Total = Price * Qty
            GrdTotal = CDec(Total * TaxRate)
            GrdTotal += Total
            lblCalcTotalPrice.Text = FormatCurrency(Total, 2)
            lblGrandTotalCalc.Text = FormatCurrency(GrdTotal, 2)

            Debug.WriteLine("Quantity= " & txtPartQty.Text & vbTab & "Price= " & txtPartPrice.Text)
            Debug.WriteLine("GrandTotalCalc= " & lblCalcTotalPrice.Text & vbTab & "TotalPrice= " & lblCalcTotalPrice.Text)
        End If

        Try
            If txtPartPrice.Text <> "" And txtPartQty.Text <> "" Then
                    'Check for Null.  You cannot calculate on null.
                lblCalcTotalPrice.Text = CStr(CDbl(txtPartPrice.Text) * CDbl(txtPartQty.Text))
                lblCalcTotalPrice.Text = FormatCurrency(lblCalcTotalPrice.Text, 2)
            End If

            Catch ex As Exception
                'Do this if an unexpected error/exception occurs
                'Unexpected because we already checked for null and numeric valuse
                '& is for concatenation, vbCRLF - Carriage Return/Line Feed
            Messagebox.Show("Unexpected Error: " & vbCrLF & ex.Message, "Error Message")
        End Try
    End Sub
#End Region

#Region "PMT"
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
            'Get the Biologix Cost, validating at the same time
        If Not Double.TryParse(txtCost.Text, BiologixCost) Then
            lblMessage.Text = "The cost must be a number!"
            InputOk = False
        End If

            'Get the down payment, validating at the same time
        If Not Double.TryParse(txtDownPayment.Text, DownPayment) Then
            lblMessage.Text = "Down payment must be a number!"
            InputOk = False
        End If

            'Get the number of months, validating at the same time
        If Not Double.TryParse(txtMonths.Text, Months) Then
            lblMessage.Text = "Months must be an integer!"
            InputOk = False
        End If

        If InputOk = True Then
                'Caclulate the loan amount and monthly payments.
            Loan = BiologixCost = DownPayment
            MonthlyPayment = Pmt(AnnualRate / MONTHS_YEARS, Months, -Loan)

                'Clear the list box message label.
            lstOutput.Itmes.Clear()
            lblMessage.Text = String.Empty
            lstOutput.Items.Add("Month" & vbTab & vbTab & "Payment" & vbTab & vbTab & "Interest" & vbTab & vbTab & "Principal")
            For Me.Count = 1 To CInt(Months)
                    'Calculate the interest for this period.
                Interest = IPmt(AnnualRate / MONTHS_YEARS, Count, Months, -Loan)
                    'Calculate the principal for this period
                Principal = PPmt(AnnaulRate / MONTHS_YEARS, Count, Months, -Loan)
                    'Start building the Output string with the month.
                lstOutput.Items.Add(Count & vbTab & vbTab & FormatCurrency(MonthlyPayment, 2) & vbTab & vbTab & FormatCurrency(Interest, 2) & vbTab & vbTab &
                FormatCurrency(Principal, 2))
            Next
        End If
    End Sub

    Private Sub btnClear2_Click(sender As Object, e As EventArgs) Handles btnClear2.Click
            'Reset the interest rate, clear the text boxes, the list box, and the message label.
            'See default interest rate for new Chemical loans.
        radNew.Checked = True
        AnnaulRate = NEW_RATE
        txtAnnualRate.Text = NEW_RATE.ToString("p")
        txtCost.Clear()
        txtDownPayment.Clear()
        txtMonths.Clear()
        lstOutput.Items.Clear()
        lblMessage.Text = String.Empty
        txtCost.Focus() 'Reset the focus to txtCost.
    End Sub

    Private Sub btnExit2_Click(sanders As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub radNew_CheckedChanged(sender As Object, e As EventArgs) Handles radNew.CheckedChanged
            'If the new radio button is checked, then the user has selected a new chemical loan.
        If radNew.Checked = True Then
            AnnaulRate = NEW_RATE
            txtAnnualRate.Text = NEW_RATE.ToString("p")
            lstOutput.Items.Clear()
        End If
    End Sub
    
    Private Sub radUsed_CheckedChanged(sender As Object, e As EventArgs) Handles radUsed.CheckedChanged
            'If the used radio button is checked, then the user has selected a used biological loan.
        If radUsed.Checked = True Then
            AnnualRate = USED_RATE
            txtAnnualRate.Text = USED_RATE.ToString("p")
            lstOutput.Items.Clear()
        End If
    End Sub
#End Region

#Region "Menu Forms"
    Private Sub mnuFractals_Click(sender As Object, e As EventArgs) Handles mnuFractals.Click
        frmFractals.ShowDialog()
    End Sub

    Private Sub mnuLoops_Click(sender As Object, e As EventArgs) Handles mnuLoops.Click
        frmLoops.ShowDialog()
    End Sub

    Private Sub mnuAnimation_Click(sender As Object, e As EventArgs) Handles mnuAnimation.Click
        frmAnimation.ShowDialog()
    End Sub

    Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        frmAbout.ShowDialog()
    End Sub

    Private Sub ArrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArrayToolStripMenuItem.Click
        frmArray.ShowDialog()
    End Sub

    Private Sub mnuMenu_Click(sender As Object, e As EventArgs) Handles mnuMenu.Click
        frmMemoPad.ShowDialog()
    End Sub

    Private Sub mnuFont_Click(sender As Object, e As EventArgs) Handles mnuFonts.Click
            'Font Charactors Customization
        Response = CType(MsgBox("Do you want to customize the labels? If No, you will be customizing text?", vbYesNo, "Font Customize"), Windows.Forms.DialogResult)
        If Response = vbYes Then
            FontDialog1.ShowDialog()
            For Each Me.ctrlFormItems In colFormItems
                ctrlFormItems.Font = FontDialog1.Font
            Next
        Else
            FontDialog1.ShowDialog()
            For Each Me.ctrlFormItems In colFormTexts
                ctrlFormItems.Font = FontDialog1.Font
            Next
        End If
    End Sub

    Private Sub mnuColor_Click(sender As Object, e As EventArgs) Handles mnuColor.Click
            'Font Color Cusotomization
        Response = CType(MsgBox("Do you want to customize the labels? If No, you will be customizing Text?", vbYesNo, "Color Customize"), Windows.Forms.DialogResult)
        If Response = vbYes Then
            ColorDialog1.ShowDialog()
            For Each Me.ctrlFormItems In colFormTexts
                ctrlFormItems.ForeColor = ColorDialog1.Color
            Next
        End If
    End Sub

    Private Sub mnuReset_Click(sender As Object, e As EventArgs) Handles mnuReset.Click
        Dim fontResetFont As New Font("Perpetua Tilting MT", 11.25, FontStyle.Bold)
        Dim fontResetFont2 As New Font("Perpetua Tilting MT", 9.75, FontStyle.Bold)
        Dim colorResetFont As New Color()
            'Collections of Labels
        For Each Me.ctrlFormItems In colFormItems
            ctrlFormItems.Font = fontResetFont
            ctrlFormItems.ForeColor = colorResetFont
        Next
            'Collections of Texts
        For Each Me.ctrlFormItems In colFormTexts
            ctrlFormItems.Font = fontResetFont2
            ctrlFormItems.ForeColor = colorResetFont
        Next
    End Sub

    Private Sub mnuPasswordReset_Click(sender As Object, e As EventArgs) Handles mnuPasswordReset.Click
            'Password Creation Reset and File save as sys.dll
        Dim FileName As System.IO.StreamWriter
        Dim InputLine As String

        Try
            FileName = File.CreateText("sys.dll")
            InputLine = InputBox("Enter Password", "Resetting Password")
            FileName.WriteLine(InputLine)
            FileName.Close()
        Catch
            MessageBox.Show("Sorry, the file cannot be created.")
        End Try
    End Sub

    Private Sub mnuHelp_Click(sender As Object, e As EventArgs) Handles mnuHelp.Click
        Process.Start("NotePad", "../Hints.txt")
    End Sub

    Private Sub ColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles cmuColor.Click
        mnuColor.PerformClick()
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles cmuFont.Click
        mnuFont.PerformClick()
    End Sub

    Private Sub CopyToolSripMenuItem1_Click(sender As Object, e As EventArgs) Handles cmuReset.Click
        mnuReset.PerformClick
    End Sub

    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub
#End Region

#Region "Browser"
    Private Sub tsbGo_Click(sender As Object, e As EventArgs) Handles tsbGo.Click
        webBrowser.Navigate(tssURL.Text)
        tssWebLoad.Text = "Loading Site... "
        Timer2.Enabled = True
    End Sub

    Private Sub tsbBack_Click(sender As Object, e As EventArgs) Handles tsbBack.Click
        webBrowser.GoBack()
        tssWebLoad.Text = "Going back to Site... "
        Timer2.Enabled = True
    End Sub

    Private Sub tsbForward_Click(sender As Object, e As EventArgs) Handles tsbForward.Click
        webBrowser.GoForward()
        tssWebLoad.Text = "Going to previous Site... "
        Timer2.Enabled = True
    End Sub

    Private Sub tsbRefresh_Click(sender As Object, e As EventArgs) Handles tsbRefresh.Click
        webBrowser.GoHome()
        tssWebLoad.Text = "Going Home... "
        Timer2.Enabled = True
    End Sub

    Private Sub tssURL_KeyDown(sender As Object, e As EventArgs) Handles tssURL.KeyDown
        If e.KeyCode = Keys.Enter Then
            webBrowser.Navigate(tssURL.Text)
            tssWebLoad.Text = "Loading Site... "
            Timer2.Enabled = True
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        tspWebloader.Value = tspWebloader.Value + 10
        If tspWebloader.Value = 100 Then
            Timer2.Enabled = False
            tspWebloader.Value = 0
            tssWebLoad.Text = "Ready"
        End If
    End Sub

    Private Sub tsbContact_Click(sender As Object, e As EventArgs) Handles tsbContact.Click
        MessageBox.Show("Company Email: dnaBiologix@gmail.com", "Email Contact:", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    End Class
#End Region




