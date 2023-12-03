Module Module1
  Public InvCount, Qty, intQty As Integer
  Public Total, GrdTotal, Price, Grd As Decimal
  Public TaxRate As Decimal = 0.065D
  Public Response As Integer
  
  'Array
  Public arrGrd As Decimal
  Public arrCount As Integer
  Public ArrayInv As Integer = 100
  Public arrDescription(ArrayInv) As String
  Public arrPartNumber(ArrayInV) As Integer
  Public arrPartQty(ArrayInv) As Integer
  Public arrPartPrice(ArrayInv) as Decimal
  Public arrPartTotal(ArrayInv) as Decimal
  'End Array
  
  'PMT, Class-level constants
  Public Const MONTHS_YEARS As Double = 12 'Months per year
  Public Const NEW_RATE As Double = 0.05 'Interest rate, Chemical
  Public Const USED_RATE As Double = 0.08 'Interest rate, Biological
  Public AnnaulRate As Double = New_RATE 'Class-level variable to hold the annual interest rate
  Public BiologixCost, DownPayment, Loan, MonthlyPayment, Interest, Principal As Double 'Biologix cost
  Public cnt As Integer
  Public Months As Double
  Public Out As String 'Used to hold a line of output
  Public InputOk As Boolean = True
  'End PMT
End Module
