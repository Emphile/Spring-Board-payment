Option Explicit
Dim Innercity As Double
Dim Basicsalary As Double
Dim OverTime As Double
Dim Tax As Double
Dim Pension As Double
Dim StudentLoan As Double

Dim Period As String
Dim ddate As Date
Dim MonthNum As Integer
Dim NIPayment As Double


Private Sub Cobpayref_Change()
If Cobpayref.Text = ("784357") Then
       TxtAddress.Text = "12 allen way"
       TxtEmployeeName.Text = "peter parker"
       TxtPostcode.Text = "WU561P"
       CobGender.Text = "male"
    
       lblpayref.Caption = Evaluate("RANDBETWEEN(1000, 9999999)")
       TxtEmployerName.Text = "EverGreen Ltd"
    
       lbltaxcode.Caption = "5678"
       lblNINumber.Caption = "N1234231"
       lblNIcode.Caption = "NI1234231"


ElseIf Cobpayref.Text = ("784467") Then
         TxtAddress.Text = "13 westen way"
         TxtEmployeeName.Text = "Victor Samson"
         TxtPostcode.Text = "WB581P"
         CobGender.Text = "Female"
    
        lblpayref.Caption = Evaluate("RANDBETWEEN(1000, 9999999)")
        TxtEmployerName.Text = "EverGreen Ltd"
        
        lbltaxcode.Caption = "3658"
        lblNINumber.Caption = "N1234097"
        lblNIcode.Caption = "NI12344398"



ElseIf Cobpayref.Text = ("784548") Then
        TxtAddress.Text = "12 Bassey way"
        TxtEmployeeName.Text = "Emmanuel Moses"
        TxtPostcode.Text = "WP361N"
        CobGender.Text = "male"
        
        lblpayref.Caption = Evaluate("RANDBETWEEN(1000, 9999999)")
        TxtEmployerName.Text = "EverGreen Ltd"
        
        lbltaxcode.Caption = "5968"
        lblNINumber.Caption = "N1259231"
        lblNIcode.Caption = "NI9334231"

End If



End Sub

Private Sub Cobzoom_Change()
If Cobzoom.Text = 50 Then
    UserForm1.Zoom = 50

ElseIf Cobzoom.Text = 100 Then
    UserForm1.Zoom = 100
    UserForm1.Width = 1030
    UserForm1.Height = 800
    
ElseIf Cobzoom.Text = 120 Then
    UserForm1.Zoom = 120
    UserForm1.Width = 1328.25
    UserForm1.Height = 900
    
ElseIf Cobzoom.Text = 140 Then
    UserForm1.Zoom = 140
    UserForm1.Width = 1500
    UserForm1.Height = 1000
    
ElseIf Cobzoom.Text = 160 Then
    UserForm1.Zoom = 160
    UserForm1.Width = 1600
    UserForm1.Height = 1200
    
ElseIf Cobzoom.Text = 180 Then
    UserForm1.Zoom = 180
    UserForm1.Width = 1900
    UserForm1.Height = 1400
    
ElseIf Cobzoom.Text = 200 Then
    UserForm1.Zoom = 200
    UserForm1.Width = 2040
    UserForm1.Height = 1800
    
End If
End Sub

Private Sub Comaddpayment_Click()


Dim wks As Worksheet
Dim AddNew As Range
Set wks = Sheet1
Set AddNew = wks.Range("A65356").End(xlUp).Offset(1, 0)

AddNew.Offset(0, 0).Value = TxtEmployeeName.Text
AddNew.Offset(0, 1).Value = TxtAddress.Text
AddNew.Offset(0, 2).Value = TxtPostcode.Text
AddNew.Offset(0, 3).Value = CobGender.Text
AddNew.Offset(0, 4).Value = lblpayref.Caption
AddNew.Offset(0, 5).Value = TxtEmployerName.Text
AddNew.Offset(0, 6).Value = TxtBasicSalary.Text
AddNew.Offset(0, 7).Value = TxtInnerCity.Text
AddNew.Offset(0, 8).Value = Txtovertime.Text
AddNew.Offset(0, 9).Value = lblGrosspay.Caption
AddNew.Offset(0, 10).Value = lbltax.Caption
AddNew.Offset(0, 11).Value = lblpension.Caption
AddNew.Offset(0, 12).Value = lblstudentloan.Caption
AddNew.Offset(0, 13).Value = lblNIpayment.Caption
AddNew.Offset(0, 14).Value = lblDeductions.Caption
AddNew.Offset(0, 15).Value = lblDate.Caption
AddNew.Offset(0, 16).Value = lbltaxperiod.Caption
AddNew.Offset(0, 17).Value = lbltaxcode.Caption
AddNew.Offset(0, 18).Value = lblNINumber.Caption
AddNew.Offset(0, 19).Value = lblNIcode.Caption
AddNew.Offset(0, 20).Value = lbltaxablepay.Caption
AddNew.Offset(0, 21).Value = lblpensionablepay.Caption
AddNew.Offset(0, 22).Value = lblNetpay.Caption





End Sub

Private Sub CommandButton1_Click()
Dim iExit As VbMsgBoxResult

iExit = MsgBox("Confirm if you want to exit", vbQuestion + vbYesNo, "payroll System")

If iExit = vbYes Then
Unload Me
End If



End Sub


Private Sub CommandButton2_Click()
TxtInnerCity.Text = "0.00"
TxtBasicSalary.Text = "0.00"
Txtovertime.Text = "0.00"
lblGrosspay.Caption = ""
TxtAddress.Text = ""
TxtEmployeeName.Text = ""
TxtPostcode.Text = ""
CobGender.Clear
lblpayref.Caption = ""
TxtEmployerName.Text = ""
lbltax.Caption = ""
lblpension.Caption = ""
lblstudentloan.Caption = ""
lblNIpayment.Caption = ""
lblDeductions.Caption = ""
lbltaxperiod.Caption = ""
lbltaxcode.Caption = ""
lblNINumber.Caption = ""
lblNIcode.Caption = ""
lbltaxablepay.Caption = ""
lblpensionablepay.Caption = ""
lblNetpay.Caption = ""
LstpaySlip.Clear
Cobpayref.Clear


Call UserForm_Initialize

End Sub

Private Sub ComPayref_Click()
lblpayref.Caption = Evaluate("RANDBETWEEN(1000, 99999999)")
End Sub

Private Sub Comtotal_Click()
 Innercity = Val(TxtInnerCity.Text)
 Basicsalary = Val(TxtBasicSalary.Text)
 OverTime = Val(Txtovertime.Text)
 
 lblGrosspay.Caption = Innercity + Basicsalary + OverTime
lblGrosspay.Caption = Format(lblGrosspay.Caption, "£#,##0.00")

Tax = (((Innercity + Basicsalary + OverTime) * 9) / 100)
Pension = (((Innercity + Basicsalary + OverTime) * 12) / 100)
StudentLoan = (((Innercity + Basicsalary + OverTime) * 5) / 100)
NIPayment = (((Innercity + Basicsalary + OverTime) * 3) / 100)

lbltax.Caption = Format(Tax, "£#,##0.00")
lblpension.Caption = Format(Pension, "£#,##0.00")
lblstudentloan.Caption = Format(StudentLoan, "£#,##0.00")
lblNIpayment.Caption = Format(NIPayment, "£#,##0.00")

lblDeductions.Caption = Format(Tax + Pension + StudentLoan + NIPayment, "£#,##0.00")
lblNetpay.Caption = Format(lblGrosspay.Caption - lblDeductions.Caption, "£#,##0.00")

 ddate = Format(Date, "medium Date")
 MonthNum = Month(ddate)
 lbltaxperiod.Caption = MonthNum
 Period = lbltaxperiod.Caption
 
 lbltaxablepay.Caption = Format(Tax * Period, "£#,##0.00")
 lblpensionablepay.Caption = Format(Pension * Period, "£#,##0.00")
 
 
 LstpaySlip.AddItem ("EverGreen Ltd")
 LstpaySlip.AddItem ("======Pay Slip======")
 LstpaySlip.AddItem ("wages ref" + vbTab + vbTab + lblpayref.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("pay ref" + vbTab + vbTab + Cobpayref.Text)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("Name" + vbTab + vbTab + TxtEmployeeName.Text)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("tax period" + vbTab + vbTab + lbltaxperiod.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("NI Number" + vbTab + vbTab + lblNINumber.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("taxable pay" + vbTab + lbltaxablepay.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("pensionable pay" + vbTab + lblpensionablepay.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("Gross pay" + vbTab + vbTab + lblGrosspay.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("Deductions" + vbTab + vbTab + lblDeductions.Caption)
 LstpaySlip.AddItem (" ")
 LstpaySlip.AddItem ("Net pay" + vbTab + vbTab + lblNetpay.Caption)


End Sub

Private Sub Frame11_Click()

End Sub



Private Sub Frame2_Click()

End Sub

Private Sub Label46_Click()

End Sub

Private Sub Label59_Click()

End Sub

Private Sub lblGrosspay_Click()

End Sub

Private Sub listpayslip_Change()

End Sub



Private Sub TxtBasicSalary_Change()


End Sub
Private Sub TxtBasicSalary_Enter()
TxtBasicSalary.Text = ""
TxtBasicSalary.SetFocus

End Sub


Private Sub Txtstudentloan_Change()

End Sub

Private Sub Txttax_Change()

End Sub

Private Sub TxtBasicSalary_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If TxtBasicSalary.Text = "" Then
TxtBasicSalary.Text = "0.00"
End If

End Sub

Private Sub TxtInnerCity_Change()

End Sub

Private Sub TxtInnerCity_Enter()
TxtInnerCity.Text = ""
TxtInnerCity.SetFocus
End Sub

Private Sub TxtInnerCity_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If TxtInnerCity.Text = "" Then
TxtInnerCity.Text = "0.00"
End If

End Sub

Private Sub Txtovertime_Change()

End Sub

Private Sub Txtovertime_Enter()
Txtovertime.Text = ""
Txtovertime.SetFocus
End Sub

Private Sub Txtovertime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Txtovertime.Text = "" Then
Txtovertime.Text = "0.00"
End If
End Sub

Private Sub UserForm_Initialize()
lblDate.Caption = Format(Date, "medium Date")

CobGender.AddItem ("Female")
CobGender.AddItem ("Male")

Cobpayref.AddItem ("784357")
Cobpayref.AddItem ("784467")
Cobpayref.AddItem ("784548")
Cobpayref.AddItem ("7843623")
Cobpayref.AddItem ("784359")

Cobzoom.AddItem ("50")
Cobzoom.AddItem ("100")
Cobzoom.AddItem ("120")
Cobzoom.AddItem ("140")
Cobzoom.AddItem ("160")
Cobzoom.AddItem ("180")
Cobzoom.AddItem ("0")

End Sub

Private Sub wagesref_Click()

End Sub
