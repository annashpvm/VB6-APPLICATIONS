VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form pay_cal 
   Caption         =   "PAY CALCULATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122224641
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122224641
         CurrentDate     =   39359
      End
      Begin VB.Label Label4 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Report From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4080
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
      Begin VB.CommandButton Exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   870
         Left            =   1080
         Picture         =   "pay_calculation.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton pay_process 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&PROCESS"
         Height          =   870
         Left            =   120
         Picture         =   "pay_calculation.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.ComboBox cmb_month 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.ComboBox cmb_year 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "YEAR"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "MONTH"
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label pay_label 
      Caption         =   "SALARY / WAGES CALCULATION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   1230
      TabIndex        =   0
      Top             =   735
      Width           =   7275
   End
   Begin VB.Shape Shape1 
      Height          =   2610
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   8595
   End
End
Attribute VB_Name = "pay_cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim employee_code As Double
Dim employee_type As Integer
Dim aw As Integer
Dim el As Double   'eligible leave
Dim pl As Double   'Permission leave with loss of pay
Dim lp1 As Double  'Loss of pay
Dim lp2 As Double  '
Dim ab As Double   'Absent
Dim dh As Double   'Declare holiday
Dim wd As Double   'working days
Dim grosspay As Double
Dim earning_amount As Double
Dim deduction_amount As Double
Dim earning As Double
Dim pfamount As Double
Dim deduct_amount2 As Double
Dim attn_deduct_days As Double
Dim attn_deduct_amt  As Double
Dim tea_allowance As Double
Dim vdaamt As Double
Dim fdaamt As Double
Dim pfeligible As Double
Dim week_off_day As String
Private Sub cmb_month_Click()
    find_dates
End Sub
Private Sub cmb_year_Click()
    find_dates
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
''    If pay_calchk = 0 Then
''       pay_label.Caption = pay_label + " STAFF"
''    ElseIf pay_calchk = 1 Then
''       pay_label.Caption = pay_label + " WORKER"
''    Else
''       pay_label.Caption = pay_label + " RETAINER"
''    End If
    
    With cmb_month
        .AddItem "January"
        .ItemData(.NewIndex) = 1
        .AddItem "February"
        .ItemData(.NewIndex) = 2
        .AddItem "March"
        .ItemData(.NewIndex) = 3
        .AddItem "April"
        .ItemData(.NewIndex) = 4
        .AddItem "May"
        .ItemData(.NewIndex) = 5
        .AddItem "June"
        .ItemData(.NewIndex) = 6
        .AddItem "July"
        .ItemData(.NewIndex) = 7
        .AddItem "August"
        .ItemData(.NewIndex) = 8
        .AddItem "September"
        .ItemData(.NewIndex) = 9
        .AddItem "October"
        .ItemData(.NewIndex) = 10
        .AddItem "November"
        .ItemData(.NewIndex) = 11
        .AddItem "December"
        .ItemData(.NewIndex) = 12
    End With
''    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''
''    End With
''    cmb_year.Text = "2015"
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
End Sub
'''Private Sub pay_process_Click()
'''    If Trim(cmb_month.Text) = "" Or Trim(cmb_year.Text) = "" Then Exit Sub
'''    vdaamt = 0
'''    If pay_calchk = 1 Then
'''       Set paydb = New ADODB.Connection
'''       Set payrs = New ADODB.Recordset
'''       sql = "select * from emp_vda where v_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and v_year = " & Val(Trim(cmb_year.Text) & " and v_company = '" & company_code & "'")
'''       paydb.Open pay
'''       payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
'''       If Not payrs.EOF Then
'''          vdaamt = payrs.Fields("v_vdaamount")
'''          fdaamt = payrs.Fields("v_fdaamount")
'''       Else
'''          MsgBox ("Please Enter VDA Amount in VDA master")
'''          Exit Sub
'''       End If
'''    End If
'''    Set paydb2 = New ADODB.Connection
'''    Set payrs2 = New ADODB.Recordset
'''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
'''    If pay_calchk = 0 Then
'''       sql2 = "Select * from emp_salary where s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and (s_emptype = 0 or s_emptype = 1)"
'''    Else
'''       sql2 = "Select * from emp_salary where s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and (s_emptype = 2 or s_emptype = 3)"
'''    End If
'''    paydb2.Open pay
'''    payrs2.Open sql2, paydb2, adOpenDynamic, adLockOptimistic
'''    If Not payrs2.EOF Then
'''       While Not payrs2.EOF
'''           payrs2.Delete
'''           payrs2.MoveNext
'''       Wend
'''    End If
'''    Set paydb = New ADODB.Connection
'''    Set payrs = New ADODB.Recordset
'''    If pay_calchk = 0 Then
'''       sql = "Select * from emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1)"
'''    Else
'''       sql = "Select * from emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3)"
'''    End If
'''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
'''    paydb.Open pay
'''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
'''    emp_chk = 0
'''    While Not payrs.EOF
'''       If payrs.Fields("emp_type") <> 4 Then
'''       emp_code = payrs.Fields("emp_code")
'''       employee_code = payrs.Fields("emp_code")
'''       employee_type = payrs.Fields("emp_type")
'''       week_off_day = Format(payrs.Fields("emp_holiday"), "dddd")
'''       el = 0
'''       pl = 0
'''       wd = 0
'''       aw = 0
'''       dh = 0
'''       ab = 0
'''       lp1 = 0
'''       lp2 = 0
'''       find_attendance
'''       find_actualworkdays
'''''       If pay_calchk = 1 Then
'''''          find_actualworkdays
'''''       Else
'''''          aw = 26
'''''       End If
'''       Set paydb2 = New ADODB.Connection
'''       Set payrs2 = New ADODB.Recordset
'''       sql2 = "Select * from emp_salary"
'''       paydb2.Open pay
'''       payrs2.Open sql2, paydb2, adOpenDynamic, adLockOptimistic
'''       payrs2.AddNew
'''       payrs2.Fields("s_company") = company_code
'''       payrs2.Fields("s_month") = cmb_month.ItemData(cmb_month.ListIndex)
'''       payrs2.Fields("s_year") = Val(cmb_year.Text)
'''       payrs2.Fields("s_empcode") = employee_code
'''       payrs2.Fields("s_emptype") = employee_type
'''       payrs2.Fields("s_avlworkdays") = aw
'''       If pay_calchk = 1 Then
'''          payrs2.Fields("s_actworkdays") = wd
'''       Else
'''          wd = 26 - (el + ab + lp2 + lp1 + pl)
'''          payrs2.Fields("s_actworkdays") = wd
'''       End If
'''       payrs2.Fields("s_dec_holiday") = dh
'''       payrs2.Fields("s_eli_leave") = el
'''       payrs2.Fields("s_absent") = ab
'''       payrs2.Fields("s_layoff") = lp2
'''       payrs2.Fields("s_lossofpay") = lp1 + pl
'''       payrs2.Fields("s_basic") = payrs.Fields("emp_basic")
'''       payrs2.Fields("s_serwt") = payrs.Fields("emp_serwt")
'''       payrs2.Fields("s_splpay") = payrs.Fields("emp_splpay")
'''       If employee_type = 2 Then
'''          payrs2.Fields("s_vda") = vdaamt
'''          payrs2.Fields("s_fda") = fdaamt
'''       Else
'''          payrs2.Fields("s_vda") = 0
'''          payrs2.Fields("s_fda") = 0
'''       End If
'''       payrs2.Fields("s_hra") = payrs.Fields("emp_hra")
'''       payrs2.Fields("s_etotal") = payrs.Fields("emp_basic") + payrs.Fields("emp_serwt") + payrs.Fields("emp_splpay") + payrs.Fields("emp_fda") + vdaamt
'''       oneday = Round(payrs2.Fields("s_etotal") / aw, 4)
''''Attendance & tea allowance  formula
'''       attend_allowance = payrs.Fields("emp_attall")
'''       tea_allowance = payrs.Fields("emp_teaall")
'''       If employee_type = 2 Or employee_type = 3 Then
'''          If (el + pl + lp1) > 3 Then
'''             attn_deduct_days = el + pl + lp1 - 3
'''          Else
'''             attn_deduct_days = 0
'''          End If
'''          If ab > 0 Then attn_deduct_days = attn_deduct_days + ab
'''          attn_deduct_amt = attn_deduct_days * 5    'per day rs.5 for attendence deduction
'''          If attend_allowance > attn_deduct_amt Then
'''             attend_allowance = attend_allowance - attn_deduct_amt
'''          Else
'''             attend_allowance = 0
'''          End If
'''          If (wd + el + pl) >= aw Then
'''             tea_allowance = payrs.Fields("emp_teaall")
'''          Else
'''             tea_allowance = Round((Round((payrs.Fields("emp_teaall") / aw), 2) * (wd + el + pl)), 2)
'''          End If
'''       End If
'''       earning_amount = payrs2.Fields("s_etotal")
'''       dec_holiday_amount = Round(oneday * dh, 2)
'''       deduction_amount = Round((ab + lp1 + lp2 + pl) * oneday, 2)
'''       earning = earning_amount - deduction_amount + dec_holiday_amount
'''       If employee_type = 0 Or employee_type = 2 Then
'''          pfamount = Round(earning * 0.12, 0)
'''       Else
'''          pfamount = 0
'''       End If
'''       grosspay = earning + payrs.Fields("emp_attall") + payrs.Fields("emp_splall") + payrs.Fields("emp_teaall") + payrs.Fields("emp_medall") + _
'''                  payrs.Fields("emp_washall") + payrs.Fields("emp_convall") + payrs.Fields("emp_lta") + payrs.Fields("emp_mazall") + payrs.Fields("emp_fuelall") + _
'''                  payrs.Fields("emp_profall") + payrs.Fields("emp_phoneall") + payrs.Fields("emp_cityall") + payrs.Fields("emp_othall") + payrs.Fields("emp_hra") + _
'''                  payrs.Fields("emp_mealsall") + payrs.Fields("emp_eduall")
'''       netpay = grosspay - (pfamount + payrs.Fields("emp_lic") + payrs.Fields("emp_rd") + payrs.Fields("emp_houserent") + deduct_amount2)
'''       payrs2.Fields("s_earning") = earning
'''       payrs2.Fields("s_attall") = attend_allowance
'''       payrs2.Fields("s_splall") = payrs.Fields("emp_splall")
'''       payrs2.Fields("s_teaall") = tea_allowance
'''       payrs2.Fields("s_medall") = payrs.Fields("emp_medall")
'''       payrs2.Fields("s_washall") = payrs.Fields("emp_washall")
'''       payrs2.Fields("s_convall") = payrs.Fields("emp_convall")
'''       payrs2.Fields("s_lta") = payrs.Fields("emp_lta")
'''       payrs2.Fields("s_mazall") = payrs.Fields("emp_mazall")
'''       payrs2.Fields("s_fuelall") = payrs.Fields("emp_fuelall")
'''       payrs2.Fields("s_profall") = payrs.Fields("emp_profall")
'''       payrs2.Fields("s_phoneall") = payrs.Fields("emp_phoneall")
'''       payrs2.Fields("s_cityall") = payrs.Fields("emp_cityall")
'''       payrs2.Fields("s_mealsall") = payrs.Fields("emp_mealsall")
'''       payrs2.Fields("s_eduall") = payrs.Fields("emp_eduall")
'''       payrs2.Fields("s_other") = payrs.Fields("emp_othall")
'''       payrs2.Fields("s_grosspay") = Round(grosspay, 2)
'''       payrs2.Fields("s_pf") = pfamount
'''       payrs2.Fields("s_lic") = payrs.Fields("emp_lic")
'''       payrs2.Fields("s_rd") = payrs.Fields("emp_rd")
'''       payrs2.Fields("s_houserent") = payrs.Fields("emp_houserent")
'''       payrs2.Fields("s_otherdeductions") = deduct_amount2
'''       payrs2.Fields("s_netpay") = netpay
'''       payrs2.Fields("s_deptcode") = payrs.Fields("emp_dept")
'''       payrs2.Update
'''       End If
'''       payrs.MoveNext
'''    Wend
'''    MsgBox ("Processing over")
'''    Beep
'''End Sub
'''''Function find_attendance()
'''''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
'''''    Set paydb2 = New ADODB.Connection
'''''    Set payrs2 = New ADODB.Recordset
'''''    If pay_calchk = 0 Then
'''''       sql2 = "select * from attn_entry where month(attn_date) = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcode = " & employee_code & " and attn_company = '" & company_code & "' and (attn_emptype = 1 or attn_emptype = 0) "
'''''    Else
'''''       sql2 = "select * from attn_entry where month(attn_date) = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcode = " & employee_code & " and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3) "
'''''    End If
'''''    paydb2.Open pay
'''''    payrs2.Open sql2, paydb2, adOpenDynamic, adLockOptimistic
'''''    While Not payrs2.EOF
'''''       If RTrim(week_off_day) = UCase(RTrim(Format(payrs2.Fields("attn_date"), "dddd"))) And payrs2.Fields("attn_status") = 5 Then
'''''          att_status = 13
'''''       Else
'''''          att_status = payrs2.Fields("attn_status")
'''''       End If
'''''       Select Case att_status
'''''       Case 0, 10, 11
'''''            wd = wd + 1
'''''       Case 2
'''''            el = el + 1
'''''       Case 3
'''''            el = el + 0.5
'''''            wd = wd + 0.5
'''''       Case 13
'''''            dh = dh + 1
'''''       Case 7
'''''            ab = ab + 1
'''''       Case 4
'''''            pl = pl + 1
'''''       Case 8
'''''            lp1 = lp1 + 1
'''''       Case 9
'''''            lp1 = lp1 + 0.5
'''''            wd = wd + 0.5
'''''       Case 6
'''''            lp2 = lp2 + 1
'''''       Case 12
'''''            ab = ab + 0.5
'''''            wd = wd + 0.5
'''''       Case 14
'''''            pl = pl + 0.5
'''''       Case 15
'''''            pl = pl + 0.5
'''''            el = el + 0.5
'''''       End Select
'''''       payrs2.MoveNext
'''''    Wend
'''''End Function
'''''Function find_actualworkdays()
'''''    deduct_amount2 = 0
'''''    Set paydb2 = New ADODB.Connection
'''''    Set payrs2 = New ADODB.Recordset
'''''    If pay_calchk = 0 Then
'''''       sql2 = "select * from monthly_deduction where e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_emp_code = " & employee_code & " and e_company = '" & company_code & "' and (e_emp_type = 0 or e_emp_type = 1)"
'''''    Else
'''''       sql2 = "select * from monthly_deduction where e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_emp_code = " & employee_code & " and e_company = '" & company_code & "' and (e_emp_type = 2 or e_emp_type = 3)"
'''''    End If
'''''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
'''''    paydb2.Open pay
'''''    payrs2.Open sql2, paydb2, adOpenDynamic, adLockOptimistic
'''''    While Not payrs2.EOF
'''''       aw = payrs2.Fields("e_avail_workdays")
'''''       deduct_amount2 = deduct_amount2 + payrs2.Fields("e_ded_amount")
'''''       payrs2.MoveNext
'''''    Wend
'''''End Function

Private Sub pay_process_Click()
On Error GoTo err_handler
    Dim esi_eligible, deposit_amt As Double
    Dim esi_contri, ei_esi_ded, ei_esi_ded2   As Double
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
        MsgBox "Out Of Financial Year", vbInformation, "Message"
        Exit Sub
    End If
  
    Dim newESIEligibleYN, ESIEligibleFor As String
    
    Dim ESI_EL_Amount1, ESI_EL_Amount2, newESIEligible_Amount As Double
    
    
    Dim cyear As Integer
    Dim payrs As New ADODB.Recordset
    Dim payrs2 As New ADODB.Recordset
    
    Dim payrs_ot As New ADODB.Recordset
    
    sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       If payrs("pay_salary_lock") = "Y" Then
          MsgBox ("Salary Process status is LOCKED . Can't Continue...")
          payrs.Close
          Exit Sub
       End If
    End If
    payrs.Close
    
    
    
    sql = "select * from comp_mas where comp_code = '" & company_code & "'"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
        pfeligible = payrs("comp_pf_eligible")
        esi_eligible = payrs("comp_esi_eligible")
        esi_contri = payrs("comp_esi_emp1_contri")
    End If
    payrs.Close
    Dim el_gpay, el_pf, el_attn_all As Double
    Dim vda, fda As String
    el_gpay = 0
    vdaamt = 0
    fdaamt = 0
    vda = 0
    fda = 0
    If Trim(cmb_month.Text) = "" Or Trim(cmb_year.Text) = "" Then
       MsgBox ("Select Month / Year...")
       Exit Sub
    End If
''''for checking Deduction amount checking for employees
''    If data_source = "H" Then ''head office checking
''        sql = "select emp_name from attn_entry a , emp_mas b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcat = 'S' and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and emp_workplace = 'CBE' and attn_empcode not in (select e_emp_code from monthly_deduction , emp_mas  where e_company = emp_company  and e_emp_code = emp_code and emp_workplace = 'CBE' and e_company = " & company_code & " and e_emp_cat = 'S' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
''    Else
''        If emptype_chk = 0 Then
''           sql = "select emp_name from attn_entry a , emp_mas b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcat = 'S' and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and attn_empcode not in (select e_emp_code from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'S' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
''        ElseIf emptype_chk = 1 Then
''           sql = "select emp_name from attn_entry a , emp_mas b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcat = 'W' and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and attn_empcode not in (select e_emp_code from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'W' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
''        ElseIf emptype_chk = 2 Then
''           sql = "select emp_name from attn_entry a , emp_mas b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcat = 'M' and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and attn_empcode not in (select e_emp_code from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'M' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
''        ElseIf emptype_chk = 3 Then
''           sql = "select emp_name from attn_entry a , emp_voupay_mast  b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcat = 'R' and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and attn_empcode not in (select e_emp_code from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'R' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
''        End If
''    End If

    sql = "select emp_name,emp_fpcode from attn_entry a , emp_mas b where attn_company = " & company_code & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_company = emp_company  and attn_empcode = emp_code and attn_salary_days >0 and attn_empcode not in (select e_emp_code from monthly_deduction where e_company = " & company_code & "  and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " )"
    

    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       MsgBox ("Deduction Details are not entered for the employee " + payrs.Fields("emp_name") + "  Emp code : " + Str(payrs.Fields("emp_fpcode")))
       payrs.Close
       Exit Sub
    Else
       payrs.Close
    End If
    
''    If emptype_chk = 0 Then
''  ''    sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S' and s_empcode in (select emp_code from emp_mas  where emp_company = " & company_code & "  " & loc & ")"
''        sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat in ('S','M') "
''
''        paydb.Execute sql
''    ElseIf emptype_chk = 1 Then
'' ''     sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W' and s_empcode in (select emp_code from emp_mas  where emp_company = " & company_code & "  " & loc & ")"
''        sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W' "
''        paydb.Execute sql
''    ElseIf emptype_chk = 3 Then
''        sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'R' "
''        paydb.Execute sql
''
''    ElseIf emptype_chk = 2 Then
''        sql = "select * from attn_entry a , emp_mas b  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat"
''        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''        While Not payrs.EOF
''            sql2 = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = '" & payrs("attn_empcat") & "' and s_empcode = '" & payrs("attn_empcode") & "'"
''            paydb.Execute sql2
''            payrs.MoveNext
''        Wend
''        payrs.Close
''    End If
    
        sql = "delete from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text)
        paydb.Execute sql
    
    sql2 = "Select * from emp_salary"
    payrs2.Open sql2, paydb, adOpenDynamic, adLockOptimistic
'''''    If emptype_chk = 0 Then
'''''       sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'S' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat = 'S' and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code  and a.attn_company = b.emp_company"
'''''    Else
'''''       sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'W' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat = 'W' and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code  and a.attn_company = b.emp_company"
'''''    End If
'''    If emptype_chk = 0 Then
'''       sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and e_emp_cat in ('S')   and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat in ('S','M') and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code and a.attn_company = b.emp_company and attn_salary_days > 0  order by attn_empcode"
'''    ElseIf emptype_chk = 1 Then
'''       sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'W' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat = 'W' and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code and a.attn_company = b.emp_company  and attn_salary_days > 0  order by attn_empcode"
'''    ElseIf emptype_chk = 2 Then
'''       sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction c ,emp_mas d where e_company = " & company_code & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and c.e_finyear = " & finyear & "  and c.e_emp_code = d.emp_code and d.emp_code = c.e_emp_code and c.e_company = d.emp_company group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat in ('S','M') and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "   and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code and a.attn_company = b.emp_company and attn_salary_days > 0  order by attn_empcode"
'''    ElseIf emptype_chk = 3 Then
'''       sql = "select * from attn_entry a,emp_voupay_mast b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and e_emp_cat = 'R' and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_empcat = 'R' and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code and a.attn_company = b.emp_company and attn_salary_days > 0  order by attn_empcode"
'''
'''    End If
    
    sql = "select * from attn_entry a,emp_mas b,(select e_emp_code , sum(e_ded_amount) as deduction from monthly_deduction where e_company = " & company_code & " and  e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Val(cmb_year.Text) & " and e_finyear  = " & finyear & " group by e_emp_code) c where a.attn_company = " & company_code & "  and a.attn_finyear  = " & finyear & "  and a.attn_month =  " & cmb_month.ItemData(cmb_month.ListIndex) & "  and a.attn_year = " & Val(cmb_year.Text) & " and a.attn_empcode = b.emp_code and a.attn_empcode = c.e_emp_code and a.attn_company = b.emp_company  and attn_salary_days > 0  order by attn_empcode"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    emp_chk = 0
    

    Dim pf_el_days As Double
    Dim mess_subsidy_days As Double
    While Not payrs.EOF
''       If payrs.Fields("emp_cat") = "W" Then
''          vdaamt = vda
''          fdaamt = fda
''''       Else
''          ''vdaamt = 0
''          ''fdaamt = 0
''       End If
       
''       If payrs.Fields("emp_code") = 10005 Then
''           MsgBox "Wait"
''       End If
       Dim employeecode As Integer
       
       employeecode = payrs.Fields("emp_code")
       
       
       ESI_EL_Amount1 = payrs.Fields("emp_basic") + payrs.Fields("emp_fda")
       ESI_EL_Amount2 = Round(payrs.Fields("emp_grosspay") / 2, 2)
       
       ESIEligibleFor = "B"
       newESIEligible_Amount = 0
       If ESI_EL_Amount1 >= 21000 Or ESI_EL_Amount2 >= 21000 Then
           newESIEligibleYN = "N"
       Else
           newESIEligibleYN = "Y"
       End If
       

       
       
''       If employeecode = 3146 Then
''          MsgBox (employeecode)
''       End If
       Dim otsql As String
       Dim otamt As Double
       otamt = 0
       otsql = "select * from emp_month_otwages where ot_compcode = " & company_code & " and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_year = " & Val(cmb_year.Text) & " and ot_empcode = " & employeecode
       payrs_ot.Open otsql, paydb, adOpenDynamic, adLockOptimistic
       While Not payrs_ot.EOF
          otamt = payrs_ot.Fields("ot_amount")
          payrs_ot.MoveNext
       Wend
       payrs_ot.Close
       pf_el_days = payrs.Fields("attn_work_days") + payrs.Fields("attn_el") + payrs.Fields("attn_dec_holiday")
       
''       mess_subsidy_days = payrs.Fields("attn_work_days") + payrs.Fields("attn_dec_holiday")


       
       
       grosspay = payrs.Fields("emp_basic") + _
              payrs.Fields("emp_splpay") + _
              payrs.Fields("emp_fda") + _
              payrs.Fields("emp_hra") + _
              payrs.Fields("emp_attall") + _
              payrs.Fields("emp_convall") + _
              payrs.Fields("emp_splall") + _
              payrs.Fields("emp_teaall") + _
              payrs.Fields("emp_medall") + _
              payrs.Fields("emp_healthall") + _
              payrs.Fields("emp_washall") + _
              payrs.Fields("emp_lta") + _
              payrs.Fields("emp_magall") + _
              payrs.Fields("emp_fuelall") + _
              payrs.Fields("emp_profall") + _
              payrs.Fields("emp_phoneall") + _
              payrs.Fields("emp_cityall") + _
              payrs.Fields("emp_eduall") + _
              payrs.Fields("emp_mealsall") + _
              payrs.Fields("emp_othall")
       
       ei_esi_ded = 0
       ei_esi_ded2 = 0
       Dim tdsper As Double
       tdsper = 0
''          tdsper = payrs.Fields("emp_tds_per")
          el_attn_all = 0
       
       
         Dim gpay2 As Double
     
       
       Dim pfcalc As Double


            el_gpay = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_hra") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      IIf(el_attn_all > 0, Round(el_attn_all / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0), 0) + _
                      Round(payrs.Fields("emp_convall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_splall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_teaall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_work_days"), 0) + _
                      Round(payrs.Fields("emp_medall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_healthall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_washall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_lta") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_magall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_fuelall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_profall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_phoneall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_cityall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_eduall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_mealsall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                      Round(payrs.Fields("emp_othall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)


''                If grosspay < esi_eligible And payrs.Fields("emp_esieligible") = "Y" Then
''                      ei_esi_ded2 = Round(Round(el_gpay + otamt, 0) * esi_contri / 100, 0)
''                      ei_esi_ded = Round(el_gpay + otamt, 0) * esi_contri / 100
''                      If ei_esi_ded > ei_esi_ded2 Then
''                         ei_esi_ded = Round(Round(el_gpay + otamt, 0) * esi_contri / 100, 0) + 1
''                      End If
''                End If
                

       If newESIEligibleYN = "Y" Then
           If ESI_EL_Amount1 > ESI_EL_Amount2 Then
              ESIEligibleFor = "B"
              newESIEligible_Amount = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
           Else
              ESIEligibleFor = "G"
              newESIEligible_Amount = Round(el_gpay / 2, 0)
           End If
       End If
       
       
       
       

''       el_gpay = 0
''
''       If payrs.Fields("emp_code") = "244" Or payrs.Fields("emp_code") = "276" Or payrs.Fields("emp_code") = "5004" Or payrs.Fields("emp_code") = "363" Then
''          MsgBox ("Test")
''       End If
''       If payrs.Fields("emp_pfeligible") = "Y" Then
''          If emptype_chk = 1 Then
''                pfamount = Round((Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''                   Round(payrs.Fields("emp_serwt") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''                   Round(payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''                   Round(vdaamt / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''                   Round(fdaamt / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2)) * payrs.Fields("emp_pfp") / 100, 0)
''          Else
''                pfcalc = payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days") + payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days")
''                If pfcalc > pfeligible Then
''                   pfcalc = pfeligible
''                End If
''                pfamount = Round(pfcalc * payrs.Fields("emp_pfp") / 100, 0)
''          End If
''       Else
''          pfamount = 0
''       End If



         If payrs.Fields("emp_pfeligible") = "Y" Then
         
         
''          If emptype_chk = 1 Then
''''                pfcalc = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''''                   Round(payrs.Fields("emp_serwt") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''''                   Round(payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''''                   Round(vdaamt / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2) + _
''''                   Round(fdaamt / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2)
''                pfcalc = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * pf_el_days, 0) + _
''                   Round(payrs.Fields("emp_serwt") / payrs.Fields("attn_act_wdays") * pf_el_days, 0) + _
''                   Round(payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * pf_el_days, 0) + _
''                   Round(vdaamt / payrs.Fields("attn_act_wdays") * pf_el_days, 0) + _
''                   Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * pf_el_days, 0)
''
''          Else

 '                pfcalc = payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days") + payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days")
''          End If
                    
''          pfcalc = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * pf_el_days, 0) + _
''                   Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * pf_el_days, 0)
          pfcalc = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0) + _
                   Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
                   
                    
          If pfcalc > pfeligible Then
             pfcalc = pfeligible
          End If
          
          
          pfamount = Round(pfcalc * payrs.Fields("emp_pfp") / 100, 0)
       Else
          pfamount = 0
       End If
       
''       deposit_amt = Round(payrs.Fields("emp_deposit") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       
       
       If emptype_chk = 3 And tdsper > 0 Then
            netpay = Round(Round(el_gpay, 0) - (pfamount + payrs.Fields("emp_lic") + payrs.Fields("emp_rd") + payrs.Fields("emp_houserent") + payrs.Fields("deduction") + payrs.Fields("emp_pfded") + payrs.Fields("emp_teaded") + deposit_amt + payrs.Fields("emp_wfund") + payrs.Fields("emp_bankded") + ei_esi_ded + Round(el_gpay / tdsper, 0)), 0)
       Else
            If payrs.Fields("attn_salary_days") > 0 Then
               If payrs.Fields("attn_work_days") > 0 Then
                  netpay = Round(Round(el_gpay, 0) - (pfamount + payrs.Fields("emp_lic") + payrs.Fields("emp_rd") + payrs.Fields("deduction") + deposit_amt + ei_esi_ded), 0)
               Else
                  netpay = Round(Round(el_gpay, 0) - (pfamount + payrs.Fields("emp_lic") + payrs.Fields("emp_rd") + payrs.Fields("deduction") + deposit_amt + ei_esi_ded), 0)
               End If
            Else
               netpay = Round(Round(el_gpay, 0) - (pfamount) - ei_esi_ded - deposit_amt, 0)
            End If
       End If
       
       netpay = netpay - payrs.Fields("emp_itded") + otamt
        
       ''netpay = netpay + Round(payrs.Fields("emp_mess_subsidy") / payrs.Fields("attn_act_wdays") * mess_subsidy_days, 0)
       payrs2.AddNew
       payrs2.Fields("s_company") = company_code
       payrs2.Fields("s_finyear") = finyear
       payrs2.Fields("s_month") = cmb_month.ItemData(cmb_month.ListIndex)
       payrs2.Fields("s_year") = Val(cmb_year.Text)
       payrs2.Fields("s_empcode") = payrs.Fields("attn_empcode")
       payrs2.Fields("s_empcat") = payrs.Fields("attn_empcat")
       payrs2.Fields("s_emptype") = payrs.Fields("emp_type")
       payrs2.Fields("s_deptcode") = payrs.Fields("emp_dept")
       payrs2.Fields("s_pf_eligible") = payrs.Fields("emp_pfeligible")
       payrs2.Fields("s_esi_eligible") = payrs.Fields("emp_esieligible")
       payrs2.Fields("s_avlworkdays") = payrs.Fields("attn_act_wdays")
       payrs2.Fields("s_actworkdays") = payrs.Fields("attn_work_days")
       payrs2.Fields("s_eligible_leave") = payrs.Fields("attn_el")
       payrs2.Fields("s_per_leave") = payrs.Fields("attn_pl")
       payrs2.Fields("s_absent") = payrs.Fields("attn_abs")
       payrs2.Fields("s_layoff") = payrs.Fields("attn_layoff")
       payrs2.Fields("s_dec_holiday") = payrs.Fields("attn_dec_holiday")
       payrs2.Fields("s_dec_holiday_eligible") = payrs.Fields("attn_dec_holiday_eligible")
       payrs2.Fields("s_medi_leave") = payrs.Fields("attn_ml")
       payrs2.Fields("s_week_off") = payrs.Fields("attn_week_off")
       payrs2.Fields("s_emer_leave") = payrs.Fields("attn_emer_leave_days")
       payrs2.Fields("s_wo_present") = payrs.Fields("attn_wo_present")
       payrs2.Fields("s_total_days") = payrs.Fields("attn_total_days")
       payrs2.Fields("s_eligible_days") = payrs.Fields("attn_eligible_days")
       payrs2.Fields("s_ot_days") = payrs.Fields("attn_ot_days")
       
       payrs2.Fields("s_salarydays") = payrs.Fields("attn_salary_days")
       payrs2.Fields("s_basic") = payrs.Fields("emp_basic")
       payrs2.Fields("s_splpay") = payrs.Fields("emp_splpay")
       
       
       
       
       ''If payrs.Fields("emp_cat") = "W" And payrs.Fields("emp_da_eligible") = "Y" Then
       ''  payrs2.Fields("s_vda") = vdaamt
       ''   payrs2.Fields("s_fda") = fdaamt
      '' Else
      ''    payrs2.Fields("s_vda") = 0
      ''    payrs2.Fields("s_fda") = 0
      '' End If
        payrs2.Fields("s_vda") = vdaamt
       payrs2.Fields("s_fda") = payrs.Fields("emp_fda")
      
       payrs2.Fields("s_hra") = payrs.Fields("emp_hra")
       attend_allowance = payrs.Fields("emp_attall")
       payrs2.Fields("s_attall") = attend_allowance
       payrs2.Fields("s_convall") = payrs.Fields("emp_convall")
       payrs2.Fields("s_splall") = payrs.Fields("emp_splall")
       payrs2.Fields("s_teaall") = payrs.Fields("emp_teaall")
       payrs2.Fields("s_medall") = payrs.Fields("emp_medall")
       payrs2.Fields("s_healthall") = payrs.Fields("emp_healthall")
       payrs2.Fields("s_washall") = payrs.Fields("emp_washall")
       payrs2.Fields("s_lta") = payrs.Fields("emp_lta")
       payrs2.Fields("s_magall") = payrs.Fields("emp_magall")
       payrs2.Fields("s_fuelall") = payrs.Fields("emp_fuelall")
       payrs2.Fields("s_profall") = payrs.Fields("emp_profall")
       payrs2.Fields("s_phoneall") = payrs.Fields("emp_phoneall")
       payrs2.Fields("s_cityall") = payrs.Fields("emp_cityall")
       payrs2.Fields("s_eduall") = payrs.Fields("emp_eduall")
       payrs2.Fields("s_mealsall") = payrs.Fields("emp_mealsall")
       payrs2.Fields("s_othall") = payrs.Fields("emp_othall")
       
       



       
       payrs2.Fields("s_it_ded") = payrs.Fields("emp_itded")
       
       payrs2.Fields("s_grosspay") = Round(grosspay, 2)
       payrs2.Fields("s_eligible_basic") = Round(payrs.Fields("emp_basic") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       ''payrs2.Fields("s_eligible_serwt") = Round(payrs.Fields("emp_serwt") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
     
       payrs2.Fields("s_eligible_splpay") = Round(payrs.Fields("emp_splpay") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
''       If payrs.Fields("attn_empcat") = "W" Then
''          payrs2.Fields("s_eligible_vda") = Round(vdaamt / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2)
''          payrs2.Fields("s_eligible_fda") = Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 2)
''       Else
''          payrs2.Fields("s_eligible_vda") = 0
''          payrs2.Fields("s_eligible_fda") = 0
''       End If
        payrs2.Fields("s_eligible_vda") = 0
       
       payrs2.Fields("s_eligible_fda") = Round(payrs.Fields("emp_fda") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       
       payrs2.Fields("s_eligible_hra") = Round(payrs.Fields("emp_hra") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       
       If el_attn_all > 0 Then
          payrs2.Fields("s_eligible_attall") = Round(el_attn_all / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       Else
          payrs2.Fields("s_eligible_attall") = el_attn_all
       End If
       
       payrs2.Fields("s_eligible_convall") = Round(payrs.Fields("emp_convall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_splall") = Round(payrs.Fields("emp_splall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_teaall") = Round(payrs.Fields("emp_teaall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_work_days"), 0)
       payrs2.Fields("s_eligible_medall") = Round(payrs.Fields("emp_medall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_healthall") = Round(payrs.Fields("emp_healthall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_washall") = Round(payrs.Fields("emp_washall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_lta") = Round(payrs.Fields("emp_lta") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_magall") = Round(payrs.Fields("emp_magall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_fuelall") = Round(payrs.Fields("emp_fuelall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_profall") = Round(payrs.Fields("emp_profall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_phoneall") = Round(payrs.Fields("emp_phoneall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_cityall") = Round(payrs.Fields("emp_cityall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_eduall") = Round(payrs.Fields("emp_eduall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_mealsall") = Round(payrs.Fields("emp_mealsall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       payrs2.Fields("s_eligible_othall") = Round(payrs.Fields("emp_othall") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
''       payrs2.Fields("s_mess_subsidy") = Round(payrs.Fields("emp_mess_subsidy") / payrs.Fields("attn_act_wdays") * mess_subsidy_days, 0)
       payrs2.Fields("s_eligible_grosspay") = Round(el_gpay, 0)
       payrs2.Fields("s_pf") = pfamount
       payrs2.Fields("s_lic") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_lic"), 0)
       Dim myval As Double
       If emptype_chk = 3 And tdsper > 0 Then
          myval = Round((Round(el_gpay, 0) / tdsper), 1)
          If myval - Int(myval) >= 0.5 Then
             myval = Int(myval) + 1
          Else
             myval = Round(myval)
          End If
''          payrs2.Fields("s_rd") = Round((Round(el_gpay, 0) / tdsper), 0)
          payrs2.Fields("s_rd") = myval
       Else
          payrs2.Fields("s_rd") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_rd"), 0)
       End If
       
''       payrs2.Fields("s_deposit") = Round(payrs.Fields("emp_deposit") / payrs.Fields("attn_act_wdays") * payrs.Fields("attn_salary_days"), 0)
       
       ''payrs2.Fields("s_pfded") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_pfded"), 0)
       
       
                 If newESIEligibleYN = "Y" And newESIEligible_Amount > 0 Then
                 
                      ei_esi_ded2 = Round(Round(newESIEligible_Amount + otamt, 0) * esi_contri / 100, 0)
                      ei_esi_ded = Round(newESIEligible_Amount + otamt, 0) * esi_contri / 100
                      If ei_esi_ded > ei_esi_ded2 Then
                         ei_esi_ded = Round(Round(newESIEligible_Amount + otamt, 0) * esi_contri / 100, 0) + 1
                      End If
                End If
                

                ei_esi_ded = Round(ei_esi_ded, 0)
       
       
       
       payrs2.Fields("s_esi_ded") = ei_esi_ded
       
       
       
       
       ''''''''''modified by devaraj on 25.11.2015
''       payrs2.Fields("s_teaded") = IIf(payrs.Fields("attn_work_days") > 0, payrs.Fields("emp_teaded"), 0) ''
''       payrs2.Fields("s_teaded") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_teaded"), 0)
''      payrs2.Fields("s_bankded") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_bankded"), 0)
''       payrs2.Fields("s_wfund") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_wfund"), 0)
''       payrs2.Fields("s_houserent") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("emp_houserent"), 0)
       payrs2.Fields("s_otherdeductions") = IIf(payrs.Fields("attn_salary_days") > 0, payrs.Fields("deduction"), 0)
       payrs2.Fields("s_netpay") = Round(netpay, 0)
       payrs2.Fields("s_salary_bank") = payrs.Fields("emp_bank")
       payrs2.Fields("s_otamount") = Round(otamt, 0)
       payrs2.Fields("s_esi_ded_for_amount") = Round(newESIEligible_Amount, 0)
       
       payrs2.Update

''       If employee_type = 2 Or employee_type = 3 Then
''          If (el + pl + lp1) > 3 Then
''             attn_deduct_days = el + pl + lp1 - 3
''          Else
''             attn_deduct_days = 0
''          End If
''          If ab > 0 Then attn_deduct_days = attn_deduct_days + ab
''          attn_deduct_amt = attn_deduct_days * 5    'per day rs.5 for attendence deduction
''          If attend_allowance > attn_deduct_amt Then
''             attend_allowance = attend_allowance - attn_deduct_amt
''          Else
''             attend_allowance = 0
''          End If
''          If (wd + el + pl) >= aw Then
''             tea_allowance = payrs.Fields("emp_teaall")
''          Else
''             tea_allowance = Round((Round((payrs.Fields("emp_teaall") / aw), 2) * (wd + el + pl)), 2)
''          End If
''       End If
''       earning_amount = payrs2.Fields("s_etotal")
''       dec_holiday_amount = Round(oneday * dh, 2)
''       deduction_amount = Round((ab + lp1 + lp2 + pl) * oneday, 2)
''       earning = earning_amount - deduction_amount + dec_holiday_amount
''       If employee_type = 0 Or employee_type = 2 Then
''          pfamount = Round(earning * 0.12, 0)
''       Else
''          pfamount = 0
''       End If
''       End If
       payrs.MoveNext
    Wend
    Beep
    MsgBox ("Salary / Wages Processing over")
    payrs2.Close
    payrs.Close
    Exit Sub
err_handler:
    
    MsgBox ("Problem in  Employee : " + payrs.Fields("emp_name") + " EMP CODE : " + Str(payrs.Fields("emp_code")))
    
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume


'    Set payrs = Nothing
 '   Set payrs2 = Nothing
End Sub

Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
    Dim mdays, diff As Integer
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If cmb_year.Text = "" Then Exit Sub
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
End Sub

