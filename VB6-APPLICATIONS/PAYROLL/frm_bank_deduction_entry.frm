VERSION 5.00
Begin VB.Form frm_bank_deduction_entry 
   Caption         =   "BANK DEDUCTION ENTRY"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE DEDUCTION DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6945
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   10755
      Begin VB.Frame Frame5 
         Caption         =   "DETAILS FOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   2280
         TabIndex        =   28
         Top             =   480
         Width           =   6045
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   1080
            TabIndex        =   30
            Top             =   240
            Width           =   2220
         End
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   3600
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5010
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   10305
         Begin VB.TextBox desi 
            Enabled         =   0   'False
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
            Left            =   8115
            TabIndex        =   22
            Top             =   1335
            Width           =   2040
         End
         Begin VB.TextBox dept 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2820
            TabIndex        =   21
            Top             =   1290
            Width           =   3540
         End
         Begin VB.TextBox emp_idcode 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8100
            TabIndex        =   20
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox emptype 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2820
            TabIndex        =   19
            Top             =   750
            Width           =   3525
         End
         Begin VB.ComboBox empname_cmb 
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
            Left            =   2835
            TabIndex        =   18
            Top             =   285
            Width           =   7170
         End
         Begin VB.Frame Frame4 
            Height          =   2655
            Left            =   840
            TabIndex        =   6
            Top             =   1680
            Width           =   8655
            Begin VB.TextBox txt_can_acno 
               Height          =   375
               Left            =   2760
               TabIndex        =   12
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txt_fest_loan_acno 
               Height          =   375
               Left            =   2760
               TabIndex        =   11
               Top             =   1440
               Width           =   2295
            End
            Begin VB.TextBox txt_othloan_acno 
               Height          =   375
               Left            =   2760
               TabIndex        =   10
               Top             =   2040
               Width           =   2295
            End
            Begin VB.TextBox txt_canbud_amt 
               Height          =   375
               Left            =   5520
               TabIndex        =   9
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txt_festloan_amt 
               Height          =   375
               Left            =   5520
               TabIndex        =   8
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox txt_othloan_amt 
               Height          =   375
               Left            =   5520
               TabIndex        =   7
               Top             =   2040
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Can Budget A/C Number "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   240
               TabIndex        =   17
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label Label2 
               Caption         =   "Festivel Loan A/C Number "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   1560
               Width           =   2415
            End
            Begin VB.Label Label8 
               Caption         =   "Other Loan A/C Number "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   240
               TabIndex        =   15
               Top             =   2160
               Width           =   2415
            End
            Begin VB.Label Label9 
               Caption         =   " A/C Number "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   3360
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Amount "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   5760
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Label Label7 
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6540
            TabIndex        =   27
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   225
            TabIndex        =   26
            Top             =   1335
            Width           =   2370
         End
         Begin VB.Label Label5 
            Caption         =   "Emp. Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   6540
            TabIndex        =   25
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label Label4 
            Caption         =   "Employee Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   225
            TabIndex        =   24
            Top             =   810
            Width           =   2145
         End
         Begin VB.Label Label3 
            Caption         =   "Employee Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   225
            TabIndex        =   23
            Top             =   360
            Width           =   2280
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   7560
      Width           =   2055
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   825
         Left            =   120
         Picture         =   "frm_bank_deduction_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   825
         Left            =   1080
         Picture         =   "frm_bank_deduction_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label lbl_disp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4650
      TabIndex        =   3
      Top             =   7050
      Width           =   4815
   End
End
Attribute VB_Name = "frm_bank_deduction_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim emp_chk As Integer
Dim emptypecode As Integer
Dim blank_rec_upd As Integer
Public nextname As String

Private Sub empname_cmb_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    If opt_staff.Value = True Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat = 'S' ")
    Else
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat = 'W' ")
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
    Else
       emp_idcode = payrs.Fields("emp_code")
       find_deptname (payrs.Fields("emp_dept"))
       dept.Text = dname
       find_desiname (payrs.Fields("emp_design"))
       DESI.Text = dname
       find_etypename (payrs.Fields("emp_type"))
       emptype.Text = dname
       emptypecode = payrs.Fields("emp_type")
       txt_can_acno.Text = payrs.Fields("emp_canbud_ac")
       txt_canbud_amt.Text = payrs.Fields("emp_canbud_amt")
       txt_fest_loan_acno.Text = payrs.Fields("emp_festloan_ac")
       txt_festloan_amt.Text = payrs.Fields("emp_festloan_amt")
       txt_othloan_acno.Text = payrs.Fields("emp_loan_ac")
       txt_othloan_amt.Text = payrs.Fields("emp_loan_amt")
    End If
    payrs.Close
End Sub

Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
    loc = ""
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If opt_staff.Value = True Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_status = 'A' order by emp_name")
    Else
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_status = 'A' order by emp_name")
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    emp_chk = 0
    While Not payrs.EOF
        empname_cmb.AddItem payrs("emp_name")
        payrs.MoveNext
        emp_chk = emp_chk + 1
    Wend
End Sub

Private Sub opt_staff_Click()
   load_empdata
End Sub
Public Function load_empdata()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If opt_staff.Value = True Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_status = 'A' order by emp_name")
    Else
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_status = 'A' order by emp_name")
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    emp_chk = 0
    While Not payrs.EOF
        empname_cmb.AddItem payrs("emp_name")
        payrs.MoveNext
        emp_chk = emp_chk + 1
    Wend
End Function

Private Sub opt_worker_Click()
   load_empdata
End Sub

Private Sub SAVE_Click()
On Error GoTo err_handler
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  sql = "update emp_mas set emp_canbud_ac = '" & txt_can_acno.Text & "' , emp_festloan_ac = '" & txt_fest_loan_acno.Text & "' ,emp_loan_ac = '" & txt_othloan_acno.Text & "' , emp_canbud_amt = " & Val(txt_canbud_amt.Text) & ",emp_festloan_amt= " & Val(txt_festloan_amt.Text) & " ,emp_loan_amt = " & Val(txt_othloan_amt.Text) & " where emp_name = '" & Trim(empname_cmb.Text) & "' and  emp_code = '" & Trim(emp_idcode.Text) & "' and  emp_company = " & company_code & ""
  paydb.Execute sql
  MsgBox ("Record updated")
  emptype.Text = ""
  emp_idcode.Text = ""
  dept.Text = ""
  DESI.Text = ""
  txt_can_acno.Text = ""
  txt_canbud_amt.Text = ""
  txt_fest_loan_acno.Text = ""
  txt_festloan_amt.Text = ""
  txt_othloan_acno.Text = ""
  txt_othloan_amt.Text = ""
  Exit Sub
err_handler:
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub


Public Function process_data()
  Dim sql2 As String
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  sql = "select *  from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and attn_year = " & Trim(cmb_year.Text) & " and attn_empcode not in (select e_emp_code  from monthly_deduction where e_company = " & company_code & " and e_finyear = " & finyear & "  and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Trim(cmb_year.Text) & ")"
  paydb.Execute sql
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  While Not payrs.EOF
        sql2 = "insert into monthly_deduction values ( " & payrs("attn_company") & " , " & payrs("attn_finyear") & " , '" & payrs("attn_empcode") & "' , '" & payrs("attn_empcat") & "' , " & payrs("attn_year") & " , " & payrs("attn_month") & " , 1, 0 , " & payrs("attn_act_wdays") & " )"
        paydb.Execute sql2
     payrs.MoveNext
  Wend
  payrs.Close
  
  sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  If Not payrs.EOF Then
       If payrs("pay_dedu_lock") = "Y" Then
          save.Enabled = False
          lbl_disp.Caption = "Deduction Locked .. Can't Modify"
       End If
  Else
       lbl_disp.Caption = ""
       save.Enabled = True
  End If
  payrs.Close
End Function



Private Sub Text3_Change()

End Sub

Private Sub txt_canbud_amt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii txt_canbud_amt, "N", 7, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_festloan_amt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii txt_festloan_amt, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_othloan_amt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii txt_othloan_amt, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
