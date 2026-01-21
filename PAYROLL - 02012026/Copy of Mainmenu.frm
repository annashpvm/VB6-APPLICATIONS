VERSION 5.00
Begin VB.Form temp 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame vda_frame 
      Caption         =   "VARIABLE DEARNESS ALLOWANCE ENTRY"
      Height          =   5325
      Left            =   1395
      TabIndex        =   0
      Top             =   1590
      Visible         =   0   'False
      Width           =   9330
      Begin VB.CommandButton exit2 
         Caption         =   "&Exit"
         Height          =   1095
         Left            =   3825
         TabIndex        =   9
         Top             =   3990
         Width           =   1230
      End
      Begin VB.CommandButton refresh 
         Caption         =   "&Refresh"
         Height          =   1095
         Left            =   2340
         TabIndex        =   8
         Top             =   4005
         Width           =   1230
      End
      Begin VB.CommandButton Save 
         Caption         =   "&Save"
         Height          =   1095
         Left            =   795
         TabIndex        =   7
         Top             =   3990
         Width           =   1230
      End
      Begin VB.TextBox vda_amount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   4200
         TabIndex        =   6
         Top             =   2040
         Width           =   2550
      End
      Begin VB.ComboBox month_cmb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2445
         TabIndex        =   2
         Top             =   705
         Width           =   2655
      End
      Begin VB.ComboBox year_cmb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7110
         TabIndex        =   1
         Top             =   675
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   15
         X2              =   9330
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Label Label3 
         Caption         =   "VDA AMOUNT"
         Height          =   480
         Left            =   735
         TabIndex        =   5
         Top             =   2280
         Width           =   2835
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9315
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "YEAR"
         Height          =   285
         Left            =   5955
         TabIndex        =   4
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
         Height          =   330
         Left            =   720
         TabIndex        =   3
         Top             =   750
         Width           =   1200
      End
   End
End
Attribute VB_Name = "temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub att_status_entry_Click()
   name1 = "ATTENDANCE STATUS MASTER ENTRY"
   name2 = "STATUS NAME"
   name3 = "SELECT STATUS NAME"
   name4 = "STATUS MODIFICATION"
   name5 = "STATUS DELETION"
   menuchk = 9
   master.Show
End Sub

Private Sub ATTN_REP_Click()
   at_rep_opt = 0
   attn_reports_frm.Show
End Sub

Private Sub attn_sum_Click()
   at_rep_opt = 1
   attn_reports_frm.Show
End Sub



Private Sub bonus_st_Click()
   Bonus_statement.Show
End Sub

Private Sub dec_holiday_mas_Click()
   dec_holiday.Show
End Sub


Private Sub month_deduction_worker_Click()
     emptype_chk = 1
     month_deduct.Show
End Sub

Private Sub pay_slip_Click()
     pay_slip_print.Show
End Sub

Private Sub pf_reports_Click()
    pf_reports_frm.Show
End Sub

Private Sub prod_incentive_Click()
    prodn_incentive_frm.Show
End Sub

Private Sub S_ATTENT_Click()
    emptype_chk = 0
    atten_ent.Show
End Sub

Private Sub CASTEMAS_Click()
     name1 = "CASTE MASTER ENTRY"
     name2 = "CASTE NAME"
     name3 = "SELECT CASTE NAME"
     name4 = "CASTE MODIFICATION"
     name5 = "CASTE DELETION"
     menuchk = 7
     master.Show
End Sub
Private Sub COMMMAS_Click()
     name1 = "COMMUNITY MASTER ENTRY"
     name2 = "COMMUNITY NAME"
     name3 = "SELECT COMMUNITY NAME"
     name4 = "COMMUNITY MODIFICATION"
     name5 = "COMMUNITY DELETION"
     menuchk = 6
     master.Show
End Sub

Private Sub DEDUMAS_Click()
     name1 = "DEDUCTION MASTER ENTRY"
     name2 = "DEDUCTION NAME"
     name3 = "SELECT DEDUCTION NAME"
     name4 = "DEDUCTION MASTER MODIFICATION"
     name5 = "DEDUCTION MASTER DELETE"
     menuchk = 8
     master.Show
End Sub

Private Sub DEPTMASTER_Click()
     name1 = "DEPARTMENT MASTER ENTRY"
     name2 = "DEPARTMENT NAME"
     name3 = "SELECT DEPARTMENT NAME"
     name4 = "DEPARTMENT MASTER MODIFICATION"
     name5 = "DEPARTMENT MASTER DELETE"
     menuchk = 1
     master.Show
End Sub

Private Sub DESIMAS_Click()
     name1 = "EMPLOYEE DESIGNATION MASTER ENTRY"
     name2 = "EMPLOYEE DESIGNATION NAME"
     name3 = "SELECT EMPLOYEE DESIGNATION NAME"
     name4 = "EMPLOYEE DESIGNATION MODIFICATION"
     name5 = "EMPLOYEE DESIGNATION DELETION"
     menuchk = 3
     master.Show
End Sub

Private Sub EMPMAS_Click()
     Load emp_mas_entry
     emp_mas_entry.ZOrder
     emp_mas_entry.Show
End Sub

Private Sub EMPTYPMAS_Click()
     name1 = "EMPLOYEE TYPE MASTER ENTRY"
     name2 = "EMPLOYEE TYPE NAME"
     name3 = "SELECT EMPLOYEE TYPE NAME"
     name4 = "EMPLOYEE TYPE MODIFICATION"
     name5 = "EMPLOYEE TYPE DELETION"
     menuchk = 2
     master.Show
End Sub

Private Sub exit_Click()
     Unload Me
End Sub

Private Sub exit2_Click()
    vda_frame.Visible = False
End Sub

Private Sub Form_Load()
    emptype_chk = 0
    Mainmenu.Caption = Mainmenu.Caption & "   -   " & millname
End Sub
Private Sub month_cmb_Click()
    If Trim(month_cmb.Text) = "" Then
       MsgBox ("Select month")
       Exit Sub
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
    paydb.Open pay
    sql = ("select * from emp_vda where v_year = " & Trim(year_cmb.Text) & " and v_month = " & month_cmb.ItemData(month_cmb.ListIndex) & " and v_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       vda_amount = payrs.Fields("v_vdaamount")
    Else
       vda_amount = 0
    End If
End Sub

Private Sub MONTH_DEDUCT_ENTRY_Click()
     emptype_chk = 0
     month_deduct.Show
End Sub

Private Sub pay_calc_staff_Click()
     pay_calchk = 0
     pay_cal.Show
End Sub

Private Sub pay_calc_worker_Click()
     pay_calchk = 1
     pay_cal.Show
End Sub

Private Sub pf_nominee_mas_Click()
   pf_nominee.Show
End Sub

Private Sub QLAMAS_Click()
     name1 = "QUALIFICATION MASTER ENTRY"
     name2 = "QUALIFICATION NAME"
     name3 = "SELECT QUALIFICATION NAME"
     name4 = "QUALIFICATION MODIFICATION"
     name5 = "QUALIFICATION DELETION"
     menuchk = 4
     master.Show
End Sub
Private Sub RELIGIONMAS_Click()
     name1 = "RELIGION  MASTER ENTRY"
     name2 = "RELIGION NAME"
     name3 = "SELECT RELIGION  NAME"
     name4 = "RELIGION  MODIFICATION"
     name5 = "RELIGION  DELETION"
     menuchk = 5
     master.Show
End Sub


Private Sub sal_arrear_entry_Click()
    salary_arrear_entry.Show
End Sub

Private Sub sal_st_Click()
    Salary_statement_prt.Show
End Sub

Private Sub salary_statement_Click()
    salary_summary_st.Show
End Sub

Private Sub save_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_vda where v_year = " & Trim(year_cmb.Text) & " and v_month = " & month_cmb.ItemData(month_cmb.ListIndex) & " and v_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       payrs.Fields("v_vdaamount") = vda_amount
    Else
       payrs.AddNew
       payrs.Fields("v_company") = "1"
       payrs.Fields("v_month") = month_cmb.ItemData(month_cmb.ListIndex)
       payrs.Fields("v_year") = Trim(year_cmb.Text)
       payrs.Fields("v_vdaamount") = vda_amount
       payrs.Update
       vda_amount = 0
    End If
End Sub

Private Sub vda_amount_KeyPress(KeyAscii As Integer)
  On Error GoTo err_handler
    chk_keyascii vda_amount, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
Private Sub VDA_AMT_ENTRY_Click()
     vda_frame.Visible = True
     With month_cmb
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
        .AddItem "Auguest"
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
    With year_cmb
        .AddItem "2002"
        .AddItem "2003"
        .AddItem "2004"
        .AddItem "2005"
        .AddItem "2006"
        .AddItem "2007"
        .AddItem "2008"
        .AddItem "2009"
        .AddItem "2010"
        .AddItem "2011"
        .AddItem "2012"
        .AddItem "2013"
    End With
    year_cmb.Text = "2002"
End Sub

Private Sub W_ATTEN_Click()
  emptype_chk = 1
  atten_ent.Show
End Sub

