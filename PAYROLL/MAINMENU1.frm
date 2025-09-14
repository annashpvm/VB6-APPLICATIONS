VERSION 5.00
Begin VB.MDIForm MAINMENU 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu master_entry 
      Caption         =   "MASTER      "
      Begin VB.Menu DEPTMASTER 
         Caption         =   "DEPARTMENT MASTER"
      End
      Begin VB.Menu EMPTYPMAS 
         Caption         =   "EMPLOYEE TYPE MASTER"
      End
      Begin VB.Menu DESIMAS 
         Caption         =   "DESIGNATION MASTER"
      End
      Begin VB.Menu QLAMAS 
         Caption         =   "QUALIFICATION MASTER"
      End
      Begin VB.Menu RELIGIONMAS 
         Caption         =   "RELIGION MASTER"
      End
      Begin VB.Menu COMMAS 
         Caption         =   "COMMUNITY MASTER"
      End
      Begin VB.Menu CASTEMAS 
         Caption         =   "CASTE ENTRY"
      End
      Begin VB.Menu att_status_entry 
         Caption         =   "ATTENANCE STATUS"
      End
      Begin VB.Menu DEDUMAS 
         Caption         =   "DEDUCTION MASTER"
      End
      Begin VB.Menu dec_holiday_mas 
         Caption         =   "DECLARE HOLIDAY MASTER"
      End
   End
   Begin VB.Menu EMP_MASTER 
      Caption         =   "   EMPLOYEE MASTER    "
      Begin VB.Menu empmas 
         Caption         =   "EMPLOYEE DETAILS ENTRY"
      End
      Begin VB.Menu pf_nominee_mas 
         Caption         =   "PF NOMINEE MASTER "
      End
   End
   Begin VB.Menu DATA_ENTRY 
      Caption         =   "   DATA ENTRY      "
      Begin VB.Menu S_ATTENT 
         Caption         =   "DAILY ATTENDANCE ENTRY FOR STAFF"
      End
      Begin VB.Menu W_ATTEN 
         Caption         =   "DAILY ATTENDANCE ENTRY FOR WORKER"
      End
      Begin VB.Menu month_deduct_staff 
         Caption         =   "MONTHLY DEDUCTION ENTRY FOR STAFF"
      End
      Begin VB.Menu month_deduct_worker 
         Caption         =   "MONTHLY DEDUCTION ENTRY FOR WORKER"
      End
      Begin VB.Menu vda_amt_entry 
         Caption         =   "VDA AMOUNT ENTRY"
      End
      Begin VB.Menu sal_arrear_entry 
         Caption         =   "SALARY ARREAR ENTRY"
      End
   End
   Begin VB.Menu pay_calculation 
      Caption         =   "   PAY CALCULATION      "
      Begin VB.Menu pay_calc_staff 
         Caption         =   "FOR STAFF"
      End
      Begin VB.Menu pay_calc_worker 
         Caption         =   "FOR WORKER"
      End
   End
   Begin VB.Menu reports 
      Caption         =   "   REPORTS      "
      Begin VB.Menu pay_slip 
         Caption         =   "PAY SLIP PRINTING"
      End
      Begin VB.Menu sal_st 
         Caption         =   "SALARY STATEMENT"
      End
      Begin VB.Menu pf_reports 
         Caption         =   "PF REPORTS"
      End
      Begin VB.Menu Prod_incentive 
         Caption         =   "PRODUCTION INCENTIVE REPORTS"
      End
      Begin VB.Menu salary_summary_reports 
         Caption         =   "SALARY SUMMARY REPORTS"
      End
      Begin VB.Menu attn_summary 
         Caption         =   "ATTENDANCE SUMMARY REPORTS"
      End
      Begin VB.Menu bonus_st 
         Caption         =   "BONUS STATEMENT"
      End
   End
   Begin VB.Menu windows 
      Caption         =   "   WINDOWS      "
   End
   Begin VB.Menu exit 
      Caption         =   "   EXIT     "
   End
End
Attribute VB_Name = "MAINMENU"
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

Private Sub attn_summary_Click()
   at_rep_opt = 1
   attn_reports_frm.Show
End Sub


Private Sub bonus_st_Click()
   Bonus_statement.Show
End Sub

Private Sub dec_holiday_mas_Click()
   dec_holiday.Show
End Sub
Private Sub MDIForm_Load()
   MAINMENU.Caption = millname
End Sub

Private Sub month_deduct_worker_Click()
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
    MAINMENU.Caption = MAINMENU.Caption & "   -   " & millname
End Sub

Private Sub MONTH_DEDUCT_STAFF_Click()
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

Private Sub VDA_AMT_ENTRY_Click()
     VDA_ENTRY.Show
End Sub

Private Sub W_ATTEN_Click()
  emptype_chk = 1
  atten_ent.Show
End Sub



