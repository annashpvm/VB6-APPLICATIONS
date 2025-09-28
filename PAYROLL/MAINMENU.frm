VERSION 5.00
Begin VB.MDIForm MAINMENU 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   1770
   ClientWidth     =   4680
   Icon            =   "MAINMENU.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu master_entry 
      Caption         =   "MASTERS"
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
      Begin VB.Menu mnu_bank 
         Caption         =   "BANK MASTER"
      End
      Begin VB.Menu MILLS_ENTRY 
         Caption         =   "MILLS DETAILS ENTRY"
      End
   End
   Begin VB.Menu EMP_MASTER 
      Caption         =   "   EMPLOYEE MASTER    "
      Begin VB.Menu empmas 
         Caption         =   "EMPLOYEE DETAILS ENTRY"
      End
      Begin VB.Menu mnu_emp_modifications 
         Caption         =   "EMPLOYEE DETAILS MODIFICATIONS"
      End
      Begin VB.Menu mnu_vou_payment 
         Caption         =   "VOUCHER PAYMENT MASTER"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_emp_modify 
         Caption         =   "EMPLOYEE MODIFICATION"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_emp_salary_slot 
         Caption         =   "EMPLOYEE SALARY SLOT ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_emp_modification 
         Caption         =   "EMPLOYEE WEEK OFF MODIFICATION"
      End
      Begin VB.Menu mnu_emp_worked_posion 
         Caption         =   "EMPLOYEE WORKED POISTION"
      End
      Begin VB.Menu pf_nominee_mas 
         Caption         =   "PF NOMINEE MASTER "
         Visible         =   0   'False
      End
      Begin VB.Menu employee_position_master 
         Caption         =   "EMPLOYEE POSITION MASTER"
         Begin VB.Menu mnu_production 
            Caption         =   "Production"
         End
         Begin VB.Menu mnu_mechanical 
            Caption         =   "Mechanical"
         End
      End
      Begin VB.Menu LINE1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_bank_deduction_additions 
         Caption         =   "EMPLOYEE BANK DEDUCTION & ADDITIONS ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Eligible_leave_staff 
         Caption         =   "Eligible Leave Entry for Staff"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_eligible_leave_worker 
         Caption         =   "Eligible Leave Entry for Worker"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_eligible_leave_others 
         Caption         =   "Eligible Leave Entry for Management"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_eligible_leave_retainer 
         Caption         =   "Eligible Leave Entry for Retainer"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu DATA_ENTRY 
      Caption         =   "   DATA ENTRY      "
      Begin VB.Menu mnu_biomaster_upd 
         Caption         =   "MASTER UPLOAD FROM BIO-METRIC"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bio_upload 
         Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRICS - OLD"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bio_upload_new 
         Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRICS"
         Visible         =   0   'False
      End
      Begin VB.Menu bio 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_attn_staff_entry 
         Caption         =   "ATTENDANCE ENTRY "
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu month_deduct_staff 
         Caption         =   "STAFF - MONTHLY DEDUCTION  (ALL DEDUCTIONS)"
      End
      Begin VB.Menu month_deduct_worker 
         Caption         =   "WORKER - MONTHLY DEDUCTION ENTRY  (ALL DEDUCTIONS)"
      End
      Begin VB.Menu mnu_indi_deductions_staff 
         Caption         =   "DEDUCTION ENTRY"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ot_import 
         Caption         =   "Over Time Import"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_indi_deductions_worker 
         Caption         =   "WORKER - INDIVIDUAL DEDUCTIONS"
         Visible         =   0   'False
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_salary_wages_adv_ent 
         Caption         =   "SALARY / WAGES ADVANCE ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu vda_amt_entry 
         Caption         =   "VDA AMOUNT ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu sal_arrear_entry 
         Caption         =   "SALARY ARREAR ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu LINE 
         Caption         =   "------------------  OTHER ENTRIES -----------------------------"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_emp_overtime_entry 
         Caption         =   "EMPLOYEE PRODUCTION INCENTIVE ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_employee_additional 
         Caption         =   "STAFF ADDITIONAL AMOUNT ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_worker_Addn_amt 
         Caption         =   "WORKER ADDITIONAL AMOUNT ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bank_deduction_entry 
         Caption         =   "EMPLOYEE BANK DEDUCTION ENTRY"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_retainer_entry 
         Caption         =   "RETAINER / VOUCHER PAYMENT ENTRY"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu pay_calculation 
      Caption         =   "   PAY CALCULATION      "
   End
   Begin VB.Menu reports 
      Caption         =   "   REPORTS      "
      Begin VB.Menu mnu_bio_attend 
         Caption         =   "Bio-metric Attendance Reports"
         Visible         =   0   'False
         Begin VB.Menu mnu_bio_vew 
            Caption         =   "View Details"
         End
         Begin VB.Menu mnu_bio_reports 
            Caption         =   "Monthly Reports "
         End
         Begin VB.Menu mnu_daily_reports 
            Caption         =   "Daily Reports"
         End
      End
      Begin VB.Menu rep 
         Caption         =   "-"
      End
      Begin VB.Menu pay_slip 
         Caption         =   "PAY SLIP PRINTING "
         Visible         =   0   'False
      End
      Begin VB.Menu payslip_portrait 
         Caption         =   "PAYSLIP PRINTING"
      End
      Begin VB.Menu mnu_payslip_worker 
         Caption         =   "PAYSLIP WORKER"
         Visible         =   0   'False
      End
      Begin VB.Menu sal_st 
         Caption         =   "SALARY STATEMENT"
      End
      Begin VB.Menu mnu_cost_report 
         Caption         =   "COST REPORT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_present_days 
         Caption         =   "PRESENT DAYS -WORKER(AGEWISE)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_deduction_statement 
         Caption         =   "DEDUCTION STATEMENT"
      End
      Begin VB.Menu pf_reports 
         Caption         =   "PF REPORTS"
      End
      Begin VB.Menu mnu_esi_reports 
         Caption         =   "ESI REPORTS"
      End
      Begin VB.Menu Prod_incentive 
         Caption         =   "OVER TIME REPORTS"
         Visible         =   0   'False
      End
      Begin VB.Menu salary_summary_reports 
         Caption         =   "SALARY SUMMARY REPORTS"
         Visible         =   0   'False
      End
      Begin VB.Menu ATTN_REP 
         Caption         =   "ATTENDANCE REPORT"
      End
      Begin VB.Menu attn_summary 
         Caption         =   "ATTENDANCE SUMMARY REPORTS"
      End
      Begin VB.Menu worked_layoff_monthwise 
         Caption         =   "WORKED &  LOP DAYS -MONTHWISE"
         Visible         =   0   'False
      End
      Begin VB.Menu bonus_st 
         Caption         =   "BONUS STATEMENT"
      End
      Begin VB.Menu mnu_cl_wages 
         Caption         =   "CASUAL LEAVE  WAGES / ATTENDANCE"
      End
      Begin VB.Menu mnu_Rep_overtime 
         Caption         =   "OVER TIME REPORTS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bank_statements 
         Caption         =   "BANK STATEMENTS"
      End
      Begin VB.Menu mnu_emp_list_join 
         Caption         =   "EMPLOYEE LIST FOR JOINING / RESIGNED"
      End
      Begin VB.Menu mnu_emp_details 
         Caption         =   "EMPLOYEE DETAILS"
      End
      Begin VB.Menu mnu_retirement_details_statement 
         Caption         =   "RETIREMENT DETAILS STATEMENT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_gratuity_settlement 
         Caption         =   "GRATUITY SETTLEMENT -INDIVIDUAL"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_gratuity 
         Caption         =   "GRATUITY STATEMENT"
         Visible         =   0   'False
      End
      Begin VB.Menu MNU_GROSS_PAY_STATEMENT 
         Caption         =   "GROSS PAY STATEMENT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_master_details 
         Caption         =   "MASTER DETAILS"
      End
   End
   Begin VB.Menu mnu_applications 
      Caption         =   "APPLICATION DETAILS"
      Begin VB.Menu mnu_application_inward 
         Caption         =   "Applications Inward Entry"
      End
      Begin VB.Menu mnu_application_reports 
         Caption         =   "Application Reports"
      End
   End
   Begin VB.Menu mnu_allcontrols 
      Caption         =   "  ALL CONTROLS"
      Begin VB.Menu mnu_attn_lock 
         Caption         =   "Attendenace Lock"
      End
      Begin VB.Menu mnu_deduction_lock 
         Caption         =   "Deduction Lock"
      End
      Begin VB.Menu mnu_salary_process_lock 
         Caption         =   "Salary process Lock"
      End
      Begin VB.Menu mnu_password_change 
         Caption         =   "Password Change"
      End
   End
   Begin VB.Menu windows 
      Caption         =   "   WINDOWS      "
      WindowList      =   -1  'True
      Begin VB.Menu mill_change 
         Caption         =   "MILL CHANGE"
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
      Begin VB.Menu calculator 
         Caption         =   "CALCULATOR"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_break 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_emp_master_upd 
         Caption         =   "Employee Master Updation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_excel_fileupload 
         Caption         =   "Excel file upload"
      End
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
   attn_monthwise.Show
End Sub

Private Sub attn_summary_Click()
   at_rep_opt = 1
   attn_reports_frm.Show
End Sub


Private Sub bonus_st_Click()
   Bonus_statement.Show
End Sub

Private Sub calculator_Click()
    Dim pvt_calc As Variant
    pvt_calc = Shell("calc.exe", vbMaximizedFocus)
    Exit Sub
End Sub


Private Sub CS_ATTN_Click()
    cbeattn_entry.Show
End Sub

''Private Sub employee_position_master_Click()
'' emp_mas_position.Show
''End Sub

Private Sub MILLS_ENTRY_Click()
   master_company.Show
End Sub

Private Sub dec_holiday_mas_Click()
   dec_holiday.Show
End Sub
Private Sub MDIForm_Load()
   optchk = 0
   MAINMENU.Caption = millname + " - " + fyear
   If uname = "HR-HOD" Then
      mnu_allcontrols.Visible = True
   Else
      mnu_allcontrols.Visible = False
   End If
   
   If userrights = 7 Then
      EMP_MASTER.Visible = False
      DATA_ENTRY.Visible = False
      pay_calculation.Visible = False
   Else
      EMP_MASTER.Visible = True
      DATA_ENTRY.Visible = True
      pay_calculation.Visible = True
   End If
          loc = ""
       loc2 = ""
   If data_source = "A" Then
       loc = ""
       loc2 = ""
   ElseIf data_source = "H" Then
       
   Else
       loc = " and emp_workplace = 'MILL'"
       loc2 = " and s_workplace = 'MILL'"
   End If
''   If hod <> True Then
''      MILLS_ENTRY.Visible = False
''   End If
End Sub

Private Sub mill_change_Click()
    If MAINMENU.ActiveForm Is Nothing Then
        Unload Me
        Load frm_login
        frm_login.Show

    Else
        MsgBox "Close All The Opened Forms", vbInformation, "Message"
    End If
    Exit Sub
End Sub

Private Sub mnu_application_inward_Click()
    frm_applicatoin_inward.Show
    frm_applicatoin_inward.ZOrder
End Sub

Private Sub mnu_application_reports_Click()
    frm_applicatoin_reports.Show
    frm_applicatoin_reports.ZOrder
End Sub

Private Sub mnu_attn_lock_Click()
    control_lock = 1
    frm_control_lock.Show
    frm_control_lock.ZOrder
End Sub

Private Sub mnu_attn_management_Click()
    emptype_chk = 2
''    millattn_entry.Show
    millattn_entry_new.Show
End Sub

Private Sub mnu_attn_retainer_Click()
    emptype_chk = 3
''    millattn_entry.Show
    millattn_entry_new.Show
End Sub

Private Sub mnu_attn_staff_entry_Click()
    emptype_chk = 0
    millattn_entry_new.Show
End Sub

Private Sub mnu_attn_worker_entry_Click()
  emptype_chk = 1
  millattn_entry_new.Show
End Sub

Private Sub mnu_bank_Click()
    frm_bank_master.Show
End Sub

Private Sub mnu_bank_deduction_additions_Click()
    frm_bank_deduction_entry.Show
    frm_bank_deduction_entry.ZOrder
End Sub

Private Sub mnu_bank_deduction_entry_Click()
    frm_bank_loan_monthly_deduction.Show
    frm_bank_loan_monthly_deduction.ZOrder
End Sub

Private Sub mnu_bank_statements_Click()
    frm_bank_statement.Show
    frm_bank_statement.ZOrder

End Sub

Private Sub mnu_bio_reports_Click()
   frm_rep_bio_metrics.Show
   frm_rep_bio_metrics.ZOrder
End Sub

Private Sub mnu_bio_upload_Click()
    frm_bio_metric_upload.Show
    frm_bio_metric_upload.ZOrder
End Sub

Private Sub mnu_bio_upload_new_Click()
    frm_bio_metric_upload_new.Show
    frm_bio_metric_upload_new.ZOrder
End Sub

Private Sub mnu_bio_vew_Click()
   frm_rep_bio_metrics_view.Show
   frm_rep_bio_metrics_view.ZOrder
End Sub

Private Sub mnu_biomaster_upd_Click()
   emp_updation
End Sub

Private Sub mnu_cl_wages_Click()
   cl_wages_statement.Show
   cl_wages_statement.ZOrder
End Sub

Private Sub mnu_cost_report_Click()
    frm_rep_cost_report.Show
    frm_rep_cost_report.ZOrder
End Sub

Private Sub mnu_daily_reports_Click()
    frm_rep_bio_metrics_daily.Show
    frm_rep_bio_metrics_daily.ZOrder
End Sub

Private Sub mnu_deduction_lock_Click()
    control_lock = 2
    frm_control_lock.Show
    frm_control_lock.ZOrder
End Sub

Private Sub mnu_deduction_statement_Click()
    frm_deduction_statement.Show
    frm_deduction_statement.ZOrder
End Sub

Private Sub mnu_eligible_leave_others_Click()
    emptype_chk = 2
    frm_eligible_leave_entry.Show
End Sub

Private Sub mnu_eligible_leave_retainer_Click()
    emptype_chk = 3
    frm_eligible_leave_entry.Show
End Sub

Private Sub mnu_Eligible_leave_staff_Click()
    emptype_chk = 0
    frm_eligible_leave_entry.Show
End Sub

Private Sub mnu_eligible_leave_worker_Click()
    emptype_chk = 1
    frm_eligible_leave_entry.Show
End Sub

Private Sub mnu_emp_details_Click()
    frm_rep_employee_details.Show
    frm_rep_employee_details.ZOrder
End Sub

Private Sub mnu_emp_list_join_Click()
    frm_rep_employee_details_view.Show
    frm_rep_employee_details_view.ZOrder
End Sub

Private Sub mnu_emp_modification_Click()
    emp_mas_modification.Show
    emp_mas_modification.ZOrder
End Sub

Private Sub mnu_emp_modifications_Click()
    emp_mas_detail_modifications.Show
    emp_mas_detail_modifications.ZOrder
End Sub

Private Sub mnu_emp_modify_Click()
     pwchk = 2
     Load frm_pass
     frm_pass.ZOrder
     frm_pass.Show
End Sub

Private Sub mnu_emp_overtime_entry_Click()
    frm_overtime_entry.Show
    frm_overtime_entry.ZOrder
End Sub

Private Sub mnu_emp_salary_slot_Click()
    emp_mas_slot_entry.Show
    emp_mas_slot_entry.ZOrder
End Sub

Private Sub mnu_emp_worked_posion_Click()
     emp_worked_position.Show
     emp_worked_position.ZOrder
End Sub

Private Sub mnu_employee_additional_Click()
    emptype_chk = 0
    frm_worker_additional_amount.Show
    frm_worker_additional_amount.ZOrder
End Sub

Private Sub mnu_esi_reports_Click()
    esi_reports_frm.Show
    esi_reports_frm.ZOrder
End Sub

Private Sub mnu_excel_fileupload_Click()
    frm_upload.Show
    frm_upload.ZOrder
End Sub

Private Sub mnu_gratuity_settlement_Click()
    frm_gratuity_settlement.Show
End Sub

Private Sub mnu_gratuity_Click()
''    frm_gratuity_settlement.Show
    Frm_rep_gratuity.Show
End Sub

Private Sub MNU_GROSS_PAY_STATEMENT_Click()
    frm_rep_grosspay_annual.Show
    frm_rep_grosspay_annual.ZOrder
End Sub

Private Sub mnu_mangement_month_Deductions_Click()
    emptype_chk = 2
     month_deduct.Show
End Sub

Private Sub mnu_master_details_Click()
    frm_master_details.Show
    frm_master_details.ZOrder
End Sub

Private Sub mnu_mechanical_Click()
    emp_worked_position_mechanical.Show
    emp_worked_position_mechanical.ZOrder
End Sub

Private Sub mnu_ot_import_Click()
    frm_ot_import.Show
    frm_ot_import.ZOrder
End Sub

Private Sub mnu_password_change_Click()
    frm_password_change.Show
    frm_password_change.ZOrder
End Sub

Private Sub mnu_payslip_all_Click()
     pwchk = 1
     optchk = 4
     pay_slip_print.Show
End Sub

Private Sub mnu_payslip_worker_Click()
 pay_slip_print_worker.Show
End Sub

Private Sub mnu_payslp_ho_Click()
     pwchk = 2
     optchk = 2
     pay_slip_print.Show
End Sub

Private Sub mnu_payslp_mills_Click()
     pwchk = 1
     optchk = 1
     pay_slip_print.Show

End Sub

Private Sub mnu_portait_all_Click()
     pwchk = 4
     optchk = 4
     pay_slip_print_portrait.Show

End Sub

Private Sub mnu_portrait_ho_Click()
     pwchk = 2
     optchk = 2
     pay_slip_print_portrait.Show

End Sub

Private Sub mnu_portrait_Man_Click()
     pwchk = 3
     optchk = 3
     pay_slip_print_portrait.Show

End Sub

Private Sub mnu_portrait_mills_Click()
     pwchk = 1
     optchk = 1
     pay_slip_print_portrait.Show

End Sub

Private Sub mnu_present_days_Click()
present_days.Show
End Sub

Private Sub mnu_production_Click()
    emp_worked_position_mechanical.Show
    emp_worked_position_mechanical.ZOrder
End Sub

Private Sub mnu_Rep_overtime_Click()
    frm_rep_overtime.Show
    frm_rep_overtime.ZOrder
End Sub

Private Sub mnu_retainer_deductions_Click()
     emptype_chk = 3
     month_deduct.Show
End Sub

Private Sub mnu_retainer_entry_Click()
    frm_retainer_vou_entry.Show
    frm_retainer_vou_entry.ZOrder
End Sub

Private Sub mnu_retirement_details_statement_Click()
    frm_rep_retirement_details.Show
    frm_rep_retirement_details.ZOrder
End Sub

Private Sub mnu_sal_st_ho_Click()
     pwchk = 2
     optchk = 2
     Salary_statement_prt.Show
End Sub

Private Sub mnu_sal_st_mills_Click()
    pwchk = 1
    optchk = 1
    Salary_statement_prt.Show
End Sub

Private Sub mnu_salary_process_lock_Click()
    control_lock = 3
    frm_control_lock.Show
    frm_control_lock.ZOrder
End Sub
Private Sub mnu_emp_master_upd_Click()
   frm_txt_to_sql.Show
End Sub
Private Sub mnu_indi_deductions_staff_Click()
     emptype_chk = 0
     frm_deduction_individual.Show
     frm_deduction_individual.ZOrder
End Sub

Private Sub mnu_indi_deductions_worker_Click()
     emptype_chk = 1
     frm_deduction_individual.Show
     frm_deduction_individual.ZOrder
End Sub


Private Sub mnu_salary_slots_Click()
     salary_statement_slotwise.ZOrder
     salary_statement_slotwise.Show
End Sub

Private Sub mnu_salary_st_worker_Click()
    salary_statement_worker.Show
End Sub

Private Sub mnu_salary_statement_deptwise_Click()
    salary_statement_departmentwise.Show
End Sub

Private Sub mnu_salary_wages_adv_ent_Click()
    frm_salary_advance.Show
End Sub

Private Sub mnu_salaryst_all_Click()
    pwchk = 4
    optchk = 4
    Salary_statement_prt.Show
End Sub

Private Sub mnu_slip_mgr_above_Click()
     pwchk = 3
     Load frm_pass
     frm_pass.ZOrder
     frm_pass.Show
End Sub

Private Sub mnu_sst_mgr_above_Click()
     optchk = 3
     pwchk = 4
     Load frm_pass
     frm_pass.ZOrder
     frm_pass.Show
End Sub

Private Sub mnu_vou_payment_Click()
     pwchk = 5
     Load frm_pass
     frm_pass.ZOrder
     frm_pass.Show
End Sub

Private Sub mnu_worker_Addn_amt_Click()
    emptype_chk = 1
    frm_worker_additional_amount.Show
    frm_worker_additional_amount.ZOrder
End Sub

Private Sub month_deduct_worker_Click()
     emptype_chk = 1
     month_deduct.Show
End Sub

Private Sub pay_calc_management_Click()
     pay_calchk = 2
     emptype_chk = 2
     pay_cal.Show
End Sub

Private Sub pay_calc_retainer_Click()
     pay_calchk = 3
     emptype_chk = 3
     pay_cal.Show
End Sub

Private Sub pay_calculation_Click()
     pay_calchk = 0
     emptype_chk = 0
     pay_cal.Show
End Sub

Private Sub pay_slip_Click()
     pwchk = 1
     optchk = 1
     pay_slip_print.Show

End Sub

Private Sub payslip_portrait_Click()
     pwchk = 1
     optchk = 1
     pay_slip_print_portrait.Show
End Sub

Private Sub pf_reports_Click()
    pf_reports_frm.Show
End Sub


Private Sub prod_incentive_Click()
    prodn_incentive_frm.Show
End Sub

Private Sub S_ATTENT_Click()
    emptype_chk = 0
    millattn_entry.Show
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
Private Sub COMMAS_Click()
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
     pwchk = 1
''     Load frm_pass
''     frm_pass.ZOrder
''     frm_pass.Show
''''     Load emp_mas_entry
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
    If MAINMENU.ActiveForm Is Nothing Then
''        Unload Me
        End
    Else
        MsgBox "Close All The Opened Forms", vbInformation, "Message"
    End If
    Exit Sub
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
     emptype_chk = 0
     pay_cal.Show
End Sub

Private Sub pay_calc_worker_Click()
     pay_calchk = 1
     emptype_chk = 1
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
    pwchk = 1
    optchk = 1
    Salary_statement_prt.Show
End Sub

Private Sub salary_statement_oldmonths_Click()
    salary_statement_worker_oldmonths.Show
End Sub

Private Sub salary_summary_reports_Click()
     salary_summary_st.Show
End Sub

Private Sub tvls_attn_Click()
     tvlattn_entry.Show
End Sub

Private Sub VDA_AMT_ENTRY_Click()
     VDA_ENTRY.Show
End Sub

Private Sub W_ATTEN_Click()
  emptype_chk = 1
  millattn_entry.Show
End Sub


Public Sub emp_updation()
On Error GoTo err_handler
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay

    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"


''    paydb.BeginTrans

    pst_qry = "delete from bio_empmas"
    paydb.Execute pst_qry


    

'''-
    mdb_qry = "Select * from employees as a, departments as b , companies as c where a.companyid = c.companyid and a.departmentid = b.departmentid and employeecode <> '0' "
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
    
         pst_qry = "insert into  bio_empmas (emp_company,emp_id,emp_pfcode,emp_name,emp_dept,emp_team,emp_status) values (  '" & mdbrs!CompanyFName & "',  " & mdbrs!employeeid & ", " & mdbrs!employeecode & ", '" & mdbrs!employeename & "', '" & mdbrs!departmentfname & "', '" & mdbrs!team & "', '" & mdbrs!Status & "'  )"
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close
    
     
 

    MsgBox ("Updated...")
    Exit Sub
err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
     
End Sub

Private Sub worked_layoff_monthwise_Click()
worked_days_statement.Show
End Sub
