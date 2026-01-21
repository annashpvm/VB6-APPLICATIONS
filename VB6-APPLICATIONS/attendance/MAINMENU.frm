VERSION 5.00
Begin VB.MDIForm MAINMENU 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4680
   Icon            =   "MAINMENU.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu master_entry 
      Caption         =   "MASTERS"
      Begin VB.Menu mnu_biomaster_upd 
         Caption         =   "EMPLOYEE LIST UPLOAD FROM BIO-METRIC"
      End
      Begin VB.Menu bio2 
         Caption         =   "-"
      End
      Begin VB.Menu emp_master_import 
         Caption         =   "Employee master - Import"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_attn_update 
         Caption         =   "Attendance staus -update"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_employee_requirement 
         Caption         =   "Employee Requirement"
      End
   End
   Begin VB.Menu DATA_ENTRY 
      Caption         =   "   DATA ENTRY      "
      Begin VB.Menu mnu_bio_upload 
         Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRICS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bio_upload_new 
         Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRICS - FOR MONTH"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_bio_upload_new_days 
         Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRICS "
      End
      Begin VB.Menu mnu_attend_process_modi 
         Caption         =   "Attendance Process && View FOR Individual"
      End
      Begin VB.Menu bio 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_shift_schedule 
         Caption         =   "Shift Schedule"
         Begin VB.Menu mnu_week_shift_shedule 
            Caption         =   "Weekwise Shift Schedule"
         End
         Begin VB.Menu mnu_shift_schedule_randon 
            Caption         =   "Random Shift Schedule"
         End
      End
      Begin VB.Menu mnu_shift_shedule_modification 
         Caption         =   "Shift Schedule Modification"
      End
      Begin VB.Menu mnu_c_shift 
         Caption         =   "C Shift / 12 Hrs Shift Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu frm_in_out_manual 
         Caption         =   "Manual IN/ OUT Punch"
      End
      Begin VB.Menu mnu_in_out_2 
         Caption         =   "Manual IN/ OUT Punch - for Others"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_leave_entries 
         Caption         =   "Leave Entries"
      End
      Begin VB.Menu mnu_ch_entry 
         Caption         =   "CH Entry"
      End
      Begin VB.Menu mnu_onduty_entry 
         Caption         =   "On Duty Entry"
      End
      Begin VB.Menu mnu_permission_entry 
         Caption         =   "Permission Entry"
      End
      Begin VB.Menu mnu_layoff_entry 
         Caption         =   "Layoff Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_satuday_assign 
         Caption         =   "Saturday - Absent Assignment"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_prodn_incentive 
         Caption         =   "Over Time Entry"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_mess_deduction 
         Caption         =   "Canteen Recovery Entry"
      End
      Begin VB.Menu mnu_canteen 
         Caption         =   "Canteen Expenses Entry"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_attn_corrections 
         Caption         =   "Attendance Corrections"
      End
      Begin VB.Menu mnu_dh_entry 
         Caption         =   "Declare Holiday Entry "
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_dh_eligibility 
         Caption         =   "Decalare Holiday Eligibility"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu reports 
      Caption         =   "   REPORTS      "
      Begin VB.Menu mnu_bio_attend 
         Caption         =   "Bio-metric Attendance Reports"
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
      Begin VB.Menu mnu_leave_od_reports 
         Caption         =   "Leave , Permission  && OD Reports"
      End
      Begin VB.Menu mnu_prodn_incentive_rep 
         Caption         =   "Over Time Reports"
      End
      Begin VB.Menu mnu_canteen_rep 
         Caption         =   "Canteen Report"
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
Private Sub calculator_Click()
    Dim pvt_calc As Variant
    pvt_calc = Shell("calc.exe", vbMaximizedFocus)
    Exit Sub
End Sub


Private Sub emp_master_import_Click()
    emp_import
End Sub

Private Sub frm_in_out_manual_Click()
    frm_manual_inout_punch.Show
End Sub

Private Sub MDIForm_Load()
   optchk = 0
''    paydb.Open pay
''    If adminpw = 0 Then
''       mnu_in_out_2.Enabled = False
''       frm_in_out_manual.Enabled = False
''    End If
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

Private Sub mnu_attn_lock_Click()
    control_lock = 1
    frm_control_lock.Show
    frm_control_lock.ZOrder
End Sub

Private Sub mnu_attn_retainer_Click()
    emptype_chk = 2
    millattn_entry.Show
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

Private Sub mnu_attend_process_modi_Click()
   frm_bio_process.Show
   frm_bio_process.ZOrder
End Sub

Private Sub mnu_attn_corrections_Click()
   frm_Attn_Corrections.Show
   frm_Attn_Corrections.ZOrder
End Sub

Private Sub mnu_attn_update_Click()
   attn_update
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

Private Sub mnu_bio_upload_new_days_Click()
    frm_bio_metric_upload_days.Show
    frm_bio_metric_upload_days.ZOrder
End Sub

Private Sub mnu_bio_vew_Click()
   frm_rep_bio_metrics_view.Show
   frm_rep_bio_metrics_view.ZOrder
End Sub

Private Sub mnu_biomaster_upd_Click()
 emp_updation
End Sub

Private Sub mnu_c_shift_Click()
    frm_shift_c_entry.Show
End Sub

Private Sub mnu_canteen_Click()
    frm_canteen_expenses.Show
    frm_canteen_expenses.ZOrder
End Sub

Private Sub mnu_canteen_rep_Click()
    frm_rep_canteen_details.Show
    frm_rep_canteen_details.ZOrder
End Sub

Private Sub mnu_ch_entry_Click()
    frm_leave_chentry.Show
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

Private Sub mnu_emp_overtime_entry_Click()
    frm_overtime_entry.Show
    frm_overtime_entry.ZOrder
End Sub

Private Sub mnu_emp_salary_slot_Click()
    emp_mas_slot_entry.Show
    emp_mas_slot_entry.ZOrder
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

Private Sub mnu_password_change_Click()
    frm_password_change.Show
    frm_password_change.ZOrder
End Sub

Private Sub mnu_payslip_all_Click()
     pwchk = 1
     optchk = 4
     pay_slip_print.Show
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

Private Sub exit_Click()
    If MAINMENU.ActiveForm Is Nothing Then
''        Unload Me
        End
    Else
        MsgBox "Close All The Opened Forms", vbInformation, "Message"
    End If
    Exit Sub
End Sub
Public Sub emp_updation()
On Error GoTo err_handler
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.252\Software\att2000.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.75\attendance\att2000.mdb"
     dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.75\d\attendance\att2000.mdb"
     
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\attendance\att2000.mdb"

''    \\10.0.0.252\Software\haribiom.mdb
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\eTimeTrackLite1.mdb"
''    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\eTimeTrackLite1.mdb"


''    paydb.BeginTrans
 
 
'' dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.9\eTimeTrackLite1.mdb"

    pst_qry = "delete from bio_empmas"
    paydb.Execute pst_qry
  

    mdb_qry = "Select * from employees as a, departments as b , companies as c where a.companyid = c.companyid and a.departmentid = b.departmentid and employeecode <> '0' "
    
    mdb_qry = "Select * from USERINFO"
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF
    
         If Val(mdbrs!userid) > 0 Then
''            pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name,bioemp_dept,bioemp_team,bioemp_status) values (  '" & mdbrs!CompanyFName & "',  " & mdbrs!employeeid & ", " & mdbrs!employeecode & ", '" & mdbrs!employeename & "', '" & mdbrs!departmentfname & "', '" & mdbrs!team & "', '" & mdbrs!Status & "'  )"
            pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  " & mdbrs!userid & ", " & mdbrs!Badgenumber & ", '" & mdbrs!Name & "' )"
            paydb.Execute pst_qry
         End If
         mdbrs.MoveNext
    Wend
    mdbrs.Close
    
''Inserted data for Outstation staffs
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '1767', '1767', 'THIRUNARAYAN' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10018', '10018', 'VIJAY ANAND' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10014', '10014', 'SRI PRAKASH' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '1302', '1302', 'BALASUBRAMANIAN' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10023', '10023', 'MAINKANDAN' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10005', '10005', 'THIRUNAVUKARASU' )"
    paydb.Execute pst_qry
    
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10006', '10006', 'THANGAM V' )"
    paydb.Execute pst_qry
    
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10010', '10010', 'ANANTHA PRABHU' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10030', '10030', 'VARATHARAJU D' )"
    paydb.Execute pst_qry
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10031', '10031', 'RAMESH KB' )"
    paydb.Execute pst_qry
    
    
    pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  '10007', '10007', 'RAMESH KB' )"
    paydb.Execute pst_qry
    
    
    
    pst_qry = "update bio_empmas set  bioemp_status = (case when emp_status = 'A' then 'Working' else 'Resigned' end)  , bioemp_name = emp_name ,bioemp_gender = EMP_SEX,bioemp_dept = dept_name ,bioemp_team = (case when emp_cat = 'S' then 'STAFF' else 'WORKER' end) ,bioemp_workhrs = emp_work_hrs  ,bioemp_grosspay = emp_grosspay from emp_mas a,bio_empmas b,pdept_mas c where emp_dept = dept_code and bioemp_fpcode = emp_code"
    paydb.Execute pst_qry


 

    MsgBox ("Employees are Updated...")
    Exit Sub
err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
     
End Sub

Private Sub mnu_dh_eligibility_Click()
    frm_dec_holiday_eligibility.Show
End Sub

Private Sub mnu_dh_entry_Click()
  frm_dec_holiday_entry.Show
End Sub

Private Sub mnu_employee_requirement_Click()
    mas_employee_requirement.Show
    mas_employee_requirement.ZOrder
End Sub

Private Sub mnu_in_out_2_Click()
    frm_manual_inout_punch_head.Show
End Sub

Private Sub mnu_layoff_entry_Click()
   frm_layoff_entries.Show
   frm_layoff_entries.ZOrder
End Sub

Private Sub mnu_leave_entries_Click()
    frm_leave_entries.Show
    frm_leave_entries.ZOrder
End Sub

Private Sub mnu_leave_od_reports_Click()
   frm_rep_leave_od.Show
   frm_rep_leave_od.ZOrder
End Sub

Private Sub mnu_mc_prodn_Click()
    frm_mcprod_entry.Show
End Sub

Private Sub mnu_mess_deduction_Click()
    frm_canteen_recovery.Show
End Sub

Private Sub mnu_onduty_entry_Click()
   frm_od_entries.Show
End Sub

Private Sub mnu_permission_entry_Click()
   frm_permission_entries.Show
End Sub

Private Sub mnu_prodn_incentive_Click()
   frm_overtime.Show
   frm_overtime.ZOrder
End Sub

Private Sub mnu_prodn_incentive_rep_Click()
    frm_rep_prodn_incentive.Show
    frm_rep_prodn_incentive.ZOrder
End Sub

Private Sub mnu_satuday_assign_Click()
   frm_saturday.Show
   frm_saturday.ZOrder
End Sub

Private Sub mnu_shift_schedule_randon_Click()
    frm_shift_schdule_random_new.Show
    frm_shift_schdule_random_new.ZOrder
End Sub

Private Sub mnu_shift_shedule_modification_Click()
    frm_shift_schdule_modification.Show
End Sub

Private Sub mnu_week_shift_shedule_Click()
    frm_shift_schdule.Show
End Sub



Public Sub emp_import()
On Error GoTo err_handler
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
    dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.252\Software\haribio.mdb"



Dim sw As Integer
''    mdb_qry = "Select * from employees as a, departments as b , companies as c where a.companyid = c.companyid and a.departmentid = b.departmentid and employeecode <> '0' "
    mdb_qry = "Select * from emp where status = 'Working'"
''    mdb_qry = "Select a.*,b.depcode as departmentcode,c.* from emp a, dept b, desig c where  a.depname = b.department and a.design = c.designation and status = 'Working'"

    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    While Not mdbrs.EOF

         If mdbrs!Status = "Working" Then
''            pst_qry = "insert into  emp_mas (emp_company,emp_code,emp_fpcode,emp_name,emp_sex,emp_basic, emp_dob , emp_doj, emp_type, emp_pfno, emp_uan, emp_marital, emp_grosspay, emp_basic, emp_fda,emp_hra , emp_medall, emp_lta, emp_status, emp_pfeligible, emp_bank_no, emp_esino, emp_eligible, emp_pfsalary, emp_itded ) " _
                 & " values (  1 ,  " & mdbrs!empid & ", " & mdbrs!empid & ", '" & mdbrs!empname & "', '" & mdbrs!sex & "', " & mdbrs!basicpay & ",'" & Format(mdbrs!birthdate, "MM/dd/yyyy") & "','" & Format(mdbrs!joindate, "MM/dd/yyyy") & "', '" & Left(mdbrs!staftype, 1) & "','" & mdbrs!pfno & "', '" & mdbrs!pfuan & "', '" & mdbrs!marital & "','" & mdbrs!salary & "', '" & mdbrs!basicpay & "','" & mdbrs!da & "','" & mdbrs!hra & "','" & mdbrs!medialow & "', '" & mdbrs!itravel & "','" & Left(mdbrs!Status, 1) & "', '" & mdbrs!pfflg & "','" & mdbrs!salbankac & "','" & mdbrs!esino & "','" & mdbrs!esiflg & "','" & mdbrs!pfsalary & "','" & mdbrs!itded & "' )"

''            pst_qry = "insert into  emp_mas (emp_company,emp_code,emp_fpcode,emp_name,emp_sex,emp_basic, emp_dob , emp_doj, emp_type, emp_pfno, emp_uan, emp_marital, emp_grosspay,emp_fda , emp_hra, emp_medall, emp_lta, emp_status, emp_pfeligible, emp_bank_acno, emp_esino, emp_esieligible, emp_pfsalary, emp_itded,emp_dept,emp_design ) " _
                 & " values (  1 ,  " & mdbrs!empid & ", " & mdbrs!empid & ", '" & mdbrs!empname & "', '" & mdbrs!sex & "', " & mdbrs!basicpay & ",'" & Format(mdbrs!birthdate, "MM/dd/yyyy") & "','" & Format(mdbrs!joindate, "MM/dd/yyyy") & "', '" & Left(mdbrs!staftype, 1) & "','" & mdbrs!pfno & "', '" & mdbrs!pfuan & "','" & Left(mdbrs!marital, 1) & "'," & mdbrs!salary & ",'" & mdbrs!da & "','" & mdbrs!hra & "','" & mdbrs!medialow & "', '" & mdbrs!ltravel & "','" & Left(mdbrs!Status, 1) & "', '" & mdbrs!pfflg & "','" & mdbrs!salbankac & "','" & mdbrs!esino & "','" & mdbrs!esiflg & "','" & mdbrs!pfsalary & "','" & mdbrs!itded & "','" & mdbrs!departmentcode & "','" & mdbrs!decode & "')"
                 
                 

            pst_qry = "insert into  emp_mas (emp_company,emp_code,emp_fpcode,emp_name,emp_sex,emp_basic, emp_dob , emp_doj, emp_type, emp_pfno, emp_uan, emp_marital, emp_grosspay,emp_fda , emp_hra, emp_medall, emp_lta, emp_status, emp_pfeligible, emp_bank_acno, emp_esino, emp_esieligible, emp_pfsalary, emp_itded,emp_cat,depart,desi) " _
                & " values (  1 ,  " & mdbrs!empid & ", " & mdbrs!empid & ", '" & mdbrs!empname & "', '" & mdbrs!sex & "', " & mdbrs!basicpay & ",'" & Format(mdbrs!birthdate, "MM/dd/yyyy") & "','" & Format(mdbrs!joindate, "MM/dd/yyyy") & "', '" & Left(mdbrs!staftype, 1) & "','" & mdbrs!pfno & "', '" & mdbrs!pfuan & "','" & Left(mdbrs!marital, 1) & "'," & mdbrs!salary & ",'" & mdbrs!da & "','" & mdbrs!hra & "','" & mdbrs!medialow & "', '" & mdbrs!ltravel & "','" & Left(mdbrs!Status, 1) & "', '" & mdbrs!pfflg & "','" & mdbrs!salbankac & "','" & mdbrs!esino & "','" & mdbrs!esiflg & "','" & mdbrs!pfsalary & "','" & mdbrs!itded & "', '" & Left(mdbrs!staftype, 1) & "' , '" & UCase(mdbrs!depname) & "' , '" & UCase(mdbrs!design) & "')"



            paydb.Execute pst_qry
         End If
         mdbrs.MoveNext
    Wend
    mdbrs.Close
''
''    mdb_qry = "Select * from dept"
''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    While Not mdbrs.EOF
''
''            pst_qry = "insert into  pdept_mas  values (   " & mdbrs!depcode & ", '" & mdbrs!department & "' )"
''
''            paydb.Execute pst_qry
''
''         mdbrs.MoveNext
''    Wend
''    mdbrs.Close

''
''    mdb_qry = "Select * from desig"
''    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''    While Not mdbrs.EOF
''
''            pst_qry = "insert into  pdesi_mas  values (   " & mdbrs!decode & ", '" & mdbrs!designation & "' )"
''
''            paydb.Execute pst_qry
''
''         mdbrs.MoveNext
''    Wend
''    mdbrs.Close

    MsgBox ("Updated...")
    Exit Sub
err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
     
End Sub


Public Sub attn_update()
On Error GoTo err_handler
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay
    
    
    sql = "select * from emp_mas"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        sql = "insert into attn_entry (attn_company,attn_finyear,attn_month,attn_year,attn_empcode,attn_empcat) values ('1','21','1','2021', " & payrs!emp_code & ", '" & payrs!emp_cat & "')"
        paydb.Execute sql
        sql2 = "insert into emp_eligible_leave (s_company,s_finyear,s_year,s_empcode,s_el,s_empcat,s_workplace) values ('1','21','2021', " & payrs!emp_code & ",0, '" & payrs!emp_cat & "','MIL')"
        paydb.Execute sql2
        
        
        payrs.MoveNext
    Wend
    payrs.Close
    

    MsgBox ("Updated...")
    Exit Sub
err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
     
End Sub



