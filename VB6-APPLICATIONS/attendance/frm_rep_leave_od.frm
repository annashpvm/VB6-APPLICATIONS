VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_leave_od 
   Caption         =   "LEAVE , Absent & ON DUTY Details Report"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   16695
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3360
      TabIndex        =   19
      Top             =   6840
      Width           =   1935
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   0
         Picture         =   "frm_rep_leave_od.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   960
         Picture         =   "frm_rep_leave_od.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   480
         TabIndex        =   22
         Top             =   2520
         Width           =   7335
         Begin VB.Frame Frame6 
            Caption         =   "Employee  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   25
            Top             =   120
            Width           =   4815
            Begin VB.OptionButton opt_selective_emp 
               Caption         =   "Selective"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2640
               TabIndex        =   27
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton opt_allemp 
               Caption         =   "All"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   26
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1575
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   6855
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   1410
               Left            =   240
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   24
               Top             =   120
               Width           =   6255
            End
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Select Mill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.OptionButton opt_dpm1 
            Caption         =   "DPM-1"
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
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_dpm3 
            Caption         =   "DPM-2"
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
            Height          =   195
            Left            =   2880
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_all 
            Caption         =   "All"
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
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opt_vjpm 
            Caption         =   "VJPM"
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
            Height          =   195
            Left            =   4200
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opt_cogen 
            Caption         =   "COGEN"
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
            Height          =   195
            Left            =   5520
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame frame_group 
         Caption         =   "DETAILS FOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   7245
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
            Height          =   345
            Left            =   2640
            TabIndex        =   12
            Top             =   240
            Width           =   1545
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
            Height          =   345
            Left            =   5040
            TabIndex        =   11
            Top             =   240
            Width           =   1545
         End
         Begin VB.OptionButton opt_allcat 
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   7335
         Begin VB.ComboBox cmb_rep 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label3 
            Caption         =   "Select Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1665
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   5280
         Width           =   7335
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   2400
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   5520
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   39359
         End
         Begin VB.Label Label1 
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
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Left            =   4200
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_leave_od"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mname As String
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mill = ""
    mcode = ""
    mname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD'"
    cmb_rep.AddItem "Leave Report - Datewise"
    cmb_rep.AddItem "Leave Report - Employeewise"
    cmb_rep.AddItem "CH Report - Datewise"
    cmb_rep.AddItem "OD Report - Datewise"
    cmb_rep.AddItem "OD Report - Employeewise"
    cmb_rep.AddItem "Permission Report - Datewise"
    cmb_rep.AddItem "Permission Report - Employeewise"
    opt_allcat.Value = True
     cmb_rep.Text = "Leave Report - Datewise"
    st_date.Value = Now
    end_date.Value = Now
''    cmb_rep.Text = "Daily Status Report"
    get_emplist
End Sub

Private Sub opt_allcat_Click()
    get_emplist
End Sub

Private Sub opt_cogen_Click()
    mname = "COGEN"

End Sub

Private Sub opt_dpm1_Click()
    mname = "SHVPM"
End Sub

Private Sub opt_dpm3_Click()
    mname = "SHVPM"
End Sub

Private Sub opt_staff_Click()
    get_emplist
End Sub

Private Sub opt_vjpm_Click()
    mname = "SHVPM"

End Sub

Private Sub opt_worker_Click()
    get_emplist
End Sub

Private Sub PROCESS_Click()
   Dim date1 As Date
   date1 = 1 & "/" & 1 & " /" & 1900
   
   mname = "SRI HARI VENKATESWARA PAPER MILL PVT LTD"
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   
   Cry_rep1.PrinterSelect
   Dim sw, ds, emp, mill As String
   emp = ""
   If opt_selective_emp.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_emp.ListCount > 0 Then
           For pin_row = 0 To lst_emp.ListCount - 1
               If lst_emp.Selected(pin_row) = True Then
                  If i = 0 Then
                     emp = " and ( {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
                     i = i + 1
                  Else
                     emp = emp + " or {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If emp <> "" Then emp = emp + ")"
   
   
   If opt_allcat.Value = True Then
      sw = ""
   ElseIf opt_staff.Value = True Then
      sw = "and {bio_empmas.bioemp_team} = 'STAFF'"
   Else
      sw = "and {bio_empmas.bioemp_team} = 'WORKER'"
   End If

  
  
  
   ds = sw + emp + mill
   If cmb_rep.Text = "Leave Report - Employeewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_leavedetails_empwise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_empleave.emp_leave_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_empleave.emp_leave_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf cmb_rep.Text = "Leave Report - Datewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_leavedetails_datewise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_empleave.emp_leave_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_empleave.emp_leave_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf cmb_rep.Text = "OD Report - Employeewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_oddetails_empwise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_emp_oddetails.empod_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_oddetails.empod_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf cmb_rep.Text = "OD Report - Datewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_oddetails_datewise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_emp_oddetails.empod_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_oddetails.empod_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf cmb_rep.Text = "Permission Report - Datewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_permission_datewise.rpt"
      pst_qry = "{bio_emp_permissions.empp_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_permissions.empp_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")"
      Cry_rep1.ReplaceSelectionFormula ("{bio_emp_permissions.empp_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_permissions.empp_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf cmb_rep.Text = "Permission Report - Employeewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_permission_empwise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_emp_permissions.empp_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_permissions.empp_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
      
   ElseIf cmb_rep.Text = "CH Report - Datewise" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_bio_emp_chdetails_datewise.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_emp_chleave.empch_ch_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_emp_chleave.empch_ch_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   End If
    
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1

End Sub

Private Sub opt_selective_emp_Click()
    lst_emp.Enabled = True
''    lst_dept.Visible = False
    lst_emp.Visible = True
End Sub
Private Sub opt_allemp_Click()
    lst_emp.Enabled = False
End Sub
Public Sub get_emplist()
''    If opt_staff.Value = True Then
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    Else
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    End If
''If opt_allcat.Value = True Then
''    If opt_all.Value = True Then
''       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_workplace  = 'MILL' order by emp_name "
''    Else
''       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A' and emp_workplace  = 'MILL' order by emp_name  "
''    End If
''Else
''    If opt_all.Value = True Then
''       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_cat = '" & emp_type & "' and emp_workplace  = 'MILL' order by emp_name "
''    Else
''       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A'  and emp_cat = '" & emp_type & "'  and emp_workplace  = 'MILL' order by emp_name  "
''    End If
''End If
    
   If opt_allcat.Value = True Then
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_status = 'Working' "
        Else
           sql = "Select * from  bio_empmas where bioemp_company = '" & mill & "' and bioemp_status = 'Working'"
        End If
   ElseIf opt_staff.Value = True Then
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_team = 'STAFF' and bioemp_status = 'Working'"
        Else
           sql = "Select * from  bio_empmas where bioemp_team =  'STAFF' and bioemp_company = '" & mill & "' and bioemp_status = 'Working' "
        End If
   Else
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_team <> 'STAFF' and bioemp_status = 'Working'"
        Else
           sql = "Select * from  bio_empmas where bioemp_team <>  'STAFF' and bioemp_company = '" & mill & "' and bioemp_status = 'Working'"
        End If
   
   End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_emp.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_emp.AddItem payrs("bioemp_name")
        payrs.MoveNext
    Wend
End Sub

