VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form pf_reports_frm 
   Caption         =   "PF - REPORTS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   1215
      Left            =   12720
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton opt_pf_export_employer_cont 
         Caption         =   "Online - PF Return- ECR Format - for Employer Contribution"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   6345
      End
      Begin VB.OptionButton opt_form3 
         Caption         =   "FORM - III REPORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   480
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.OptionButton opt_pf_export_emp_cont 
         Caption         =   "Online - PF Return- ECR Format - for Employee Contribution"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Visible         =   0   'False
         Width           =   6345
      End
      Begin VB.OptionButton opt_pf_export 
         Caption         =   "Online - PF Return- ECR Format"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.OptionButton opt_epf_3500 
         Caption         =   "EPF && FPF STATEMENT - FOR WOKED DAYS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   4755
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   -120
      TabIndex        =   16
      Top             =   6960
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121372673
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121372673
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker dt_joining 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121372673
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker dt_resigned 
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121372673
         CurrentDate     =   39359
      End
      Begin VB.Label Label10 
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
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
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
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   255
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "PF  - REPORTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7815
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   9390
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   600
         TabIndex        =   8
         Top             =   5640
         Width           =   8175
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
            Left            =   6150
            TabIndex        =   10
            Top             =   240
            Width           =   1335
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
            Left            =   1800
            TabIndex        =   9
            Top             =   315
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   360
            TabIndex        =   12
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4980
            TabIndex        =   11
            Top             =   315
            Width           =   885
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   3000
         TabIndex        =   4
         Top             =   6840
         Width           =   3015
         Begin VB.CommandButton Exit 
            Caption         =   "&Exit"
            Height          =   870
            Left            =   1920
            Picture         =   "pf_reports_frm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Refresh 
            Caption         =   "&Refresh"
            Height          =   870
            Left            =   960
            Picture         =   "pf_reports_frm.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton print 
            Caption         =   "&Print"
            Height          =   870
            Left            =   0
            Picture         =   "pf_reports_frm.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "SELECT STAFF / WORKER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8160
         TabIndex        =   3
         Top             =   5400
         Visible         =   0   'False
         Width           =   930
         Begin VB.OptionButton opt_all 
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
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   720
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
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
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   3120
            TabIndex        =   14
            Top             =   240
            Width           =   1695
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
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   5520
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4905
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   8160
         Begin VB.OptionButton opt_epf_st_dept 
            Caption         =   "EPF && FPF STATEMENT - Department wise"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   5115
         End
         Begin VB.Frame Frame7 
            Height          =   375
            Left            =   6600
            TabIndex        =   24
            Top             =   3720
            Visible         =   0   'False
            Width           =   855
            Begin VB.OptionButton opt_new_employer 
               Caption         =   " Employer Contribution"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   360
               TabIndex        =   27
               Top             =   960
               Width           =   2625
            End
            Begin VB.OptionButton opt_new_all 
               Caption         =   "ALL Contribution"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   720
               TabIndex        =   26
               Top             =   120
               Value           =   -1  'True
               Width           =   2625
            End
            Begin VB.OptionButton opt_new_employee 
               Caption         =   " Employee Contribution"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   360
               TabIndex        =   25
               Top             =   600
               Width           =   2625
            End
         End
         Begin VB.OptionButton opt_pf_export_new 
            Caption         =   "Online - PF Return- ECR Format - NEW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   240
            TabIndex        =   23
            Top             =   2400
            Visible         =   0   'False
            Width           =   3945
         End
         Begin VB.OptionButton opt_epf_st 
            Caption         =   "EPF && FPF STATEMENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   3315
         End
      End
   End
End
Attribute VB_Name = "pf_reports_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pst_qry As String
Private Sub cmb_month_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     find_dates
  End If
End Sub
Private Sub cmb_year_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     find_dates
  End If
End Sub
Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
    opt_epf_st.Value = True
    opt_All.Value = True
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
''    End With
''''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
''    cmb_year.Text = "2015"
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    sql = ("Select * from  emp_salary")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic

End Sub

Private Sub opt_pf_export_new_Click()
    Frame7.Visible = True
End Sub

Private Sub print_Click()
   
   If opt_pf_export.Value = True Then
      pfreturn
      Exit Sub
   End If
   
   If opt_pf_export_new.Value = True Then
      pfreturn_new
      Exit Sub
   End If
   If opt_pf_export_emp_cont.Value = True Then
      pfreturn_emp_contribution
      Exit Sub
   End If
   
   If opt_pf_export_employer_cont.Value = True Then
      pfreturn_employer_contribution
      Exit Sub
   End If
   
   
   disname = "MONTHLY EPF & FPF STATEMENT FOR THE MONTH OF "
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(1) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(2) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   If opt_pf_export.Value = True Then
    cry_rep1.Formulas(3) = ("rdate = '" & Format(st_date.Value, "yyyy/MM/dd") & "'")
   End If
   ''    rpt.ParameterFields(0) = "@parameter_name_1;" & "parameter_value" & ";TRUE"
   ''cry_rep1.ParameterFields(0) = "@repdate;" & Format(st_date.Value, "yyyy-mm-dd") & ";TRUE"
   cry_rep1.ParameterFields(0) = ""
   
   If opt_epf_st.Value = True Then
        If finyear >= 13 Then
          cry_rep1.ParameterFields(0) = "repdate;date( " & year(st_date.Value) & ", " & Month(st_date.Value) & "," & Day(st_date.Value) & ");TRUE"
   
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\epf_fpf_statement.rpt"
        Else
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\epf_fpf_statement_calc_old.rpt"
        End If
   ElseIf opt_epf_st_dept.Value = True Then
          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\epf_fpf_statement_deptwise.rpt"
   ElseIf opt_epf_3500.Value = True Then
   
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\epf_fpf_statement3500.rpt"
''        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker.rpt"

   End If

''   If opt_epf_st.Value = True Then
      If opt_staff.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and ( {emp_mas.emp_cat} = 'S' or {emp_mas.emp_cat} = 'M') and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
      ElseIf opt_worker.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_mas.emp_cat} = 'W' and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
      Else
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
      End If
''   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub

Private Sub refresh_Click()
    opt_epf_st.Value = True
    opt_staff.Value = True
End Sub

Public Sub pfreturn()
    Dim ecode As String
    Dim epfno As String
    Dim empname As String
    Dim epfwages As String
    Dim epswages As String
    Dim PF As String
    Dim eps_contribution As String
    Dim eps_wages, pfamt, epsamt, epsbalamt, lop As Double
    Dim ncp As String
    Dim pfrefund As String
    Dim e_f_name As String
    Dim e_relation As String
    Dim e_dob As String
    Dim e_doJ As String
    Dim e_sex As String
    Dim e_dol As String
    Dim e_reason As String
    Dim e_ar_epf_wages As String
    Dim e_ar_epf_ee As String
    Dim e_ar_epf_er As String
    Dim e_ar_eps As String
    Dim empr_eligible As Double
    Dim eligible, eligible2 As Double
    epfno = Space(7)
    empname = Space(85)
    epfwages = Space(10)
    epswages = Space(10)
    eps_contribution = 10
    epsamt2 = Space(10)
    epsbalamt2 = Space(10)
    PF = Space(10)
    ncp = Space(5)
    pfrefund = Space(10)
    
    e_f_name = Space(60)
    e_relation = Space(1)
    e_dob = Space(12)
    e_doJ = Space(12)
    e_sex = Space(1)
    e_dol = Space(12)
    e_reason = Space(1)
    
    e_ar_epf_wages = Space(10)
    e_ar_epf_ee = Space(10)
    e_ar_epf_er = Space(10)
    e_ar_eps = Space(10)
    Dim i As Integer
    i = 0
    Dim rs_set As New ADODB.Recordset
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 and emp_pfno = '' order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        MsgBox ("PF Numebr Not available for the employee ..." + rs_set!emp_name)
        i = i + 1
        rs_set.MoveNext
    Wend
    If i > 0 Then Exit Sub
    rs_set.Close
    Dim DoBVar, DayVar, MthVar, YrsVar As Integer
''    Dim filename As String
    CommonDialog1.Filter = "Text Files (*.txt)"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
''    filename = CommonDialog1.filename + ".txt"
''    Open "c:\rep.txt" For Output As #1
    Dim employer_pfamt As Double
    Open CommonDialog1.FileName + ".txt" For Output As #1
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        ecode = Trim(rs_set!emp_code)
''        If ecode = "368" Then
''         MsgBox "Jebes"
''        End If
        eligible = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        eligible2 = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        eps_wages = eligible
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff) / 2, 1)
        If rs_set!s_month < 9 And rs_set!s_month <= 2014 Then
            empr_eligible = IIf(eligible >= 6500, 6500, eligible)
            eps_wages = IIf(eligible >= 6500, 6500, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible >= 6500, 6500, eligible)
               eps_wages = IIf(eligible >= 6500, 6500, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If eligible >= 6500 Then
                    pfamt = 780
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                End If
            End If
        Else
            empr_eligible = IIf(eligible > 15000, 15000, eligible)
             ''empr_eligible = IIf(eligible > 15000, 1500, eligible)
            eps_wages = IIf(eligible > 15000, 15000, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible > 15000, 15000, eligible)
               eps_wages = IIf(eligible > 15000, 15000, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If eligible > 15000 Then
                    pfamt = 1800
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                End If
            End If
        
        End If
        
        employer_pfamt = Round(empr_eligible * 0.12, 0)
        epsamt = Round(eps_wages * 0.0833, 0)
        epsbalamt = Round(employer_pfamt - epsamt, 0)
        dt_joining.Value = Format(rs_set!emp_doj, "MM/dd/yyyy")
        
        If IsNull(rs_set!emp_resigneddate) Then
           dt_resigned.Value = vbNull
        Else
           dt_resigned.Value = Format(rs_set!emp_resigneddate, "MM/dd/yyyy")
        End If
        
        DoBVar = IIf((100 * Month(end_date) + Day(end_date)) < (100 * Month(rs_set!emp_dob) + Day(rs_set!emp_dob)), 1, 0)
        MthVar = DateDiff("m", rs_set!emp_dob, end_date - DoBVar) Mod 12
        DayVar = DateDiff("d", rs_set!emp_dob, end_date)
        YrsVar = DateDiff("yyyy", rs_set!emp_dob, end_date) - DoBVar
        If YrsVar >= 58 And MthVar + DayVar >= 1 Then
            epsamt = 0
            epsbalamt = employer_pfamt
        End If
        LSet epfno = Trim(rs_set!emp_pfno)
        LSet empname = Trim(rs_set!emp_name)
        
        
        
''        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
''           LSet epfwages = Trim(Format$(rs_set!s_eligible_grosspay, "0"))
''        Else
''           LSet epfwages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
''        End If
        
''        LSet epswages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
           LSet epfwages = Trim(Format$(eligible2, "0"))
        Else
           If rs_set!s_month < 9 And rs_set!s_month <= 2014 Then
              LSet epfwages = Trim(Format$(IIf(eligible2 >= 6500, 6500, eligible2), "0"))
           Else
              LSet epfwages = Trim(Format$(IIf(eligible2 > 15000, 15000, eligible2), "0"))
           End If
        End If

''        LSet PF = Trim(Format$(rs_set!s_pf, "0"))
        LSet PF = Trim(Format$(pfamt, "0"))
        
        LSet epsamt2 = Trim(Format$(epsamt, "0"))
        LSet epsbalamt2 = Trim(Format$(epsbalamt, "0"))
        LSet ncp = Trim(Format$(lop, "0"))
        LSet pfrefund = Trim(Format$(rs_set!s_pfded, "0"))
     
''        If Format(rs_set!emp_doj, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_doj, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
''        If Format(dt_joining.Value, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(dt_joining.Value, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        If dt_joining.Value >= st_date.Value And dt_joining.Value <= end_date.Value Then
        
           e_f_name = rs_set!emp_fname
           e_relation = IIf(rs_set!emp_relation = "F", "F", "S")
           e_dob = Format(rs_set!emp_dob, "dd/MM/yyyy")
           e_doJ = Format(rs_set!emp_doj, "dd/MM/yyyy")
           e_sex = rs_set!emp_sex
        Else
           e_f_name = ""
           e_relation = ""
           e_dob = ""
           e_doJ = ""
           e_sex = ""
        End If
        
        e_dol = ""
        e_reason = ""
        
        If Left(rs_set!emp_status, 1) = "R" Then
''           If Format(rs_set!emp_resigneddate, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_resigneddate, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
            If dt_resigned.Value >= st_date.Value And dt_resigned.Value <= end_date.Value Then
                If rs_set.Fields("emp_reason") = "CESSATION" Then
                   e_reason = "C"
                ElseIf rs_set.Fields("emp_reason") = "SUPERANNUATION" Then
                   e_reason = "S"
                ElseIf rs_set.Fields("emp_reason") = "RETIREMENT" Then
                   e_reason = "R"
                ElseIf rs_set.Fields("emp_reason") = "DEATH IN SERVICE" Then
                   e_reason = "D"
                ElseIf rs_set.Fields("emp_reason") = "PERMANENT DISABLEMENT" Then
                   e_reason = "P"
                End If
                If IsNull(rs_set!emp_resigneddate) = True Then
                   e_dol = ""
                Else
                   e_dol = Format(rs_set!emp_resigneddate, "dd/MM/yyyy")
                End If
           End If
           End If
        LSet e_ar_epf_wages = "0"
        LSet e_ar_epf_ee = "0"
        LSet e_ar_epf_er = "0"
        LSet e_ar_eps = "0"
               
        sql = "Select * from arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & cmb_year & " and e_company = " & company_code & " and e_empcode  = '" & ecode & "'"
        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
        If Not payrs2.EOF Then
           arrearamt = payrs2("e_amount")
        End If
        payrs2.Close
        Print #1, Trim(epfno) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epswages) + "#~#" + Trim(PF) + "#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund) + "#~#" + Trim(e_ar_epf_wages) + "#~#" + Trim(e_ar_epf_ee) + "#~#" + Trim(e_ar_epf_er) + "#~#" + Trim(e_ar_eps) + "#~#" + Trim(e_f_name) + "#~#" + Trim(e_relation) + "#~#" + Trim(e_dob) + "#~#" + Trim(e_sex) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_reason)
        rs_set.MoveNext
    Wend
    Close #1
    MsgBox ("File Saved...")
    Dim path As String
    path = CommonDialog1.FileName + ".TXT"
    If prt_stat = False Then
        retval = Shell("Edit.com " & path, vbMaximizedFocus)
    End If
    Me.MousePointer = 1
End Sub

Public Sub pfreturn_emp_contribution()
    Dim ecode As String
    Dim epfno As String
    Dim empname As String
    Dim epfwages As String
    Dim epswages As String
    Dim PF As String
    Dim eps_contribution As String
    Dim eps_wages, pfamt, epsamt, epsbalamt, lop As Double
    Dim ncp As String
    Dim pfrefund As String
    Dim e_f_name As String
    Dim e_relation As String
    Dim e_dob As String
    Dim e_doJ As String
    Dim e_sex As String
    Dim e_dol As String
    Dim e_reason As String
    Dim e_ar_epf_wages As String
    Dim e_ar_epf_ee As String
    Dim e_ar_epf_er As String
    Dim e_ar_eps As String
    Dim empr_eligible As Double
    Dim eligible, eligible2 As Double
    epfno = Space(7)
    empname = Space(85)
    epfwages = Space(10)
    epswages = Space(10)
    eps_contribution = 10
    epsamt2 = Space(10)
    epsbalamt2 = Space(10)
    PF = Space(10)
    ncp = Space(5)
    pfrefund = Space(10)
    
    e_f_name = Space(60)
    e_relation = Space(1)
    e_dob = Space(12)
    e_doJ = Space(12)
    e_sex = Space(1)
    e_dol = Space(12)
    e_reason = Space(1)
    
    e_ar_epf_wages = Space(10)
    e_ar_epf_ee = Space(10)
    e_ar_epf_er = Space(10)
    e_ar_eps = Space(10)
    Dim i As Integer
    i = 0
    Dim rs_set As New ADODB.Recordset
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 and emp_pfno = '' order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        MsgBox ("PF Numebr Not available for the employee ..." + rs_set!emp_name)
        i = i + 1
        rs_set.MoveNext
    Wend
    If i > 0 Then Exit Sub
    rs_set.Close
    Dim DoBVar, DayVar, MthVar, YrsVar As Integer
''    Dim filename As String
    CommonDialog1.Filter = "Text Files (*.txt)"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
''    filename = CommonDialog1.filename + ".txt"
''    Open "c:\rep.txt" For Output As #1
    Dim employer_pfamt As Double
    Open CommonDialog1.FileName + ".txt" For Output As #1
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        ecode = Trim(rs_set!emp_code)
''        If ecode = "169" Then
''         MsgBox "Jebes"
''        End If
       If rs_set!s_empcat = "W" Then
            eligible = (rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday)
            eligible2 = (rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday)
       Else
            eligible = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
            eligible2 = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        End If
        eps_wages = eligible
        
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff) / 2, 1)
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff), 1)
        If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
            empr_eligible = IIf(eligible >= 6500, 6500, eligible)
            eps_wages = IIf(eligible >= 6500, 6500, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible >= 6500, 6500, eligible)
               eps_wages = IIf(eligible >= 6500, 6500, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If eligible >= 6500 Then
                    pfamt = 780
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                End If
            End If
        Else
            empr_eligible = IIf(eligible > 15000, 15000, eligible)
            eps_wages = IIf(eligible > 15000, 15000, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible > 15000, 15000, eligible)
               eps_wages = IIf(eligible > 15000, 15000, eligible)
            End If
            If rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If rs_set!s_empcat = "W" Then
                   pfamt = Round(eligible * 12 / 100, 0)
                Else
                   If eligible > 15000 Then
                      pfamt = 1800
                   Else
                       pfamt = Round(rs_set!s_pf, 0)
                   End If
               End If
            End If
        
        End If
        
        employer_pfamt = Round(empr_eligible * 0.12, 0)
''Modified on 02/02/15
        epsamt = Round(eps_wages * 0.0833, 0)
        epsbalamt = Round(employer_pfamt - epsamt, 0)
        epsamt = 0
        epsbalamt = 0
        
        
        dt_joining.Value = Format(rs_set!emp_doj, "MM/dd/yyyy")
        
        If IsNull(rs_set!emp_resigneddate) Or Format(rs_set!emp_resigneddate, "MM/dd/yyyy") >= Format(end_date.Value, "MM/dd/yyyy") Then
           dt_resigned.Value = vbNull
        Else
           dt_resigned.Value = Format(rs_set!emp_resigneddate, "MM/dd/yyyy")
        End If
        
        DoBVar = IIf((100 * Month(end_date) + Day(end_date)) < (100 * Month(rs_set!emp_dob) + Day(rs_set!emp_dob)), 1, 0)
        MthVar = DateDiff("m", rs_set!emp_dob, end_date - DoBVar) Mod 12
        DayVar = DateDiff("d", rs_set!emp_dob, end_date)
        YrsVar = DateDiff("yyyy", rs_set!emp_dob, end_date) - DoBVar
        If YrsVar >= 58 And MthVar + DayVar >= 1 Then
            epsamt = 0
            epsbalamt = employer_pfamt
        End If
        LSet epfno = Trim(rs_set!emp_pfno)
        LSet empname = Trim(rs_set!emp_name)
        
        
        
''        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
''           LSet epfwages = Trim(Format$(rs_set!s_eligible_grosspay, "0"))
''        Else
''           LSet epfwages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
''        End If
        
''        LSet epswages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
           LSet epfwages = Trim(Format$(eligible2, "0"))
        Else
           If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
              LSet epfwages = Trim(Format$(IIf(eligible2 >= 6500, 6500, eligible2), "0"))
           Else
              LSet epfwages = Trim(Format$(IIf(eligible2 > 15000, 15000, eligible2), "0"))
           End If
        End If

''        LSet PF = Trim(Format$(rs_set!s_pf, "0"))
        LSet PF = Trim(Format$(pfamt, "0"))
        
        LSet epsamt2 = Trim(Format$(epsamt, "0"))
        LSet epsbalamt2 = Trim(Format$(epsbalamt, "0"))
        LSet ncp = Trim(Format$(lop, "0"))
        LSet pfrefund = Trim(Format$(rs_set!s_pfded, "0"))
     
''        If Format(rs_set!emp_doj, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_doj, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        
   ''     If Format(dt_joining.Value, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(dt_joining.Value, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        If dt_joining.Value >= Format(st_date.Value, "MM/dd/yyyy") And dt_joining <= Format(end_date.Value, "MM/dd/yyyy") Then
''        If dt_joining.Value >= st_date.Value And dt_joining.Value <= end_date.Value Then
           e_f_name = rs_set!emp_fname
           e_relation = IIf(rs_set!emp_relation = "F", "F", "S")
           e_dob = Format(rs_set!emp_dob, "dd/MM/yyyy")
           e_doJ = Format(rs_set!emp_doj, "dd/MM/yyyy")
           e_sex = rs_set!emp_sex
        Else
           e_f_name = ""
           e_relation = ""
           e_dob = ""
           e_doJ = ""
           e_sex = ""
        End If
        
        e_dol = ""
        e_reason = ""
        
        If Left(rs_set!emp_status, 1) = "R" Then
''            If dt_resigned.Value >= st_date.Value And dt_resigned.Value <= end_date.Value Then
            If dt_resigned.Value >= Format(st_date.Value, "MM/dd/yyyy") And dt_resigned.Value <= Format(end_date.Value, "MM/dd/yyyy") Then
                If rs_set.Fields("emp_reason") = "CESSATION" Then
                   e_reason = "C"
                ElseIf rs_set.Fields("emp_reason") = "SUPERANNUATION" Then
                   e_reason = "S"
                ElseIf rs_set.Fields("emp_reason") = "RETIREMENT" Then
                   e_reason = "R"
                ElseIf rs_set.Fields("emp_reason") = "DEATH IN SERVICE" Then
                   e_reason = "D"
                ElseIf rs_set.Fields("emp_reason") = "PERMANENT DISABLEMENT" Then
                   e_reason = "P"
                End If
                If IsNull(rs_set!emp_resigneddate) = True Then
                   e_dol = ""
                Else
                   e_dol = Format(rs_set!emp_resigneddate, "dd/MM/yyyy")
                End If
           End If
           End If
        LSet e_ar_epf_wages = "0"
        LSet e_ar_epf_ee = "0"
        LSet e_ar_epf_er = "0"
        LSet e_ar_eps = "0"
               
        sql = "Select * from arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & cmb_year & " and e_company = " & company_code & " and e_empcode  = '" & ecode & "'"
        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
        If Not payrs2.EOF Then
           arrearamt = payrs2("e_amount")
        End If
        payrs2.Close
        Print #1, Trim(epfno) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epswages) + "#~#" + Trim(PF) + "#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund) + "#~#" + Trim(e_ar_epf_wages) + "#~#" + Trim(e_ar_epf_ee) + "#~#" + Trim(e_ar_epf_er) + "#~#" + Trim(e_ar_eps) + "#~#" + Trim(e_f_name) + "#~#" + Trim(e_relation) + "#~#" + Trim(e_dob) + "#~#" + Trim(e_sex) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_reason)
        rs_set.MoveNext
    Wend
    Close #1
    MsgBox ("File Saved...")
    Dim path As String
    path = CommonDialog1.FileName + ".TXT"
    If prt_stat = False Then
        retval = Shell("Edit.com " & path, vbMaximizedFocus)
    End If
    Me.MousePointer = 1
End Sub
Public Sub pfreturn_employer_contribution()
    Dim ecode As String
    Dim epfno As String
    Dim empname As String
    Dim epfwages As String
    Dim epswages As String
    Dim PF As String
    Dim eps_contribution As String
    Dim eps_wages, pfamt, epsamt, epsbalamt, lop As Double
    Dim ncp As String
    Dim pfrefund As String
    Dim e_f_name As String
    Dim e_relation As String
    Dim e_dob As String
    Dim e_doJ As String
    Dim e_sex As String
    Dim e_dol As String
    Dim e_reason As String
    Dim e_ar_epf_wages As String
    Dim e_ar_epf_ee As String
    Dim e_ar_epf_er As String
    Dim e_ar_eps As String
    Dim empr_eligible As Double
    Dim eligible, eligible2 As Double
    epfno = Space(7)
    empname = Space(85)
    epfwages = Space(10)
    epswages = Space(10)
    eps_contribution = 10
    epsamt2 = Space(10)
    epsbalamt2 = Space(10)
    PF = Space(10)
    ncp = Space(5)
    pfrefund = Space(10)
    
    e_f_name = Space(60)
    e_relation = Space(1)
    e_dob = Space(12)
    e_doJ = Space(12)
    e_sex = Space(1)
    e_dol = Space(12)
    e_reason = Space(1)
    
    e_ar_epf_wages = Space(10)
    e_ar_epf_ee = Space(10)
    e_ar_epf_er = Space(10)
    e_ar_eps = Space(10)
    Dim i As Integer
    i = 0
    Dim rs_set As New ADODB.Recordset
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 and emp_pfno = '' order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        MsgBox ("PF Numebr Not available for the employee ..." + rs_set!emp_name)
        i = i + 1
        rs_set.MoveNext
    Wend
    If i > 0 Then Exit Sub
    rs_set.Close
    Dim DoBVar, DayVar, MthVar, YrsVar As Integer
''    Dim filename As String
    CommonDialog1.Filter = "Text Files (*.txt)"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
''    filename = CommonDialog1.filename + ".txt"
''    Open "c:\rep.txt" For Output As #1
    Dim employer_pfamt As Double
    Open CommonDialog1.FileName + ".txt" For Output As #1
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        ecode = Trim(rs_set!emp_code)
''        If ecode = "169" Then
''         MsgBox "Jebes"
''        End If
       If rs_set!s_empcat = "W" Then
            eligible = (rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday)
            eligible2 = (rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday)
       Else
            eligible = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
            eligible2 = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        End If
        eps_wages = eligible
        
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff) / 2, 1)
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff), 1)
        If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
            empr_eligible = IIf(eligible >= 6500, 6500, eligible)
            eps_wages = IIf(eligible >= 6500, 6500, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible >= 6500, 6500, eligible)
               eps_wages = IIf(eligible >= 6500, 6500, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If eligible >= 6500 Then
                    pfamt = 780
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                End If
            End If
        Else
            empr_eligible = IIf(eligible > 15000, 15000, eligible)
            eps_wages = IIf(eligible > 15000, 15000, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible > 15000, 15000, eligible)
               eps_wages = IIf(eligible > 15000, 15000, eligible)
            End If
            If rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
            Else
                If rs_set!s_empcat = "W" Then
                   pfamt = Round(eligible * 12 / 100, 0)
                Else
                   If eligible > 15000 Then
                      pfamt = 1800
                   Else
                       pfamt = Round(rs_set!s_pf, 0)
                   End If
               End If
            End If
        
        End If
        
        employer_pfamt = Round(empr_eligible * 0.12, 0)
''Modified on 02/02/15
        epsamt = Round(eps_wages * 0.0833, 0)
        epsbalamt = Round(employer_pfamt - epsamt, 0)
''        epsamt = 0
  ''      epsbalamt = 0
        
        
        dt_joining.Value = Format(rs_set!emp_doj, "MM/dd/yyyy")
        
        If IsNull(rs_set!emp_resigneddate) Or Format(rs_set!emp_resigneddate, "MM/dd/yyyy") >= Format(end_date.Value, "MM/dd/yyyy") Then
           dt_resigned.Value = vbNull
        Else
           dt_resigned.Value = Format(rs_set!emp_resigneddate, "MM/dd/yyyy")
        End If
        
        DoBVar = IIf((100 * Month(end_date) + Day(end_date)) < (100 * Month(rs_set!emp_dob) + Day(rs_set!emp_dob)), 1, 0)
        MthVar = DateDiff("m", rs_set!emp_dob, end_date - DoBVar) Mod 12
        DayVar = DateDiff("d", rs_set!emp_dob, end_date)
        YrsVar = DateDiff("yyyy", rs_set!emp_dob, end_date) - DoBVar
        If YrsVar >= 58 And MthVar + DayVar >= 1 Then
            epsamt = 0
            epsbalamt = employer_pfamt
        End If
        LSet epfno = Trim(rs_set!emp_pfno)
        LSet empname = Trim(rs_set!emp_name)
        
        
        
''        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
''           LSet epfwages = Trim(Format$(rs_set!s_eligible_grosspay, "0"))
''        Else
''           LSet epfwages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
''        End If
        
''        LSet epswages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
           LSet epfwages = Trim(Format$(eligible2, "0"))
        Else
           If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
              LSet epfwages = Trim(Format$(IIf(eligible2 >= 6500, 6500, eligible2), "0"))
           Else
              LSet epfwages = Trim(Format$(IIf(eligible2 > 15000, 15000, eligible2), "0"))
           End If
        End If

''        LSet PF = Trim(Format$(rs_set!s_pf, "0"))
        LSet PF = Trim(Format$(pfamt, "0"))
        
        LSet epsamt2 = Trim(Format$(epsamt, "0"))
        LSet epsbalamt2 = Trim(Format$(epsbalamt, "0"))
        LSet ncp = Trim(Format$(lop, "0"))
        LSet pfrefund = Trim(Format$(rs_set!s_pfded, "0"))
     
''        If Format(rs_set!emp_doj, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_doj, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        
   ''     If Format(dt_joining.Value, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(dt_joining.Value, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        If dt_joining.Value >= Format(st_date.Value, "MM/dd/yyyy") And dt_joining <= Format(end_date.Value, "MM/dd/yyyy") Then
''        If dt_joining.Value >= st_date.Value And dt_joining.Value <= end_date.Value Then
           e_f_name = rs_set!emp_fname
           e_relation = IIf(rs_set!emp_relation = "F", "F", "S")
           e_dob = Format(rs_set!emp_dob, "dd/MM/yyyy")
           e_doJ = Format(rs_set!emp_doj, "dd/MM/yyyy")
           e_sex = rs_set!emp_sex
        Else
           e_f_name = ""
           e_relation = ""
           e_dob = ""
           e_doJ = ""
           e_sex = ""
        End If
        
        e_dol = ""
        e_reason = ""
        
        If Left(rs_set!emp_status, 1) = "R" Then
''            If dt_resigned.Value >= st_date.Value And dt_resigned.Value <= end_date.Value Then
            If dt_resigned.Value >= Format(st_date.Value, "MM/dd/yyyy") And dt_resigned.Value <= Format(end_date.Value, "MM/dd/yyyy") Then
                If rs_set.Fields("emp_reason") = "CESSATION" Then
                   e_reason = "C"
                ElseIf rs_set.Fields("emp_reason") = "SUPERANNUATION" Then
                   e_reason = "S"
                ElseIf rs_set.Fields("emp_reason") = "RETIREMENT" Then
                   e_reason = "R"
                ElseIf rs_set.Fields("emp_reason") = "DEATH IN SERVICE" Then
                   e_reason = "D"
                ElseIf rs_set.Fields("emp_reason") = "PERMANENT DISABLEMENT" Then
                   e_reason = "P"
                End If
                If IsNull(rs_set!emp_resigneddate) = True Then
                   e_dol = ""
                Else
                   e_dol = Format(rs_set!emp_resigneddate, "dd/MM/yyyy")
                End If
           End If
           End If
        LSet e_ar_epf_wages = "0"
        LSet e_ar_epf_ee = "0"
        LSet e_ar_epf_er = "0"
        LSet e_ar_eps = "0"
               
        sql = "Select * from arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & cmb_year & " and e_company = " & company_code & " and e_empcode  = '" & ecode & "'"
        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
        If Not payrs2.EOF Then
           arrearamt = payrs2("e_amount")
        End If
        payrs2.Close
''        Print #1, Trim(epfno) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epswages) + "#~#" + Trim(PF) + "#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund) + "#~#" + Trim(e_ar_epf_wages) + "#~#" + Trim(e_ar_epf_ee) + "#~#" + Trim(e_ar_epf_er) + "#~#" + Trim(e_ar_eps) + "#~#" + Trim(e_f_name) + "#~#" + Trim(e_relation) + "#~#" + Trim(e_dob) + "#~#" + Trim(e_sex) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_reason)
        Print #1, Trim(epfno) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epswages) + "#~#" + "0" + "#~#" + "0" + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund) + "#~#" + Trim(e_ar_epf_wages) + "#~#" + Trim(e_ar_epf_ee) + "#~#" + Trim(e_ar_epf_er) + "#~#" + Trim(e_ar_eps) + "#~#" + Trim(e_f_name) + "#~#" + Trim(e_relation) + "#~#" + Trim(e_dob) + "#~#" + Trim(e_sex) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_doJ) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_dol) + "#~#" + Trim(e_reason)
        rs_set.MoveNext
    Wend
    Close #1
    MsgBox ("File Saved...")
    Dim path As String
    path = CommonDialog1.FileName + ".TXT"
    If prt_stat = False Then
        retval = Shell("Edit.com " & path, vbMaximizedFocus)
    End If
    Me.MousePointer = 1
End Sub


Public Sub find_dates()
    Dim mdays, diff As Integer
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
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


Public Sub pfreturn_new()
    Dim uan As String
    Dim ecode As String
    Dim epfno As String
    Dim empname As String
    Dim epfwages As String
    Dim epswages As String
    Dim PF As String
    Dim eps_contribution As String
    Dim eps_wages, pfamt, epsamt, epsbalamt, lop As Double
    Dim ncp As String
    Dim pfrefund As String
    Dim e_f_name As String
    Dim e_relation As String
    Dim e_dob As String
    Dim e_doJ As String
    Dim e_sex As String
    Dim e_dol As String
    Dim e_reason As String
    Dim e_ar_epf_wages As String
    Dim e_ar_epf_ee As String
    Dim e_ar_epf_er As String
    Dim e_ar_eps As String
    Dim empr_eligible As Double
    Dim eligible, eligible2 As Double
    
    uan = Space(12)
    epfno = Space(7)
    empname = Space(85)
    epfwages = Space(10)
    epswages = Space(10)
    eps_contribution = 10
    epsamt2 = Space(10)
    epsbalamt2 = Space(10)
    PF = Space(10)
    ncp = Space(5)
    pfrefund = Space(10)
    
    e_f_name = Space(60)
    e_relation = Space(1)
    e_dob = Space(12)
    e_doJ = Space(12)
    e_sex = Space(1)
    e_dol = Space(12)
    e_reason = Space(1)
    
    e_ar_epf_wages = Space(10)
    e_ar_epf_ee = Space(10)
    e_ar_epf_er = Space(10)
    e_ar_eps = Space(10)
    Dim i As Integer
    i = 0
    Dim rs_set As New ADODB.Recordset
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 and emp_uan = '' order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        MsgBox ("UAN Numebr Not available for the employee ..." + rs_set!emp_name)
        i = i + 1
        rs_set.MoveNext
    Wend
    If i > 0 Then Exit Sub
    rs_set.Close
    Dim DoBVar, DayVar, MthVar, YrsVar As Integer
''    Dim filename As String
    CommonDialog1.Filter = "Text Files (*.txt)"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
''    filename = CommonDialog1.filename + ".txt"
''    Open "c:\rep.txt" For Output As #1
    Dim employer_pfamt As Double
    Open CommonDialog1.FileName + ".txt" For Output As #1
    pst_qry = "select * from emp_salary a , emp_mas b where s_company = emp_company and s_empcode = emp_code and s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & "  and s_pf_eligible = 'Y' and s_pf >0 order by emp_pfno"
    rs_set.Open pst_qry, paydb, 1, 2
    While Not rs_set.EOF
        ecode = Trim(rs_set!emp_code)
''        If ecode = "5004" Then
''         MsgBox "SELVARAJ"
''        End If

      If rs_set!s_empcat = "W" Then
        eligible = Round((rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday), 2)
        eligible2 = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        eligible2 = Round((rs_set!s_basic + rs_set!s_serwt + rs_set!s_fda + rs_set!s_vda + rs_set!s_splpay) / rs_set!s_avlworkdays * (rs_set!s_actworkdays + rs_set!s_eligible_leave + rs_set!s_dec_holiday_eligible + rs_set!s_dec_holiday), 2)

      Else
        eligible = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
        eligible2 = rs_set!s_eligible_basic + rs_set!s_eligible_serwt + rs_set!s_eligible_fda + rs_set!s_eligible_vda + rs_set!s_eligible_splpay
      End If
        eps_wages = eligible
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + (rs_set!s_layoff) / 2, 1)
        lop = Round(rs_set!s_per_leave + rs_set!s_absent + rs_set!s_layoff, 1)
        If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
            empr_eligible = IIf(eligible >= 6500, 6500, eligible)
            eps_wages = IIf(eligible >= 6500, 6500, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible >= 6500, 6500, eligible)
               eps_wages = IIf(eligible >= 6500, 6500, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
                pfamt = Round(eligible * 0.12, 0)
            Else
                If eligible >= 6500 Then
                    pfamt = 780
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                    pfamt = Round(eligible * 0.12, 0)
                End If
            End If
        Else
            empr_eligible = IIf(eligible > 15000, 15000, eligible)
             ''empr_eligible = IIf(eligible > 15000, 1500, eligible)
            eps_wages = IIf(eligible > 15000, 15000, eligible)
            If rs_set!s_empcat = "S" Then
               eligible = IIf(eligible > 15000, 15000, eligible)
               eps_wages = IIf(eligible > 15000, 15000, eligible)
            End If
            If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
                pfamt = Round(rs_set!s_pf, 0)
                pfamt = Round(eligible * 0.12, 0)
            Else
                If eligible > 15000 Then
                    pfamt = 1800
                Else
                    pfamt = Round(rs_set!s_pf, 0)
                    pfamt = Round(eligible * 0.12, 0)
                End If
            End If
        
        End If
        
        employer_pfamt = Round(empr_eligible * 0.12, 0)
        
        epsamt = Round(eps_wages * 0.0833, 0)
               
        epsbalamt = Round(employer_pfamt - epsamt, 0)
        
        
        
        dt_joining.Value = Format(rs_set!emp_doj, "MM/dd/yyyy")
        
        If IsNull(rs_set!emp_resigneddate) Then
           dt_resigned.Value = vbNull
        Else
           dt_resigned.Value = Format(rs_set!emp_resigneddate, "MM/dd/yyyy")
        End If
        
        DoBVar = IIf((100 * Month(end_date) + Day(end_date)) < (100 * Month(rs_set!emp_dob) + Day(rs_set!emp_dob)), 1, 0)
        MthVar = DateDiff("m", rs_set!emp_dob, end_date - DoBVar) Mod 12
        DayVar = DateDiff("d", rs_set!emp_dob, end_date)
        YrsVar = DateDiff("yyyy", rs_set!emp_dob, end_date) - DoBVar
        If YrsVar >= 58 And MthVar + DayVar >= 1 Then
            epsamt = 0
            epsbalamt = employer_pfamt
        End If
        
        epswages = epfwages
        
        LSet uan = Trim(rs_set!emp_uan)
        
        LSet epfno = Trim(rs_set!emp_pfno)
        LSet empname = Trim(rs_set!emp_name)
        
        
        
''        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
''           LSet epfwages = Trim(Format$(rs_set!s_eligible_grosspay, "0"))
''        Else
''           LSet epfwages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
''        End If
        
''        LSet epswages = Trim(Format$(IIf(rs_set!s_eligible_grosspay >= 6500, 6500, rs_set!s_eligible_grosspay), "0"))
        If rs_set!s_empcat = "W" Or rs_set!s_empcat = "M" Then
           LSet epfwages = Trim(Format$(eligible2, "0"))
        Else
           If rs_set!s_month < 9 And rs_set!s_year <= 2014 Then
              LSet epfwages = Trim(Format$(IIf(eligible2 >= 6500, 6500, eligible2), "0"))
           Else
              LSet epfwages = Trim(Format$(IIf(eligible2 > 15000, 15000, eligible2), "0"))
           End If
        End If

''        LSet PF = Trim(Format$(rs_set!s_pf, "0"))
        LSet PF = Trim(Format$(pfamt, "0"))
        
        LSet epsamt2 = Trim(Format$(epsamt, "0"))
        LSet epsbalamt2 = Trim(Format$(epsbalamt, "0"))
        
        
        LSet ncp = Trim(Format$(lop, "0"))
        LSet pfrefund = Trim(Format$(rs_set!s_pfded, "0"))
     
''        If Format(rs_set!emp_doj, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_doj, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
''        If Format(dt_joining.Value, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(dt_joining.Value, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
        If dt_joining.Value >= st_date.Value And dt_joining.Value <= end_date.Value Then
        
           e_f_name = rs_set!emp_fname
           e_relation = IIf(rs_set!emp_relation = "F", "F", "S")
           e_dob = Format(rs_set!emp_dob, "dd/MM/yyyy")
           e_doJ = Format(rs_set!emp_doj, "dd/MM/yyyy")
           e_sex = rs_set!emp_sex
        Else
           e_f_name = ""
           e_relation = ""
           e_dob = ""
           e_doJ = ""
           e_sex = ""
        End If
        
        e_dol = ""
        e_reason = ""
        
        If Left(rs_set!emp_status, 1) = "R" Then
''           If Format(rs_set!emp_resigneddate, "MM/dd/yyyy") >= Format(st_date.Value, "MM/dd/yyyy") And Format(rs_set!emp_resigneddate, "MM/dd/yyyy") <= Format(end_date.Value, "MM/dd/yyyy") Then
            If dt_resigned.Value >= st_date.Value And dt_resigned.Value <= end_date.Value Then
                If rs_set.Fields("emp_reason") = "CESSATION" Then
                   e_reason = "C"
                ElseIf rs_set.Fields("emp_reason") = "SUPERANNUATION" Then
                   e_reason = "S"
                ElseIf rs_set.Fields("emp_reason") = "RETIREMENT" Then
                   e_reason = "R"
                ElseIf rs_set.Fields("emp_reason") = "DEATH IN SERVICE" Then
                   e_reason = "D"
                ElseIf rs_set.Fields("emp_reason") = "PERMANENT DISABLEMENT" Then
                   e_reason = "P"
                End If
                If IsNull(rs_set!emp_resigneddate) = True Then
                   e_dol = ""
                Else
                   e_dol = Format(rs_set!emp_resigneddate, "dd/MM/yyyy")
                End If
           End If
           End If
        LSet e_ar_epf_wages = "0"
        LSet e_ar_epf_ee = "0"
        LSet e_ar_epf_er = "0"
        LSet e_ar_eps = "0"
               
        sql = "Select * from arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & cmb_year & " and e_company = " & company_code & " and e_empcode  = '" & ecode & "'"
        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
        If Not payrs2.EOF Then
           arrearamt = payrs2("e_amount")
        End If
        payrs2.Close
        
        LSet epsamt2 = Trim(Format$(epsamt, "0"))
        LSet epsbalamt2 = Trim(Format$(epsbalamt, "0"))
        
''        Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epswages) + "#~#" + Trim(epswages) + "#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund) + "#~#"
''        If opt_new_all.Value = True Then
''           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
''        ElseIf opt_new_employee.Value = True Then
''           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(PF) + "#~#" + "0" + "#~#" + "0" + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
''        Else
''           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + "0" + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
''        End If
        
        If opt_new_all.Value = True Then
           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#0#~#" + Trim(PF) + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
        ElseIf opt_new_employee.Value = True Then
           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#0#~#" + Trim(PF) + "#~#" + "0" + "#~#" + "0" + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
        Else
           Print #1, Trim(uan) + "#~#" + Trim(empname) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#" + Trim(epfwages) + "#~#0#~#" + "0" + "#~#" + Trim(epsamt2) + "#~#" + Trim(epsbalamt2) + "#~#" + Trim(ncp) + "#~#" + Trim(pfrefund)
        End If
        
        rs_set.MoveNext
    Wend
    Close #1
    MsgBox ("File Saved...")
    Dim path As String
    path = CommonDialog1.FileName + ".TXT"
    If prt_stat = False Then
        retval = Shell("notepad " & path, vbMaximizedFocus)
    End If
    Me.MousePointer = 1
End Sub

