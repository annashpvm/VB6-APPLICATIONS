VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form salary_statement_worker_oldmonths 
   Caption         =   "Salary statement worker for Old Months"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   600
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   153419777
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   153419777
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "SALARY STATEMENT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   9240
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   840
         TabIndex        =   11
         Top             =   3720
         Width           =   7335
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
            Left            =   5280
            TabIndex        =   13
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
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4200
            TabIndex        =   15
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3480
         TabIndex        =   8
         Top             =   5040
         Width           =   1695
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "salary_statement_worker_oldmonths.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "salary_statement_worker_oldmonths.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.ListBox deduct_list 
         Height          =   285
         Left            =   7680
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT "
         Height          =   3105
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   7320
         Begin VB.OptionButton opt_wp 
            Caption         =   "PERMANENT WORKER - Consolidated"
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton opt_wt 
            Caption         =   "TRAINEE  WORKER - Consolidated"
            Height          =   300
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   3195
         End
         Begin VB.OptionButton opt_wp_pdays 
            Caption         =   "PERMANENT WORKER - for PRESENT DAYS"
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   4095
         End
         Begin VB.OptionButton opt_wt_pdays 
            Caption         =   "TRAINEE  WORKER  - for PRESENT DAYS"
            Height          =   300
            Left            =   240
            TabIndex        =   5
            Top             =   1680
            Width           =   4395
         End
         Begin VB.OptionButton opt_wp_layoff 
            Caption         =   "PERMANENT WORKER - for PRODUCTION STOPPAGE DAYS"
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   2040
            Width           =   5895
         End
         Begin VB.OptionButton opt_wt_layoff 
            Caption         =   "TRAINEE  WORKER  - for PRODUCTION STOPPAGE DAYS"
            Height          =   300
            Left            =   240
            TabIndex        =   3
            Top             =   2400
            Width           =   5595
         End
      End
      Begin VB.ListBox std_deduct_lst 
         Height          =   285
         Left            =   7920
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   4560
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "SELECT DEDUCTION LIST "
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
         Left            =   4920
         TabIndex        =   16
         Top             =   4680
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   360
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      Caption         =   "DONT USE THIS REPORT FOR  JUNE-2015 AND AFTER MONTHS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "salary_statement_worker_oldmonths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_month_Click()
find_dates
End Sub

Private Sub cmb_year_Click()
find_dates
End Sub

Private Sub Command1_Click()
Set salary_statement_staff.rpt = rptprint
With salary_statement_staff.rpt
    .ExportOptions.DiskFileName = "xx.pdf"
    .ExportOptions.DestinationType = crEDTDiskFile
    .ExportOptions.FormatType = crEFTPortableDocFormat
    .ExportOptions.PDFExportAllPages = True
    .Export False
    .PrinterSetup 0
End With

End Sub

Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
''    opt_sp.Enabled = True


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
''''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
''    cmb_year.Text = "2015"
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
      
    
End Sub
''
''Private Sub opt_sp_Click()
''    sql = ("Select * from  pdedu_mas order by pdedu_name")
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    std_deduct_lst.Clear
''    deduct_list.Clear
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 2 Then
''           std_deduct_lst.AddItem payrs(1)
''           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
''        End If
''        If payrs.Fields("pdedu_type") = 4 Then
''           deduct_list.AddItem payrs(1)
''           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
''        End If
''        payrs.MoveNext
''    Wend
''End Sub
''
''Private Sub opt_wp_Click()
''    sql = ("Select * from  pdedu_mas order by pdedu_name")
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    std_deduct_lst.Clear
''    deduct_list.Clear
''
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 3 Then
''           std_deduct_lst.AddItem payrs(1)
''           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
''        End If
''        If payrs.Fields("pdedu_type") = 4 Then
''           deduct_list.AddItem payrs(1)
''           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
''        End If
''        payrs.MoveNext
''    Wend
''End Sub

Private Sub PROCESS_Click()

   ded1_code = " "
   ded2_code = " "
   ded3_code = " "
   ded4_code = " "
   ded5_code = " "
   ded6_code = " "
   ded7_code = " "
   ded8_code = " "
   ded9_code = " "
   ded10_code = " "
   If opt_wp.Value = True Then disname = "WORKER WAGES STATEMENT FOR THE MONTH OF "
   If opt_wp_pdays.Value = True Then disname = "WORKER WAGES STATEMENT FOR THE MONTH OF "
   If opt_wp_layoff.Value = True Then disname = "WORKER WAGES STATEMENT FOR THE MONTH OF "

   If opt_wt.Value = True Then disname = "TRAINEE WORKER WAGES STATEMENT FOR THE MONTH OF "
   If opt_wt_layoff.Value = True Then disname = "TRAINEE WORKER WAGES STATEMENT FOR THE MONTH OF "
   If opt_wt_pdays.Value = True Then disname = "TRAINEE WORKER WAGES STATEMENT FOR THE MONTH OF "
''   If opt_worker_all.Value = True Then disname = "WORKER WAGES STATEMENT FOR THE MONTH OF "
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   Dim ds As String
''   If optchk = 2 Then
''      ds = " and {emp_mas.emp_classification} = 'A'"
''   ElseIf optchk = 1 Then
''      ds = " and {emp_mas.emp_classification} = 'B'"
''   Else
''    ds = ""
''   End If
   cry_rep1.Formulas(5) = ""
   If optchk = 1 Then
''      ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_classification} = 'B'"
      ds = " and {emp_mas.emp_workplace} = 'MILL'"
      cry_rep1.Formulas(5) = ("e_cat = ' VPT -  BELOW MANAGERS   '")
   ElseIf optchk = 2 Then
      ds = " and {emp_mas.emp_workplace} = 'CBE' "
      cry_rep1.Formulas(5) = ("e_cat = ' CBE - STAFF SALARY  '")
   ElseIf optchk = 3 Then
      ds = " and {emp_mas.emp_classification} = 'A'"
      cry_rep1.Formulas(5) = ("e_cat = ' ABOVE MANAGERS   '")
   Else
      ds = ""
   End If
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   cry_rep1.Formulas(6) = ""
   
   If opt_wp.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_a4_foroldmonths.rpt"
      cry_rep1.Formulas(5) = ""
   ElseIf opt_wt.Value = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_trainee_a4_foroldmonths.rpt"
        cry_rep1.Formulas(5) = ""
   ElseIf opt_wp_pdays = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_forlayoff_foroldmonths.rpt"
        cry_rep1.Formulas(5) = ("cond = 0")
        cry_rep1.Formulas(6) = ("speriod = 'FOR WORKING DAYS'")
   ElseIf opt_wt_pdays = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_trainee_forlayoff_foroldmonths.rpt"
        cry_rep1.Formulas(5) = ("cond = 0")
        cry_rep1.Formulas(6) = ("speriod = 'FOR WORKING DAYS'")
   ElseIf opt_wp_layoff = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_forlayoff_foroldmonths.rpt"
        cry_rep1.Formulas(5) = ("cond = 1")
        cry_rep1.Formulas(6) = ("speriod = 'FOR PRODUCTION STOPPAGE DAYS'")
    ElseIf opt_wt_layoff = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_trainee_forlayoff_foroldmonths.rpt"
        cry_rep1.Formulas(5) = ("cond = 1")
        cry_rep1.Formulas(6) = ("speriod = 'FOR PRODUCTION STOPPAGE DAYS'")
   
   End If
   
   If opt_wp.Value = True Or opt_wp_pdays.Value = True Or opt_wp_layoff.Value = True Then
        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2 and {emp_salary.s_salarydays} > 0  " & ds & "")
''   ElseIf opt_worker_all = True Then
''        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''        "and {emp_salary.s_company} = " & company_code & " and ({emp_mas.emp_type} = 2 or {emp_mas.emp_type} = 3 ) and {emp_salary.s_salarydays} > 0  " & ds & "")
   Else
        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 and {emp_salary.s_salarydays} > 0  " & ds & "")
   End If
   
'   If opt_sp.Value = True Then
'      pst_qry = ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 0 and {emp_salary.s_salarydays} > 0  " & ds & " and ({emp_mas.emp_status}  ='A' or ({emp_mas.emp_status}  ='R' and month({emp_mas.emp_resigneddate}) <=  " & cmb_month.ItemData(cmb_month.ListIndex) & " and year({emp_mas.emp_resigneddate}) <=  " & Val(cmb_year.Text) & "  ))")
'      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'                                        " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 0 and {emp_salary.s_salarydays} > 0  " & ds & " ")
'   Else
'      If opt_st.Value = True Then
'         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1 and {emp_salary.s_salarydays} > 0  " & ds & "")
'      Else
'         If opt_wp.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2 and {emp_salary.s_salarydays} > 0  " & ds & "")
''         Else
''            If opt_wt.Value = True Then
''               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                       "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 and {emp_salary.s_salarydays} > 0  " & ds & "")
''            End If
''         End If
''      End If
''   End If
   c = 1
   Dim item As String
   pin_fms = 5
   If std_deduct_lst.ListCount > 0 Then
      For pin_row = 0 To std_deduct_lst.ListCount - 1
''          If std_deduct_lst.Selected(pin_row) = True Then
            cry_rep1.Formulas(pin_fms) = "ded" & (c) & "_code = " & Val(std_deduct_lst.ItemData(pin_row))
            std_deduct_lst.ListIndex = Val(pin_row)
            cry_rep1.Formulas(pin_fms + 1) = "ded" & (c) & " = '" & Trim$(std_deduct_lst.Text) & "'"
            pin_sel_item = Val(pin_sel_item) + 1
            c = c + 1
            pin_fms = pin_fms + 2
''          End If
      Next pin_row
   End If
   If deduct_list.ListCount > 0 Then
      For pin_row = 0 To deduct_list.ListCount - 1
          If pin_fms > 21 Then Exit For
          If deduct_list.Selected(pin_row) = True Then
             cry_rep1.Formulas(pin_fms) = "ded" & (c) & "_code = " & Val(deduct_list.ItemData(pin_row))
             deduct_list.ListIndex = Val(pin_row)
             cry_rep1.Formulas(pin_fms + 1) = "ded" & (c) & " = '" & Trim$(deduct_list.Text) & "'"
             pin_sel_item = Val(pin_sel_item) + 1
             c = c + 1
             pin_fms = pin_fms + 2
          Else
             cry_rep1.Formulas(pin_fms) = ""
             cry_rep1.Formulas(pin_fms + 1) = ""
             pin_fms = pin_fms + 2
          End If
        Next pin_row
   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
 End Sub


Public Sub find_dates()
''    If cmb_month.ListIndex = -1 Then Exit Sub
    If cmb_year.Text = "" Then Exit Sub
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




