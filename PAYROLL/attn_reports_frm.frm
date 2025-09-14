VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form attn_reports_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ATTENDANCE REPORTS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   5175
      Left            =   9960
      TabIndex        =   25
      Top             =   840
      Width           =   7335
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton opt_selective_dept 
            Caption         =   "Selective"
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
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton opt_alldept 
            Caption         =   "All"
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
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   2160
         TabIndex        =   26
         Top             =   120
         Width           =   5055
         Begin VB.ListBox lst_dept 
            Enabled         =   0   'False
            Height          =   4110
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SELECT SEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   16200
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton Opt_MALE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "MALE"
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
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Opt_female 
         BackColor       =   &H00C0E0FF&
         Caption         =   "FEMALE"
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
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Opt_all 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ALL"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " "
      Height          =   735
      Left            =   1680
      TabIndex        =   13
      Top             =   8760
      Visible         =   0   'False
      Width           =   5655
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121438209
         CurrentDate     =   42187
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121438209
         CurrentDate     =   42187
      End
      Begin VB.Label Label3 
         Caption         =   "FROM"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "TO"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ATTENDANCE   - REPORTS"
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
      Height          =   7695
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   9390
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   2520
         TabIndex        =   31
         Top             =   6120
         Width           =   3375
         Begin VB.CommandButton Exit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "E&xit"
            Height          =   750
            Left            =   2280
            Picture         =   "attn_reports_frm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Refresh 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Refresh"
            Height          =   750
            Left            =   1200
            Picture         =   "attn_reports_frm.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton print 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Print"
            Height          =   750
            Left            =   120
            Picture         =   "attn_reports_frm.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   15
         Left            =   5520
         TabIndex        =   11
         Top             =   6000
         Width           =   615
      End
      Begin VB.Frame rep_opt 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SELECT THE REPORT"
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
         Height          =   2505
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   9015
         Begin VB.OptionButton opt_forthe_month_absent 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Absent  Leave For the MONTH - Emp wise"
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
            Height          =   345
            Left            =   240
            TabIndex        =   35
            Top             =   840
            Width           =   5490
         End
         Begin VB.OptionButton opt_forthe_uptomonth 
            BackColor       =   &H00C0E0FF&
            Caption         =   "UPTO CURRENT MONTH - Empwise"
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
            Height          =   345
            Left            =   240
            TabIndex        =   19
            Top             =   2040
            Width           =   3810
         End
         Begin VB.OptionButton opt_forthe_month 
            BackColor       =   &H00C0E0FF&
            Caption         =   "For the  MONTH - Emp wise"
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
            Height          =   345
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   3810
         End
         Begin VB.OptionButton opt_date 
            BackColor       =   &H00C0E0FF&
            Caption         =   "DATE"
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
            Height          =   450
            Left            =   6240
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton opt_umon 
            BackColor       =   &H00C0E0FF&
            Caption         =   "UPTO CURRENT MONTH - Departmentwise"
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
            Height          =   345
            Left            =   4920
            TabIndex        =   2
            Top             =   1800
            Visible         =   0   'False
            Width           =   5490
         End
         Begin VB.OptionButton opt_fmon 
            BackColor       =   &H00C0E0FF&
            Caption         =   "CURRENT MONTH - Department wise"
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
            Height          =   345
            Left            =   5160
            TabIndex        =   1
            Top             =   1080
            Visible         =   0   'False
            Width           =   3570
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SELECT STAFF / WORKER"
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
         Height          =   1020
         Left            =   1080
         TabIndex        =   8
         Top             =   3480
         Width           =   6510
         Begin VB.OptionButton opt_sw 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00C0E0FF&
            Caption         =   "STAFF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   2160
            TabIndex        =   10
            Top             =   360
            Width           =   1770
         End
         Begin VB.OptionButton opt_worker 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   1770
         End
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
         Left            =   6600
         TabIndex        =   5
         Top             =   4800
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
         Left            =   2520
         TabIndex        =   4
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   1320
         TabIndex        =   7
         Top             =   4920
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   5520
         TabIndex        =   6
         Top             =   4920
         Width           =   885
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4005
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "attn_reports_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    sql = "select * from pdept_mas   order by dept_name"
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_dept.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("dept_name")
        payrs.MoveNext
    Wend
        
        
    If at_rep_opt = 1 Then
       rep_opt.Visible = True
    End If
    opt_sw.Value = True

    opt_forthe_month.Value = True
    st_date.Value = Now
    end_date.Value = Now
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
    sql = ("Select * from  emp_salary")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic

End Sub

Private Sub opt_alldept_Click()
     lst_dept.Enabled = False
End Sub

Private Sub opt_date_Click()
Frame5.Visible = True
End Sub

Private Sub opt_fmon_Click()
 Frame5.Visible = False
End Sub

Private Sub opt_selective_dept_Click()
     lst_dept.Enabled = True
End Sub

Private Sub opt_umon_Click()
 Frame5.Visible = False
End Sub

Private Sub print_Click()
   Dim dept, sw As String
   dept = ""
   If opt_selective_dept.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                     dept = " and ( {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                     i = i + 1
                  Else
                     dept = dept + " or {pdept_mas.dept_name}= '" & lst_dept.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If

   If dept <> "" Then dept = dept + ")"
   sw = ""
   If opt_staff.Value = True Then
      disname = "STAFFS ATTENDANCE DETAILS FOR THE MONTH OF "
      sw = " and {emp_salary.s_empcat} = 'S'"
   ElseIf opt_worker.Value = True Then
      disname = "WORKERS ATTENDANCE DETAILS FOR THE MONTH OF "
      sw = " and {emp_salary.s_empcat} = 'W'"
   ElseIf opt_sw.Value = True Then
      disname = "ATTENDANCE DETAILS FOR THE MONTH OF "
   End If
   
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
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\attn_summary_report.rpt"
  If opt_forthe_month.Value = True Then
     cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0 " & sw & dept)
  ElseIf opt_forthe_month_absent.Value = True Then
     cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\attn_summary_report_absent.rpt"
     cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_avlworkdays} > {emp_salary.s_salarydays} and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0 " & sw & dept)

  
  ElseIf opt_forthe_uptomonth.Value = True Then
     cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0 " & sw & dept)
  
  End If

   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1

''   If at_rep_opt = 0 Then
''      If opt_staff.Value = True Then disname = "STAFFS ATTENDANCE STATEMENT FOR THE MONTH OF "
''      If opt_worker.Value = True Then disname = "WORKERS ATTENDANCE STATEMENT FOR THE MONTH OF "
''           If opt_sw.Value = True Then disname = "ATTENDANCE DETAILS FOR THE MONTH OF "
''   Else
''      If opt_umon.Value = True Then
''         If opt_staff.Value = True Then disname = "STAFFS ATTENDANCE SUMMARY UPTO THE MONTH OF "
''         If opt_worker.Value = True Then disname = "WORKERS ATTENDANCE SUMMARY UPTO THE MONTH OF "
''         If opt_sw.Value = True Then disname = "ATTENDANCE SUMMARY UPTO THE MONTH OF "
''       ElseIf opt_date.Value = True Then
''         If opt_staff.Value = True Then disname = "STAFFS ATTENDANCE SUMMARY FOR "
''         If opt_worker.Value = True Then disname = "WORKERS ATTENDANCE SUMMARY FOR "
''         If opt_sw.Value = True Then disname = "ATTENDANCE SUMMARY FOR  "
''      Else
''         If opt_staff.Value = True Then disname = "STAFFS ATTENDANCE DETAILS FOR THE MONTH OF "
''         If opt_worker.Value = True Then disname = "WORKERS ATTENDANCE DETAILS FOR THE MONTH OF "
''         If opt_sw.Value = True Then disname = "ATTENDANCE DETAILS FOR THE MONTH OF "
''      End If
''   End If
''   If Trim(cmb_month.Text) = "" Then
''      MsgBox ("Select the Reporting Month")
''      Exit Sub
''   End If
''
''

''   MousePointer = vbDefault
''   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   cry_rep1.PrinterSelect
''   cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
''   cry_rep1.Formulas(1) = ("sthead = '" & disname & "'")
''   If opt_date.Value = True Then
''        cry_rep1.Formulas(2) = ("rmonth = '" & Format(st_date.Value, "dd/MM/yyyy") & " TO " & Format(end_date.Value, "dd/MM/yyyy") & "'")
''   Else
''        cry_rep1.Formulas(2) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
''   End If
''   cry_rep1.Formulas(3) = ("r2month = " & cmb_month.ItemData(cmb_month.ListIndex))
''   If at_rep_opt = 0 Then
''         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\attendance_report.rpt"
''      If opt_staff.Value = True And Opt_all.Value = True Then
''           pst_qry = ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                                " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and  {emp_salary.s_salarydays} > 0 ")
''
''           cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                                " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and  {emp_salary.s_salarydays} > 0 ")
''
''      ElseIf opt_staff.Value = True And Opt_MALE.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                                " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'M' and  {emp_salary.s_salarydays} > 0 ")
''
''      ElseIf opt_staff.Value = True And Opt_female.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                                " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''      ElseIf opt_worker.Value = True And Opt_all.Value = True Then
''           cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W'  and  {emp_salary.s_salarydays} > 0 ")
''
''      ElseIf opt_worker.Value = True And Opt_MALE.Value = True Then
''           cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'M' and {emp_salary.s_salarydays} > 0 ")
''
''      ElseIf opt_worker.Value = True And Opt_female.Value = True Then
''           cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'F'  and  {emp_salary.s_salarydays} > 0 ")
''
''      End If
''   Else
''      If opt_forthe_month.Value = True Or opt_forthe_uptomonth.Value = True Then
''          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\attn_summary_report.rpt"
''      Else
''          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\attn_summary_report_dept.rpt"
''      End If
''      If opt_staff.Value = True And Opt_all.Value = True Then
''         If opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_salary.s_salarydays} > 0 ")
''         Else
''              cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and  {emp_salary.s_salarydays} > 0 ")
''         End If
''       ElseIf opt_staff.Value = True And Opt_MALE.Value = True Then
''         If opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'M' and {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'M' and {emp_salary.s_salarydays} > 0 ")
''         Else
''              cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'M' and {emp_salary.s_salarydays} > 0 ")
''         End If
''       ElseIf opt_staff.Value = True And Opt_female.Value = True Then
''         If opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month}<= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''         Else
''              cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''         End If
''      ElseIf opt_worker.Value = True And Opt_all.Value = True Then
''         If opt_forthe_uptomonth.Value = True Or opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_salary.s_salarydays} > 0 ")
''
''         Else
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_salary.s_salarydays} > 0 ")
''
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' ")
''
''
''
''         End If
''      ElseIf opt_worker.Value = True And Opt_MALE.Value = True Then
''         If opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'M' and  {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'M' and  {emp_salary.s_salarydays} > 0 ")
''         Else
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'M' and  {emp_salary.s_salarydays} > 0 ")
''         End If
''      ElseIf opt_worker.Value = True And Opt_female.Value = True Then
''         If opt_umon.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} >= 1 and {emp_salary.s_month} <= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''         ElseIf opt_date.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} >= " & year(st_date.Value) & " and {emp_salary.s_year} <= " & year(end_date.Value) & " and {emp_salary.s_month} >= " & Month(st_date.Value) & " and {emp_salary.s_month}<= " & Month(end_date.Value) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'F' and  {emp_salary.s_salarydays} > 0 ")
''         Else
''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_mas.emp_sex} = 'F' and {emp_salary.s_salarydays} > 0 ")
''         End If
''      End If
''   End If
''
''
''   cry_rep1.WindowState = crptMaximized
''   cry_rep1.Connect = gst_repconnect
''   cry_rep1.Action = 1
End Sub

Private Sub refresh_Click()
    opt_staff.Value = True
    opt_fmon.Value = True
End Sub


