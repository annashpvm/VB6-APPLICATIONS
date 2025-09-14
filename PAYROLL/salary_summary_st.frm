VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salary_summary_st 
   Caption         =   "SALARY STATEMENTS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "SELECT"
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
      Height          =   1095
      Left            =   2625
      TabIndex        =   5
      Top             =   2610
      Width           =   6510
      Begin VB.OptionButton rep_p 
         Caption         =   "PERMANENT"
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
         Height          =   495
         Left            =   270
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton rep_all 
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
         Height          =   495
         Left            =   4545
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton rep_t 
         Caption         =   "TRAINEE"
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
         Height          =   495
         Left            =   2295
         TabIndex        =   7
         Top             =   375
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SALARY STATEMENTS"
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
      Height          =   6555
      Left            =   1260
      TabIndex        =   0
      Top             =   900
      Width           =   9390
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   990
         Left            =   3480
         Picture         =   "salary_summary_st.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4920
         Width           =   1110
      End
      Begin VB.CommandButton Refresh 
         Caption         =   "&Refresh"
         Height          =   990
         Left            =   2400
         Picture         =   "salary_summary_st.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4920
         Width           =   1110
      End
      Begin VB.CommandButton print 
         Caption         =   "&Print"
         Height          =   990
         Left            =   1290
         Picture         =   "salary_summary_st.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4920
         Width           =   1110
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
         Left            =   2745
         TabIndex        =   9
         Top             =   3210
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
         Left            =   6540
         TabIndex        =   10
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Frame Frame2 
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
         Height          =   1095
         Left            =   1365
         TabIndex        =   1
         Top             =   420
         Width           =   6510
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
            Height          =   450
            Left            =   4680
            TabIndex        =   4
            Top             =   390
            Width           =   1170
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
            Height          =   450
            Left            =   345
            TabIndex        =   2
            Top             =   285
            Width           =   1770
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
            Height          =   450
            Left            =   2280
            TabIndex        =   3
            Top             =   360
            Width           =   1770
         End
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
         Left            =   5535
         TabIndex        =   15
         Top             =   3240
         Width           =   885
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
         Left            =   1380
         TabIndex        =   13
         Top             =   3255
         Width           =   1050
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   3285
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "salary_summary_st"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    opt_staff.Value = True
    rep_p.Value = True
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
''        .AddItem "2007"
''        .AddItem "2008"
''        .AddItem "2009"
''        .AddItem "2010"
''        .AddItem "2011"
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

Private Sub print_Click()
   If opt_staff.Value = True Then
      If rep_p.Value = True Then disname = "PERMENANT STAFFS SALARY SUMMARY FOR THE MONTH OF "
      If rep_t.Value = True Then disname = "TRAINEE STAFFS SALARY SUMMARY FOR THE MONTH OF "
      If rep_all.Value = True Then disname = "ALL STAFFS SALARY SUMMARY FOR THE MONTH OF "
   End If
   If opt_worker.Value = True Then
      If rep_p.Value = True Then disname = "PERMENANT WORKERS SALARY SUMMARY FOR THE MONTH OF "
      If rep_t.Value = True Then disname = "TRAINEE WORKERS SALARY SUMMARY FOR THE MONTH OF "
      If rep_all.Value = True Then disname = "ALL WORKERS SALARY SUMMARY FOR THE MONTH OF "
   End If
   If opt_All.Value = True Then
      If rep_p.Value = True Then disname = "PERMENANT STAFFS & WORKERS SALARY SUMMARY FOR THE MONTH OF "
      If rep_t.Value = True Then disname = "TRAINEE STAFFS & WORKERS SALARY SUMMARY FOR THE MONTH OF "
      If rep_all.Value = True Then disname = "ALL STAFFS & WORKERS SALARY SUMMARY FOR THE MONTH OF "
   End If
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   gst_repconnect = "dsn=servall;uid=sa;pwd=serdat"
     cry_rep1.PrinterSelect
   cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(1) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(2) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\salary_summary_reports.rpt"
   If opt_staff.Value = True Then
         If rep_p.Value = True Then
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                               "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_emptype} = " & 0)
''''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''''                                               "and {emp_salary.s_company} = " & company_code & " ")
''
''                    cry_rep1.ReplaceSelectionFormula ("")
         Else
           If rep_t.Value = True Then
              cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                 "and {emp_salary.s_company} = " & company_code & " and  {emp_salary.s_emptype} = " & 1)
           Else
              cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                 "and {emp_salary.s_company} = " & company_code & " and ({emp_salary.s_emptype} = 0 or {emp_salary.s_emptype} = 1)")
           End If
        End If
   Else
        If opt_worker.Value = True Then
             If rep_p.Value = True Then
                cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                   "and {emp_salary.s_company} = " & company_code & " and  {emp_salary.s_emptype} = " & 2)
             Else
               If rep_t.Value = True Then
                  cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                     "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_emptype} = " & 3)
               Else
                  cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                 "and {emp_salary.s_company} = " & company_code & " and  ({emp_salary.s_emptype} =  2 or {emp_salary.s_emptype} = 3)")
                                                     
               End If
            End If
        Else
            If rep_p.Value = True Then
                cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                   "and {emp_salary.s_company} = " & company_code & " and  ({emp_salary.s_emptype} =  0 or {emp_salary.s_emptype} = 2)")
            Else
               If rep_t.Value = True Then
                  cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                     "and {emp_salary.s_company} = " & company_code & " and ({emp_salary.s_emptype} = 1 or {emp_salary.s_emptype} = 3)")
               Else
                  cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                                     "and {emp_salary.s_company} = " & company_code & "")
                 

               End If
            End If
        End If
   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub

Private Sub refresh_Click()
    opt_staff.Value = True
    rep_p.Value = True
End Sub




    
