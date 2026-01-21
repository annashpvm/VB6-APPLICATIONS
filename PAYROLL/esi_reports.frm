VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form esi_reports_frm 
   Caption         =   "ESI Reports"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   1080
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   5895
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   8055
         Begin VB.OptionButton opt_esi_dept 
            Caption         =   "ESI Statement - Department wise"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   3855
         End
         Begin VB.OptionButton opt_rep2 
            Caption         =   "ESI Statment - for online file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   480
            TabIndex        =   17
            Top             =   1200
            Width           =   3015
         End
         Begin VB.OptionButton opt_rep1 
            Caption         =   "ESI Statement"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   3855
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
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   570
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
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
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   2880
         TabIndex        =   7
         Top             =   4200
         Width           =   3015
         Begin VB.CommandButton print 
            Caption         =   "&Print"
            Height          =   870
            Left            =   0
            Picture         =   "esi_reports.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Refresh 
            Caption         =   "&Refresh"
            Height          =   870
            Left            =   960
            Picture         =   "esi_reports.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Exit 
            Caption         =   "&Exit"
            Height          =   870
            Left            =   1920
            Picture         =   "esi_reports.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   2760
         Width           =   8175
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
            Left            =   1860
            TabIndex        =   4
            Top             =   315
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
            Left            =   6150
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
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
            TabIndex        =   6
            Top             =   315
            Width           =   885
         End
         Begin VB.Label Label2 
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
            TabIndex        =   5
            Top             =   345
            Width           =   1050
         End
      End
      Begin VB.Label Label1 
         Caption         =   " ESI STATEMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   615
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "esi_reports_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
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

Private Sub print_Click()
 
   disname = "MONTHLY ESI STATEMENT FOR THE MONTH OF "
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
   If opt_rep1.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\ESI_statement.rpt"
   ElseIf opt_esi_dept.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\ESI_statement_deptwise.rpt"
   Else
   
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\esi_statement_forfile.rpt"
   End If
    
    If opt_staff.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_mas.emp_cat} = 'S'  and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_esi_ded} > 0")
      ElseIf opt_worker.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_mas.emp_cat} = 'W' and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_esi_ded} > 0")
      Else
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_esi_ded} > 0")
''         Cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                            "and {emp_salary.s_company} = " & company_code & "  and  {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_esi_ded} > 0")
      End If
     

  
''''      If opt_staff.Value = True Then
''''         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''''                                            "and {emp_salary.s_company} = " & company_code & " and ( {emp_mas.emp_cat} = 'S' or {emp_mas.emp_cat} = 'M') and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
''''      ElseIf opt_worker.Value = True Then
''''         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''''                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_mas.emp_cat} = 'W' and {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
''''      Else
''''         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''''                                            "and {emp_salary.s_company} = " & company_code & " and  {emp_salary.s_eligible_grosspay} > 0 and {emp_salary.s_pf} > 0")
''''      End If
''''
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub

Private Sub refresh_Click()
opt_All.Value = True
End Sub
