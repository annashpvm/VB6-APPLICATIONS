VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form worked_days_statement 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "WORKED && LAYOFF DAYS STATEMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5115
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   10575
      Begin VB.Frame Frame4 
         Caption         =   "REPORTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   9615
         Begin VB.OptionButton opt_rep2 
            Caption         =   "LOP Statement"
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
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton opt_rep1 
            Caption         =   "Worked and Layoff Statement"
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
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   5775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "DETAILS FOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   9585
         Begin VB.OptionButton opt_wp 
            Caption         =   "PERMANENT WORKER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5220
            TabIndex        =   12
            Top             =   480
            Width           =   1950
         End
         Begin VB.OptionButton opt_sp 
            Caption         =   "PERMANENT STAFF "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   285
            TabIndex        =   11
            Top             =   360
            Width           =   1770
         End
         Begin VB.OptionButton opt_st 
            Caption         =   "TRAINEE STAFF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2640
            TabIndex        =   10
            Top             =   480
            Width           =   1740
         End
         Begin VB.OptionButton opt_wt 
            Caption         =   "TRAINEE WORKER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   7320
            TabIndex        =   9
            Top             =   480
            Width           =   1905
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ACTIVE / RESIGNED"
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
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   4200
         Width           =   9495
         Begin VB.OptionButton opt_emptype_resigned 
            Caption         =   "RESIGNED"
            Height          =   375
            Left            =   6840
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptype_active 
            Caption         =   "ACTIVE"
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptypeall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "EXIT"
         Height          =   705
         Left            =   840
         Picture         =   "worked_days_statement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton PRINT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "PRINT"
         Height          =   705
         Left            =   120
         Picture         =   "worked_days_statement.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "worked_days_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    opt_sp.Value = True
''    With cmb_year
''''        .AddItem "2011-2012"
''        .AddItem "2012-2013"
''        .AddItem "2013-2014"
''    End With
''''    cmb_year.Text = "2007-2008"
End Sub

Private Sub print_Click()
   If opt_sp.Value = True Then disname = "PERMENENT STAFF STATEMENT FOR THE YEAR OF "
   If opt_st.Value = True Then disname = "TEMPORARY STAFF STATEMENT FOR THE YEAR OF "
   If opt_wp.Value = True Then disname = "PERMENENT WORKER STATEMENT FOR THE YEAR OF "
   If opt_wt.Value = True Then disname = "TEMPORARY WORKER STATEMENT FOR THE YEAR OF "
''   syear = Val(Left(cmb_year.Text, 4))
''   eyear = Val(Mid(cmb_year.Text, 6, 4))
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
''   If opt3.Value = True Then
''      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_for3500.rpt"
''   ElseIf opt2.Value = True Then
''      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_existing.rpt"
''   ElseIf opt4.Value = True Then
''      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_for3500p_dhless.rpt"
''   Else
   
   Dim ryear As Integer
   ryear = Left(fyear, 4)
   
   
   cry_rep1.Formulas(0) = ("report_year = '" & ryear & "'")
''   cry_rep1.Formulas(1) = ("bonus_per = " & Val(txt_bonus.Text) & "")
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
   
     If opt_rep1.Value = True Then
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\worked_days_statement.rpt"
     Else
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\lop_days_statement.rpt"
         cry_rep1.Formulas(3) = ""

     End If
''   End If
''   cry_rep1.ReportFileName = "\\annadurai\d\payroll\bonus_statement.rpt"
'   If opt1.Value = True Then
'      cry_rep1.Formulas(4) = ("opt=0")
'   Else
'      cry_rep1.Formulas(4) = ("opt=1")
'   End If
   Dim qry1 As String
   qry1 = ""
   If opt_emptypeall.Value = True Then
      qry1 = ""
   ElseIf opt_emptype_active.Value = True Then
      If qry1 <> "" Then
         qry1 = qry1 + " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
      Else
         qry1 = " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
      End If
   ElseIf opt_emptype_resigned.Value = True Then
      If qry1 <> "" Then
         qry1 = qry1 + " and {emp_mas.emp_status} = 'R'"
      Else
         qry1 = " and {emp_mas.emp_status} = 'R'"
      End If
   End If
   
''   If opt_sp.Value = True Then
''      pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 0"
''   ElseIf opt_st.Value = True Then
''      pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1"
''   ElseIf opt_wp.Value = True Then
''      pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2"
''   ElseIf opt_wt.Value = True Then
''      pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3"
''   End If
   
   If opt_sp.Value = True Then
      pst_qry = "{emp_salary.s_year} =  " & ryear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 0"
   ElseIf opt_st.Value = True Then
      pst_qry = "{emp_salary.s_year} =  " & ryear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1"
   ElseIf opt_wp.Value = True Then
      pst_qry = "{emp_salary.s_year} =  " & ryear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2"
   ElseIf opt_wt.Value = True Then
      pst_qry = "{emp_salary.s_year} =  " & ryear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3"
   End If
   
   
   
   cry_rep1.ReplaceSelectionFormula pst_qry & qry1
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub





