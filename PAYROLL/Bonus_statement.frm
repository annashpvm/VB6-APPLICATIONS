VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Bonus_statement 
   BackColor       =   &H00C0E0FF&
   Caption         =   "BONUS STATEMENT PRINTING"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4680
      TabIndex        =   8
      Top             =   7920
      Width           =   1695
      Begin VB.CommandButton PRINT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "PRINT"
         Height          =   705
         Left            =   120
         Picture         =   "Bonus_statement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   720
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "EXIT"
         Height          =   705
         Left            =   840
         Picture         =   "Bonus_statement.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BONUS STATEMENT"
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
      Height          =   7395
      Left            =   660
      TabIndex        =   0
      Top             =   360
      Width           =   11175
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
         Height          =   855
         Left            =   600
         TabIndex        =   13
         Top             =   4920
         Visible         =   0   'False
         Width           =   9015
         Begin VB.OptionButton opt_emptypeall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_emptype_active 
            Caption         =   "ACTIVE"
            Height          =   375
            Left            =   3480
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptype_resigned 
            Caption         =   "RESIGNED"
            Height          =   375
            Left            =   6840
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4455
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   9855
         Begin VB.OptionButton opt5 
            Caption         =   "Bonus Statement  - All (Excluding Basic && DA)"
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
            Left            =   600
            TabIndex        =   32
            Top             =   1800
            Value           =   -1  'True
            Width           =   5415
         End
         Begin VB.OptionButton opt4 
            Caption         =   "Bonus Statement  - All (Basic && DA)"
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
            Left            =   600
            TabIndex        =   31
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Frame Frame6 
            Height          =   1215
            Left            =   6000
            TabIndex        =   28
            Top             =   240
            Width           =   3015
            Begin VB.OptionButton opt_Abstract 
               Caption         =   "Abstract"
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
               TabIndex        =   30
               Top             =   600
               Width           =   1935
            End
            Begin VB.OptionButton opt_detail 
               Caption         =   "Detailed"
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
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.OptionButton opt3 
            Caption         =   "Bonus Statement  - NON PF"
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
            Left            =   600
            TabIndex        =   27
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton opt2 
            Caption         =   "Bonus Statement  - PF"
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
            Left            =   600
            TabIndex        =   26
            Top             =   600
            Width           =   3375
         End
         Begin VB.Frame Frame5 
            Height          =   2055
            Left            =   960
            TabIndex        =   17
            Top             =   2280
            Width           =   8055
            Begin VB.ComboBox cmb_month_to 
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
               TabIndex        =   23
               Top             =   1440
               Width           =   2655
            End
            Begin VB.ComboBox cmb_year_to 
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
               Left            =   6360
               TabIndex        =   22
               Top             =   1440
               Width           =   1335
            End
            Begin VB.ComboBox cmb_month_from 
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
               TabIndex        =   19
               Top             =   600
               Width           =   2655
            End
            Begin VB.ComboBox cmb_year_from 
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
               Left            =   6360
               TabIndex        =   18
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "YEAR"
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
               Height          =   285
               Left            =   5205
               TabIndex        =   25
               Top             =   1440
               Width           =   885
            End
            Begin VB.Label Label4 
               Caption         =   "Month To"
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
               Height          =   330
               Left            =   480
               TabIndex        =   24
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "YEAR"
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
               Height          =   285
               Left            =   5160
               TabIndex        =   21
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label2 
               Caption         =   "Month From"
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
               Height          =   330
               Left            =   480
               TabIndex        =   20
               Top             =   600
               Width           =   1335
            End
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Bonus Statement  - All"
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
            Left            =   600
            TabIndex        =   12
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.TextBox txt_bonus 
         Height          =   495
         Left            =   5640
         TabIndex        =   7
         Top             =   6480
         Width           =   1335
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
         Height          =   465
         Left            =   480
         TabIndex        =   1
         Top             =   3720
         Visible         =   0   'False
         Width           =   705
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
            TabIndex        =   5
            Top             =   480
            Width           =   1905
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
            TabIndex        =   4
            Top             =   480
            Width           =   1740
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
            TabIndex        =   2
            Top             =   360
            Width           =   1770
         End
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
            TabIndex        =   3
            Top             =   480
            Width           =   1950
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Bonus Percentage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   6360
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Bonus_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo2_Change()

End Sub

Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    With cmb_month_from
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
''''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''
''        .Text = "2015"
''    End With
    With cmb_year_from
        .AddItem Left(fyear, 4) - 1
        .AddItem Mid(fyear, 6, 4) - 1
    End With
    
    
    With cmb_month_to
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
''''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''
''        .Text = "2015"
''    End With
    With cmb_year_to
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    
    cmb_month_from.ListIndex = 9
    cmb_month_to.ListIndex = 8
    opt_sp.Value = True
''    With cmb_year
''''        .AddItem "2011-2012"
''        .AddItem "2012-2013"
''        .AddItem "2013-2014"

''    End With
''''    cmb_year.Text = "2007-2008"
End Sub

Private Sub print_Click()
   If opt_sp.Value = True Then disname = "BONUS STATEMENT FOR THE PERIOD " + cmb_month_from.Text + "-" + cmb_year_from.Text + " TO " + cmb_month_to.Text + "-" + cmb_year_to.Text
   If opt_st.Value = True Then disname = "BONUS STATEMENT FOR THE PERIOD " + cmb_month_from.Text + "-" + cmb_year_from.Text + " TO " + cmb_month_to.Text + "-" + cmb_year_to.Text
   If opt_wp.Value = True Then disname = "BONUS STATEMENT FOR THE PERIOD " + cmb_month_from.Text + "-" + cmb_year_from.Text + " TO " + cmb_month_to.Text + "-" + cmb_year_to.Text
   If opt_wt.Value = True Then disname = "BONUS STATEMENT FOR THE PERIOD " + cmb_month_from.Text + "-" + cmb_year_from.Text + " TO " + cmb_month_to.Text + "-" + cmb_year_to.Text
''   syear = Val(Left(cmb_year.Text, 4))
''   eyear = Val(Mid(cmb_year.Text, 6, 4))
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect

   If opt_detail.Value = True Then
      If opt4.Value = True Then
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_BASIC_DA.rpt"
      ElseIf opt5.Value = True Then
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_BASIC_DA_Excluding.rpt"
      Else
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement.rpt"
      End If
   Else
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\bonus_statement_bank.rpt"
   End If

''   cry_rep1.ReportFileName = "\\annadurai\d\payroll\bonus_statement.rpt"
   cry_rep1.Formulas(0) = ("report_year = '" & fyear & "'")
   cry_rep1.Formulas(1) = ("bonus_per = " & Val(txt_bonus.Text) & "")
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
   If opt1.Value = True Then
      cry_rep1.Formulas(4) = ("opt=0")
   Else
      cry_rep1.Formulas(4) = ("opt=1")
   End If
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
   
   

   pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code

   
   If opt1.Value = True Or opt4.Value = True Or opt5.Value = True Then
      pst_qry = "{emp_salary.s_company}= " & company_code & " and (({emp_salary.s_year} = " & Val(cmb_year_from.Text) & "  and {emp_salary.s_month} >= " & cmb_month_from.ItemData(cmb_month_from.ListIndex) & " ) or ({emp_salary.s_year} =  " & Val(cmb_year_to.Text) & " and {emp_salary.s_month} <= " & cmb_month_to.ItemData(cmb_month_to.ListIndex) & ") )"
   Else
      If opt2.Value = True Then
          pst_qry = " {emp_mas.emp_pfeligible} = 'Y' and {emp_salary.s_company}= " & company_code & " and (({emp_salary.s_year} = " & Val(cmb_year_from.Text) & "  and {emp_salary.s_month} >= " & cmb_month_from.ItemData(cmb_month_from.ListIndex) & " ) or ({emp_salary.s_year} =  " & Val(cmb_year_to.Text) & " and {emp_salary.s_month} <= " & cmb_month_to.ItemData(cmb_month_to.ListIndex) & ") )"
      Else
          pst_qry = " {emp_mas.emp_pfeligible} = 'N' and {emp_salary.s_company}= " & company_code & " and (({emp_salary.s_year} = " & Val(cmb_year_from.Text) & "  and {emp_salary.s_month} >= " & cmb_month_from.ItemData(cmb_month_from.ListIndex) & " ) or ({emp_salary.s_year} =  " & Val(cmb_year_to.Text) & " and {emp_salary.s_month} <= " & cmb_month_to.ItemData(cmb_month_to.ListIndex) & ") )"
      End If
   End If
   


   cry_rep1.ReplaceSelectionFormula pst_qry & qry1
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub




