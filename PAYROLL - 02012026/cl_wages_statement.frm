VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form cl_wages_statement 
   BackColor       =   &H00C0E0FF&
   Caption         =   "BONUS STATEMENT PRINTING"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   9240
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   39359
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
         TabIndex        =   23
         Top             =   240
         Width           =   1095
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
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   840
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4680
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
      Begin VB.CommandButton PRINT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "PRINT"
         Height          =   705
         Left            =   120
         Picture         =   "cl_wages_statement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   720
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "EXIT"
         Height          =   705
         Left            =   840
         Picture         =   "cl_wages_statement.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CL WAGES STATEMENT"
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
      Height          =   6075
      Left            =   660
      TabIndex        =   0
      Top             =   360
      Width           =   10575
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
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   9015
         Begin VB.OptionButton opt_emptypeall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_emptype_active 
            Caption         =   "ACTIVE"
            Height          =   375
            Left            =   3480
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptype_resigned 
            Caption         =   "RESIGNED"
            Height          =   375
            Left            =   6840
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3735
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   9615
         Begin VB.Frame Frame5 
            Height          =   855
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   9255
            Begin VB.OptionButton opt_bank 
               Caption         =   "Bank Statement"
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
               Left            =   6720
               TabIndex        =   28
               Top             =   240
               Width           =   2055
            End
            Begin VB.OptionButton optWages_daywages 
               Caption         =   "Day Wages with Salary"
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
               Left            =   1800
               TabIndex        =   27
               Top             =   240
               Width           =   2535
            End
            Begin VB.OptionButton opt_attn 
               Caption         =   "Attendance"
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
               Left            =   4680
               TabIndex        =   26
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optWages 
               Caption         =   "Day Wages"
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
               TabIndex        =   25
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
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
            Left            =   3360
            TabIndex        =   17
            Top             =   2880
            Width           =   1335
         End
         Begin VB.OptionButton opt3 
            Caption         =   "Non PF Members"
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
            TabIndex        =   16
            Top             =   2160
            Width           =   5775
         End
         Begin VB.OptionButton opt2 
            Caption         =   "PF Member"
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
            TabIndex        =   15
            Top             =   1800
            Width           =   4215
         End
         Begin VB.OptionButton opt1 
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
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   480
            TabIndex        =   10
            Top             =   1440
            Value           =   -1  'True
            Width           =   3615
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
            Left            =   2160
            TabIndex        =   18
            Top             =   2880
            Width           =   885
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
   End
End
Attribute VB_Name = "cl_wages_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo2_Change()

End Sub

Private Sub cmb_year_Click()
    If cmb_year.Text = "" Then Exit Sub
    Dim d1 As Date
    mdays = 31
    
    end_date = DateValue("12/31/" + cmb_year.Text)
    st_date = DateValue("01/07/" + cmb_year.Text)
    


    If end_date.Value > Now Then
       end_date.Value = Now + 1
    End If
End Sub

Private Sub exit_Click()
   Unload Me
End Sub

 Private Sub Form_Load()
    st_date.Value = Now
    end_date.Value = Now
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
End Sub

Private Sub print_Click()
   

    
    
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   If optWages.Value = True Then
      disname = "CASUAL LEAVE WAGES STATEMENT FOR THE PERIOD " + cmb_year.Text
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\cl_wages_statement.rpt"
   ElseIf optWages_daywages.Value = True Then
      disname = "CASUAL LEAVE SALARY DAYS STATEMENT FOR THE PERIOD " + cmb_year.Text
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\cl_wages_statement_with_salary.rpt"
   ElseIf opt_attn.Value = True Then
      disname = "CASUAL LEAVE SALARY DAYS STATEMENT FOR THE PERIOD " + cmb_year.Text
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\cl_wages_Attn_statement.rpt"

   Else
      disname = "CASUAL LEAVE SALARY DAYS STATEMENT FOR THE PERIOD " + cmb_year.Text
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\cl_wages_statement_Bank.rpt"
   End If


   cry_rep1.Formulas(0) = ("report_year = '" & fyear & "'")
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
   
   

  pst_qry = "{emp_salary.s_finyear} =  " & finyear & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.EMP_DOJ} <   date(" & Format(st_date.Value, "yyyy,mm,dd") & ") "

   
   If opt1.Value = True Then
          pst_qry = "{emp_salary.s_company}= " & company_code & " and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_mas.EMP_DOJ} <   date(" & Format(st_date.Value, "yyyy,mm,dd") & ") "
   Else
      If opt2.Value = True Then
''          pst_qry = " {emp_mas.emp_pfeligible} = 'Y' and {emp_salary.s_company}= " & company_code & " and (({emp_salary.s_year} = " & Val(cmb_year_from.Text) & "  and {emp_salary.s_month} >= " & cmb_month_from.ItemData(cmb_month_from.ListIndex) & " ) or ({emp_salary.s_year} =  " & Val(cmb_year_to.Text) & " and {emp_salary.s_month} <= " & cmb_month_to.ItemData(cmb_month_to.ListIndex) & ") )"
          pst_qry = "{emp_mas.emp_pfeligible} = 'Y' and {emp_salary.s_company}= " & company_code & " and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_mas.EMP_DOJ} <   date(" & Format(st_date.Value, "yyyy,mm,dd") & ") "
      Else
  ''        pst_qry = " {emp_mas.emp_pfeligible} = 'N' and {emp_salary.s_company}= " & company_code & " and (({emp_salary.s_year} = " & Val(cmb_year_from.Text) & "  and {emp_salary.s_month} >= " & cmb_month_from.ItemData(cmb_month_from.ListIndex) & " ) or ({emp_salary.s_year} =  " & Val(cmb_year_to.Text) & " and {emp_salary.s_month} <= " & cmb_month_to.ItemData(cmb_month_to.ListIndex) & ") )"
          pst_qry = "{emp_mas.emp_pfeligible} = 'N' and {emp_salary.s_company}= " & company_code & " and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_mas.EMP_DOJ} <   date(" & Format(st_date.Value, "yyyy,mm,dd") & ") "
          
      End If
   End If
   


   cry_rep1.ReplaceSelectionFormula pst_qry & qry1
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub




