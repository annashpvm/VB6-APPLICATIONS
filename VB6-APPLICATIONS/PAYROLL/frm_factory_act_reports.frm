VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_factory_act_reports 
   Caption         =   "FACTORY ACT REPORTS"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15750
   WindowState     =   2  'Maximized
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
      Height          =   8475
      Left            =   555
      TabIndex        =   5
      Top             =   0
      Width           =   9240
      Begin VB.Frame Frame2 
         Caption         =   "SELECT "
         Height          =   4305
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   8760
         Begin VB.OptionButton opt_form15_leavewithwages 
            Caption         =   "Form-15 Leave with Wages"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   17
            Top             =   1800
            Width           =   4455
         End
         Begin VB.OptionButton opt_mustor_roll 
            Caption         =   "Muster Roll"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   16
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton opt_wages_fact_act 
            Caption         =   "Wages StateMent (Factory Act)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   3615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3480
         TabIndex        =   11
         Top             =   7080
         Width           =   1695
         Begin VB.CommandButton cmd_print 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "frm_factory_act_reports.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "frm_factory_act_reports.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "SALARY FOR THE MONTH OF "
         Height          =   1215
         Left            =   480
         TabIndex        =   6
         Top             =   5520
         Width           =   8415
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
            TabIndex        =   8
            Top             =   480
            Width           =   2775
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
            Left            =   5280
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4200
            TabIndex        =   9
            Top             =   480
            Width           =   885
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   975
      TabIndex        =   0
      Top             =   8880
      Visible         =   0   'False
      Width           =   4935
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   115802113
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   115802113
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_factory_act_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_print_Click()
   Dim mcode As Integer
   mcode = 1
    disname = "SALARY STATEMENT FOR THE MONTH OF "
          
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   cry_rep1.Formulas(2) = ("millname= '" & compname & "'")

  If opt_wages_fact_act.Value = True Then
     cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
     cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
     cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_fact_act.rpt"
     cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
  ElseIf opt_mustor_roll.Value = True Then

           ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 "
           cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
           cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
           
           
           cry_rep1.Formulas(2) = ("report_month = '" & cmb_month.Text & "'")
           cry_rep1.Formulas(3) = ("report_year = '" & cmb_year.Text & "'")
   
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\muster_roll.rpt"
           
           cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " ")
  ElseIf opt_form15_leavewithwages.Value = True Then
     cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
     cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
     cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\wages_statement_fact_act_form15.rpt"
     cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
    
  End If
  cry_rep1.WindowState = crptMaximized
  cry_rep1.Connect = gst_repconnect
  cry_rep1.Action = 1
End Sub

Private Sub EXIT_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    dt_doj_from = Now
    
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
    
        With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    

End Sub

Private Sub cmb_month_Click()
find_dates
End Sub


Private Sub cmb_year_Click()
find_dates
End Sub

Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
    
    Dim mdays, diff As Integer
    Dim d1 As Date
    If cmb_year.Text = "" Then Exit Sub
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

