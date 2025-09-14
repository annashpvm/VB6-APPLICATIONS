VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_deduction_statement 
   Caption         =   "DEDUCTION REPORTS"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Select "
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      Begin VB.OptionButton opt_ho 
         Caption         =   "HO"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt_mill 
         Caption         =   "MILL"
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
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt_ho_mill 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3315
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   9390
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   2880
         TabIndex        =   12
         Top             =   1920
         Width           =   3375
         Begin VB.CommandButton cmd_exit 
            Caption         =   "&Exit"
            Height          =   990
            Left            =   2160
            Picture         =   "frm_deduction_statement.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Width           =   1110
         End
         Begin VB.CommandButton cmd_refresh 
            Caption         =   "&Refresh"
            Height          =   990
            Left            =   1080
            Picture         =   "frm_deduction_statement.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   1110
         End
         Begin VB.CommandButton cmd_print 
            Caption         =   "&Print"
            Height          =   990
            Left            =   0
            Picture         =   "frm_deduction_statement.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   6495
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
            Left            =   4800
            TabIndex        =   9
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
            Left            =   960
            TabIndex        =   8
            Top             =   270
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
            Left            =   120
            TabIndex        =   11
            Top             =   315
            Width           =   810
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
            Left            =   3720
            TabIndex        =   10
            Top             =   300
            Width           =   885
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
         Height          =   495
         Left            =   8280
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   375
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
            TabIndex        =   6
            Top             =   285
            Width           =   1770
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
            TabIndex        =   5
            Top             =   285
            Width           =   1770
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
            Height          =   450
            Left            =   4680
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   7695
      Begin VB.ComboBox cmb_deduction 
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
         Left            =   2160
         TabIndex        =   23
         Top             =   1080
         Width           =   5085
      End
      Begin VB.Frame Frame7 
         Height          =   615
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   4215
         Begin VB.OptionButton opt_selective 
            Caption         =   "SELECTIVE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   22
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton opt_all_ded 
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
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton opt_rep1 
         Caption         =   "Monthly deduction Details"
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
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.Label ded_label 
         Caption         =   "DEDUCTION"
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
         Height          =   315
         Left            =   480
         TabIndex        =   24
         Top             =   1080
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   -120
      Top             =   2865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_deduction_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
   Unload Me
End Sub

Private Sub cmd_Exit_Click()
        Unload Me
End Sub

Private Sub cmd_print_Click()
   If opt_staff.Value = True Then disname = "STAFF DEDUCTION STATEMENT FOR THE MONTH OF "
   If opt_worker.Value = True Then disname = "WORKER STATEMENT FOR THE MONTH OF "
   If opt_all.Value = True Then disname = "DEDUCTION STATEMENT FOR THE MONTH OF "
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   Dim ds As String
   If opt_ho.Value = True Then
      ds = " and {emp_mas.emp_classification} = 'A'"
   ElseIf opt_mill.Value = True Then
      ds = " and {emp_mas.emp_classification} = 'B'"
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
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\deduction_statement.rpt"
''   If opt_staff.Value = True Then
''      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S' " & ds & "")
''   ElseIf opt_worker.Value = True Then
''      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' " & ds & " ")
''   Else
''      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                              "and {emp_salary.s_company} = " & company_code & "" & ds & "")
''   End If
   
   If opt_all_ded.Value = True Then
      cry_rep1.ReplaceSelectionFormula ("{monthly_deduction.e_ded_year} = " & Val(cmb_year.Text) & " and {monthly_deduction.e_ded_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                              "and {monthly_deduction.e_company}= " & company_code & "")
   Else
       If cmb_deduction.ListIndex = -1 Then
          MsgBox ("Select Deduction Name...")
          Exit Sub
        End If
      cry_rep1.ReplaceSelectionFormula ("{monthly_deduction.e_ded_year} = " & Val(cmb_year.Text) & " and {monthly_deduction.e_ded_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                              "and {monthly_deduction.e_company}= " & company_code & " and {monthly_deduction.e_ded_code} = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & " ")
       
   End If
                                              

pst_qry = "{monthly_deduction.e_ded_year} = " & Val(cmb_year.Text) & " and {monthly_deduction.e_ded_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                              "and {monthly_deduction.e_company} = " & company_code & ""


   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1

End Sub

Private Sub Form_Load()
       Dim payrs As New ADODB.Recordset
    sql = "Select * from pdedu_mas"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_deduction.AddItem payrs(1)
        cmb_deduction.ItemData(cmb_deduction.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
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
