VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_rep_grosspay_annual 
   Caption         =   "GROSS PAY ANNUAL STATEMENT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   840
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "GROSS PAY STATEMENT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   9240
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   7215
         Begin VB.OptionButton opt_gp_eligible 
            Caption         =   "GROSS PAY - EARNED"
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
            Left            =   3960
            TabIndex        =   16
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton opt_gp_actual 
            Caption         =   "GROSS PAY - ACTUAL"
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
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT "
         Height          =   2145
         Left            =   840
         TabIndex        =   7
         Top             =   1200
         Width           =   7200
         Begin VB.OptionButton opt_sp 
            Caption         =   "PERMANENT STAFF"
            Height          =   405
            Left            =   480
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   3615
         End
         Begin VB.OptionButton opt_wt 
            Caption         =   "TRAINEE  WORKER"
            Height          =   300
            Left            =   480
            TabIndex        =   11
            Top             =   1560
            Width           =   4515
         End
         Begin VB.OptionButton opt_wp 
            Caption         =   "PERMANENT WORKER"
            Height          =   285
            Left            =   480
            TabIndex        =   10
            Top             =   1080
            Width           =   4215
         End
         Begin VB.OptionButton opt_st 
            Caption         =   "TRAINEE STAFF"
            Height          =   420
            Left            =   480
            TabIndex        =   9
            Top             =   600
            Width           =   4155
         End
         Begin VB.CheckBox chk_deptwise 
            Caption         =   "DEPARTMENTWISE"
            Height          =   255
            Left            =   4080
            TabIndex        =   8
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3600
         TabIndex        =   4
         Top             =   4440
         Width           =   1695
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "frm_rep_grosspay_annual.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "frm_rep_grosspay_annual.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   2160
         TabIndex        =   1
         Top             =   3480
         Width           =   4695
         Begin VB.TextBox txt_fincode 
            Height          =   285
            Left            =   2400
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   720
            Visible         =   0   'False
            Width           =   255
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
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_rep_grosspay_annual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmb_year_Click()
Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  mas_finyear where fin_year ='" & cmb_year.Text & "' ")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    txt_fincode.Text = payrs!fin_code
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  mas_finyear where fin_code > 11")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_year.AddItem payrs("fin_year")
        cmb_year.ItemData(cmb_year.NewIndex) = payrs("fin_code")
        payrs.MoveNext
    Wend

End Sub

Private Sub PROCESS_Click()
Dim resigned_date As Date
Dim fin As String
If Trim(cmb_year.Text) = "" Then
      MsgBox ("Select the Reporting Year")
      Exit Sub
   End If
fin = 20 & "" & Val(txt_fincode.Text) & ""
resigned_date = 1 & "/" & 4 & "/" & fin & ""
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   
   If opt_gp_actual.Value = True Then
       cry_rep1.Formulas(5) = ("gptype = 0")
   Else
       cry_rep1.Formulas(5) = ("gptype = 1")
   End If
   cry_rep1.Formulas(1) = ("report_year = '" & cmb_year.Text & "'")
''   cry_rep1.Formulas(4) = ("rmonth = '" & cmb_year.Text & "'")
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(3) = ("fin_year = " & Val(txt_fincode.Text) & "")
''   If chk_deptwise.Value = True Then
''        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_grosspay_annual.rpt"
''   Else
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_grosspay_annual.rpt"

''   End If
   If opt_sp.Value = True Then
''      pst_qry = "{emp_salary.s_finyear} =  " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_resigneddate) > " & Format(resigned_date, "yyyy/MM/dd") & " "
      pst_qry = "{emp_salary.s_finyear} =  " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and  {emp_mas.emp_type} = 0  "
'      Cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_finyear} = " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 0 ")
   ElseIf opt_st.Value = True Then

      pst_qry = "{emp_salary.s_finyear} =  " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1  "
'      Cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_finyear} = " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1")
   ElseIf opt_wp.Value = True Then

      pst_qry = "{emp_salary.s_finyear} =  " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2  "
'      Cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_finyear} = " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 2")
   ElseIf opt_wt.Value = True Then

      pst_qry = "{emp_salary.s_finyear} =  " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 "
'      Cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_finyear} = " & Val(txt_fincode.Text) & " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3")
   End If
   cry_rep1.ReplaceSelectionFormula (pst_qry)
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub
