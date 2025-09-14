VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form salary_statement_departmentwise 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "SALARY STATEMENT -DEPARTMENTWISE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   7755
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   9240
      Begin VB.Frame Frame8 
         Caption         =   "SELECT REPORT"
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
         TabIndex        =   30
         Top             =   480
         Width           =   7815
         Begin VB.OptionButton opt_actual 
            Caption         =   "ACTUAL SALARY"
            Height          =   375
            Left            =   5280
            TabIndex        =   33
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton opt_abstract 
            Caption         =   "ABSTRACT"
            Height          =   375
            Left            =   3120
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton opt_statement 
            Caption         =   "STATEMENT"
            Height          =   375
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Height          =   855
         Left            =   360
         TabIndex        =   25
         Top             =   5520
         Width           =   7815
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
            TabIndex        =   27
            Top             =   240
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
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4200
            TabIndex        =   28
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "MILLWISE"
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
         Height          =   1185
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   7785
         Begin VB.Frame frame_mill 
            Caption         =   "SELECT MILL"
            Enabled         =   0   'False
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
            Height          =   135
            Left            =   1920
            TabIndex        =   19
            Top             =   960
            Width           =   615
            Begin VB.OptionButton opt_solvent 
               Caption         =   "SOLVENT"
               Height          =   375
               Left            =   4080
               TabIndex        =   24
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton opt_cogen 
               Caption         =   "COGEN"
               Height          =   375
               Left            =   3000
               TabIndex        =   23
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_vjpm 
               Caption         =   "VJPM"
               Height          =   375
               Left            =   2040
               TabIndex        =   22
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_slpb 
               Caption         =   "SLPB"
               Height          =   375
               Left            =   1200
               TabIndex        =   21
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt_dpm 
               Caption         =   "DPM"
               Height          =   375
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.OptionButton opt_millselective 
            Caption         =   "SELECTIVE"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt_millall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3720
         TabIndex        =   13
         Top             =   6600
         Width           =   1695
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "salary_statement_departmentwise.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "salary_statement_departmentwise.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT PERMENANT / TRANIEE"
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
         Height          =   825
         Left            =   360
         TabIndex        =   9
         Top             =   4680
         Width           =   7800
         Begin VB.OptionButton opt_trainee 
            Caption         =   "TRANIEES"
            Height          =   285
            Left            =   4680
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_all_per_trainee 
            Caption         =   "ALL"
            Height          =   285
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_permenent 
            Caption         =   "PERMENANT"
            Height          =   285
            Left            =   2400
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "LOCATION"
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
         Height          =   705
         Left            =   360
         TabIndex        =   5
         Top             =   3840
         Width           =   7800
         Begin VB.OptionButton opt_all_location 
            Caption         =   "ALL"
            Height          =   405
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_vpt 
            Caption         =   "MILLS"
            Height          =   405
            Left            =   2400
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_cbe 
            Caption         =   "COIMBATORE"
            Height          =   405
            Left            =   4560
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "STAFF / WORKER"
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
         Left            =   360
         TabIndex        =   1
         Top             =   2760
         Width           =   7815
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            Height          =   375
            Left            =   4560
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton opt_sw 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            Height          =   375
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   1440
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "salary_statement_departmentwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'' dt.Value = Now
opt_statement.Value = True
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

Private Sub opt_millall_Click()
    frame_mill.Enabled = False
End Sub

Private Sub opt_millselective_Click()
    frame_mill.Enabled = True
End Sub

Private Sub PROCESS_Click()
   Dim wp, qry1, qry2, qry3 As String
   Dim date1 As String
   MousePointer = vbDefault
   qry1 = ""
   qry2 = ""
   qry3 = ""
'''   date1 = Format(DateAdd("m", -54, dt.Value), "yyyy,mm,dd")
      wp = ""
''   If opt_emptypeall.Value = True Then
''   ElseIf opt_emptype_active.Value = True Then
''      If qry1 <> "" Then
''         qry1 = qry1 + " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
''      Else
''         qry1 = " ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
''      End If
''   ElseIf opt_emptype_resigned.Value = True Then
''      If qry1 <> "" Then
''         qry1 = qry1 + " and {emp_mas.emp_status} = 'R'"
''      Else
''         qry1 = " {emp_mas.emp_status} = 'R'"
''      End If
''   End If
   
   If opt_sw.Value = True Then
   ElseIf opt_staff.Value = True Then
      If qry1 <> "" Then
          qry1 = qry1 + " and {vew_payroll_emp_mas.emp_cat} <> 'W'"
      Else
          qry1 = "{vew_payroll_emp_mas.emp_cat} <> 'W'"
      End If
   ElseIf opt_worker.Value = True Then
      If qry1 <> "" Then
         qry1 = qry1 + "and {vew_payroll_emp_mas.emp_cat} = 'W'"
      Else
         qry1 = "{vew_payroll_emp_mas.emp_cat} = 'W'"
      End If
   End If
        
   If opt_all_per_trainee.Value = True Then
   ElseIf opt_permenent.Value = True Then
       If qry1 <> "" Then
          qry1 = qry1 + "and ({vew_payroll_emp_mas.emp_type} = 0  or {vew_payroll_emp_mas.emp_type} = 2) "
       Else
          qry1 = "({vew_payroll_emp_mas.emp_type} = 0  or {vew_payroll_emp_mas.emp_type} = 2) "
       End If
   ElseIf opt_trainee.Value = True Then
       If qry1 <> "" Then
          qry1 = qry1 + "and ({vew_payroll_emp_mas.emp_type} = 1  or {vew_payroll_emp_mas.emp_type} = 3) "
       Else
          qry1 = "({vew_payroll_emp_mas.emp_type} = 1  or {vew_payroll_emp_mas.emp_type} = 3) "
       End If
   End If
   
   If opt_millselective.Value = True Then
       If opt_dpm.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {vew_payroll_emp_mas.emp_company} = 1"
            Else
               qry1 = "{vew_payroll_emp_mas.emp_company} = 1"
            End If
       End If
       If opt_slpb.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {vew_payroll_emp_mas.emp_company} = 2"
            Else
               qry1 = "{vew_payroll_emp_mas.emp_company} = 2"
            End If
       End If
       If opt_vjpm.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {vew_payroll_emp_mas.emp_company} = 3"
            Else
               qry1 = "{vew_payroll_emp_mas.emp_company} = 3"
            End If
       End If
       If opt_cogen.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {vew_payroll_emp_mas.emp_company} = 5"
            Else
               qry1 = "{vew_payroll_emp_mas.emp_company} = 5"
            End If
       End If
       If opt_solvent.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {vew_payroll_emp_mas.emp_company} = 8"
            Else
               qry1 = "{vew_payroll_emp_mas.emp_company} = 8"
            End If
       End If
       
   End If
   
   
   pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_payroll_emp_mas]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
              & " drop view [dbo].[vew_payroll_emp_mas] "
    paydb.Execute (pst_qry)
    pst_qry = " create view vew_payroll_emp_mas as " _
              & " select EMP_COMPANY,EMP_CODE,EMP_CAT,EMP_TYPE,EMP_NAME,EMP_DEPT,EMP_DESIGN,EMP_STATUS,EMP_CLASSIFICATION,EMP_WORKPLACE,EMP_SALARY_SLOT from emp_voupay_mast where emp_status='A'" _
              & " Union All " _
              & " select EMP_COMPANY,EMP_CODE,EMP_CAT,EMP_TYPE,EMP_NAME,EMP_DEPT,EMP_DESIGN,EMP_STATUS,EMP_CLASSIFICATION,EMP_WORKPLACE,EMP_SALARY_SLOT from emp_mas where emp_status='A'"
    paydb.Execute (pst_qry)
    
    
    
    If opt_millall.Value = True And opt_all_location.Value = True And opt_all_per_trainee.Value = True And opt_sw.Value = True Then
       qry1 = "{emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and   {emp_salary.s_year} = " & Val(cmb_year.Text) & "  and {emp_salary.s_netpay} > 0  "
    Else
        qry1 = qry1 + " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and  {emp_salary.s_year}=" & Val(cmb_year.Text) & " and  {emp_salary.s_year} = " & Val(cmb_year.Text) & "  and {emp_salary.s_netpay} > 0 "
    End If
   If opt_statement.Value = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_departmentwise.rpt"
   ElseIf opt_abstract.Value = True Then
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_abstract_departmentwise.rpt"
   Else
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_departmentwise_actual_salary.rpt"
   End If
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   
   
''   cry_rep1.Formulas(5) = ("Ason='" & Format(dt.Value, "yyyy/MM/dd") & "'")
   If opt_dpm.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'I'")
   ElseIf opt_slpb.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'II'")
   ElseIf opt_vjpm.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD'")
   ElseIf opt_cogen.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'IIC'")
   ElseIf opt_solvent.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'OIL PLANT'")
   End If
   
      cry_rep1.Formulas(2) = ("millname= '" & compname & "'")
      
   cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   
   If opt_staff.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF'")
   ElseIf opt_worker.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'WORKER'")
   ElseIf opt_sw.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF / WORKER'")
   End If
   
''   If opt_emptypeall.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT + RESIGNED EMPLOYEES'")
''   ElseIf opt_emptype_active.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT EMPLOYEES'")
''   ElseIf opt_emptype_resigned.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'RESIGNED EMPLOYEES'")
''
''   End If
   
   cry_rep1.DiscardSavedData = True
''   qry1 = "{vew_payroll_emp_mas.emp_workplace} = 'CBE' and ({vew_payroll_emp_mas.emp_type} = 1  or {vew_payroll_emp_mas.emp_type} = 3) and {vew_payroll_emp_mas.emp_company} = 1 and {emp_salary.s_month} = 12  and  {emp_salary.s_year}=2015 and  {emp_salary.s_year} = 2015  and {emp_salary.s_netpay} > 0 "
''   qry1 = "{vew_payroll_emp_mas.emp_workplace} = 'CBE'  and {vew_payroll_emp_mas.emp_type} = " 1 " and {vew_payroll_emp_mas.emp_company} = 1 and {emp_salary.s_month} = 12  and  {emp_salary.s_year}=2015 and  {emp_salary.s_year} = 2015  and {emp_salary.s_netpay} > 0  "
''   qry1 = "{vew_payroll_emp_mas.emp_type} =  1"
   cry_rep1.ReplaceSelectionFormula (qry1)
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
   Exit Sub
 End Sub




