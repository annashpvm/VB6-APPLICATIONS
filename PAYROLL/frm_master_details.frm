VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_master_details 
   Caption         =   "MASTER DETAILS"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   3735
      Left            =   11520
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   4575
      Begin VB.OptionButton opt_salary 
         Caption         =   "Salary Details - CTC"
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
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RETAINER"
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
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MASTER DETAILS"
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
      Height          =   6375
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   9390
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   600
         TabIndex        =   21
         Top             =   840
         Width           =   8175
         Begin VB.OptionButton opt_address 
            Caption         =   "Employee Details"
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
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   3135
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
         Height          =   435
         Left            =   600
         TabIndex        =   17
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
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
            Left            =   3960
            TabIndex        =   20
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
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   1215
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
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   3000
         TabIndex        =   13
         Top             =   5280
         Width           =   3015
         Begin VB.CommandButton print 
            Caption         =   "&Print"
            Height          =   870
            Left            =   0
            Picture         =   "frm_master_details.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Refresh 
            Caption         =   "&Refresh"
            Height          =   870
            Left            =   960
            Picture         =   "frm_master_details.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton Exit 
            Caption         =   "&Exit"
            Height          =   870
            Left            =   1920
            Picture         =   "frm_master_details.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   240
            Width           =   1335
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
            Left            =   4980
            TabIndex        =   12
            Top             =   315
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
            Left            =   360
            TabIndex        =   11
            Top             =   345
            Width           =   1050
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122486785
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
         Format          =   122486785
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker dt_joining 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122486785
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker dt_resigned 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122486785
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   375
      Top             =   3990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_master_details"
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

Private Sub cmb_month_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     find_dates
  End If
End Sub
Private Sub cmb_year_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     find_dates
  End If
End Sub

Public Sub find_dates()
    Dim mdays, diff As Integer
    Dim d1 As Date
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

Private Sub print_Click()
   Dim ds, pst_qry As String
''   If optchk = 2 Then
''      ds = " and {emp_mas.emp_classification} = 'A'"
''   ElseIf optchk = 1 Then
''      ds = " and {emp_mas.emp_classification} = 'B'"
''   Else
''    ds = ""
''   End If
   cry_rep1.Formulas(5) = ""
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_emp_mas_company]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                 & " drop view [dbo].[vew_emp_mas_company] "
''   gcn_servall.Execute (pst_qry)
''
''   pst_qry = "create view vew_emp_mas_company as " _
''              & " select  * from emp_mas where emp_company in [1,2,3,8]"
''
''   gcn_servall.Execute (pst_qry)
   cry_rep1.Formulas(0) = ""
   cry_rep1.Formulas(1) = ""
   cry_rep1.Formulas(4) = ""
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   If opt_staff.Value = True Then
      cry_rep1.Formulas(3) = ("sthead = 'STAFF'")
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\master_data_salary_staff.rpt"
   Else
      cry_rep1.Formulas(3) = ("sthead = 'WORKER'")
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\master_data_salary.rpt"
   End If
    cry_rep1.Formulas(3) = ("sthead = ''")
     cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\master_data_salary.rpt"
   If opt_salary.Value = True Then
      If opt_staff.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_cat} = 'S' and {emp_mas.emp_status} = 'A' and ({emp_mas.emp_company} = 1 or {emp_mas.emp_company} = 2 or {emp_mas.emp_company} = 3  or {emp_mas.emp_company} = 5)  ")
      Else
         If cmb_month.ListIndex = -1 Then Exit Sub
         If cmb_year.Text = "" Then Exit Sub
         cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
         cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
         cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W' and {emp_salary.s_salarydays} > 0  ")
   
      End If
   Else
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\master_data_address.rpt"
      If opt_staff.Value = True Then
         cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_cat} = 'S' and {emp_mas.emp_status} = 'A' and {emp_mas.emp_company} = " & company_code & "")
      Else
         cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_cat} = 'W'")
         cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_cat} = 'W' and {emp_mas.emp_status} = 'A' and {emp_mas.emp_company} = " & company_code & "")
          cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_status} = 'A' and {emp_mas.emp_company} = " & company_code & "")
      End If
   End If
  '' cry_rep1.ReplaceSelectionFormula pst_qry
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1

End Sub
