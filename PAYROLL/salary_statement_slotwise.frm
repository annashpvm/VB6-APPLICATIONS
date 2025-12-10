VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salary_statement_slotwise 
   Caption         =   "SLOT WISE SALARY STATMENT"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   1815
      Left            =   8280
      TabIndex        =   35
      Top             =   4080
      Width           =   2175
      Begin VB.OptionButton opt_cat_retain 
         Caption         =   "Retainer"
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
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton opt_cat_perm 
         Caption         =   "Permanent"
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
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton opt_cat_all 
         Caption         =   "All"
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
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt_cat_trainee 
         Caption         =   "Trainee"
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Salary Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   8280
      TabIndex        =   28
      Top             =   480
      Width           =   2295
      Begin VB.OptionButton opt_all_salary_range 
         Caption         =   "All Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton opt_selective_range 
         Caption         =   "Selective Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txt_from 
         Height          =   495
         Left            =   600
         TabIndex        =   30
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt_to 
         Height          =   495
         Left            =   600
         TabIndex        =   29
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "FROM "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   2880
      TabIndex        =   12
      Top             =   7200
      Width           =   3015
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   750
         Left            =   1920
         Picture         =   "salary_statement_slotwise.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Refresh 
         Caption         =   "&Refresh"
         Height          =   750
         Left            =   1080
         Picture         =   "salary_statement_slotwise.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton print 
         Caption         =   "&Print"
         Height          =   750
         Left            =   120
         Picture         =   "salary_statement_slotwise.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   855
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
      Height          =   6675
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.Frame Frame8 
         Caption         =   "Bank Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   24
         Top             =   360
         Width           =   6495
         Begin VB.OptionButton Opt_waccno 
            Caption         =   "With Account no"
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
            Left            =   960
            TabIndex        =   26
            Top             =   360
            Width           =   2250
         End
         Begin VB.OptionButton opt_woutaccno 
            Caption         =   "Without Account No"
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
            Left            =   3840
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select Mill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   21
         Top             =   1440
         Width           =   6495
         Begin VB.OptionButton opt_selective_mill 
            Caption         =   "Current Mill"
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
            Left            =   3840
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   2130
         End
         Begin VB.OptionButton opt_all_mills 
            Caption         =   "ALL MILLS"
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
            Left            =   960
            TabIndex        =   22
            Top             =   360
            Width           =   1530
         End
      End
      Begin VB.Frame Frame7 
         Height          =   975
         Left            =   600
         TabIndex        =   16
         Top             =   5520
         Width           =   6495
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
            Left            =   1080
            TabIndex        =   18
            Top             =   360
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
            Left            =   4800
            TabIndex        =   17
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
            Left            =   3960
            TabIndex        =   20
            Top             =   360
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
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   9
         Top             =   4560
         Width           =   6495
         Begin VB.OptionButton opt_layoffdays 
            Caption         =   "Layoff  days"
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
            Left            =   4200
            TabIndex        =   27
            Top             =   360
            Width           =   1890
         End
         Begin VB.OptionButton opt_alldays 
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
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   1530
         End
         Begin VB.OptionButton opt_Presentdays 
            Caption         =   "Worked days"
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
            Left            =   1920
            TabIndex        =   10
            Top             =   360
            Width           =   2130
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "SALARY SLOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   5
         Top             =   2520
         Width           =   6495
         Begin VB.ComboBox cmb_slot 
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
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_selective_slot 
            Caption         =   "SELECTIVE SLOT"
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
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   2130
         End
         Begin VB.OptionButton opt_all_slots 
            Caption         =   "ALL SLOTS"
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
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1530
         End
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
         Height          =   975
         Left            =   600
         TabIndex        =   1
         Top             =   3600
         Width           =   6510
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
            TabIndex        =   4
            Top             =   360
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
            TabIndex        =   3
            Top             =   360
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
            TabIndex        =   2
            Top             =   390
            Width           =   1170
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   3225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "salary_statement_slotwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    cmb_slot.AddItem "SLOT1"
    cmb_slot.AddItem "SLOT2"
    cmb_slot.AddItem "SLOT3"
    cmb_slot.AddItem "SLOT4"
    cmb_slot.AddItem "SLOT5"
    cmb_slot.AddItem "SLOT6"
    opt_alldays.Value = True
    opt_staff.Value = True
    ''rep_p.Value = True
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

Private Sub opt_abstract_Click()

End Sub

Private Sub opt_empwise_Click()

End Sub

Private Sub print_Click()
   If opt_all_slots.Value = False And opt_selective_slot.Value = False Then
      MsgBox ("Select selective slot or All slot...")
      Exit Sub
   End If
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   If opt_selective_range.Value = True Then
      If Val(txt_from.Text) = 0 Or Val(txt_to.Text) = 0 Then
         MsgBox ("Enter Salary Range ....")
         Exit Sub
      End If
      If Val(txt_from.Text) > Val(txt_to.Text) Then
         MsgBox ("Error in Salary Range ....")
         Exit Sub
      End If
      
   End If

   If opt_staff.Value = True Then disname = "SLOT WISE STAFF SALARY STATEMENT"
   If opt_worker.Value = True Then disname = "SLOT WISE WORKER WAGES STATEMENT"
   If opt_All.Value = True Then disname = "SLOT WISE STAFF / WORKERS SALARY/WAGES STATEMENT"

   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   If opt_selective_mill.Value = True Then
      cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
   Else
      cry_rep1.Formulas(0) = ("millname= 'ALL MILLS'")
   End If
   If opt_layoffdays.Value = True Then
        cry_rep1.Formulas(5) = ("cond = 1")
        cry_rep1.Formulas(6) = ("speriod = 'FOR PRODUCTION STOPPAGE DAYS'")
    ElseIf opt_Presentdays.Value = True Then
        cry_rep1.Formulas(5) = ("cond = 0")
        cry_rep1.Formulas(6) = ("speriod = 'FOR WORKING DAYS'")
    End If
    
   cry_rep1.Formulas(1) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(2) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   cry_rep1.Formulas(4) = ("slot = '" & cmb_slot.Text & "'")
   
   SSLOT = ""
   If opt_cat_all.Value = True Then
       If opt_staff.Value = True Then
          SSLOT = "and {vew_emp_mas.emp_cat} = 'S'"
       ElseIf opt_worker.Value = True Then
          SSLOT = "and {vew_emp_mas.emp_cat} = 'W'"
       End If
   End If
   If opt_cat_trainee.Value = True Then
      If opt_staff.Value = True Then
           SSLOT = "and {vew_emp_mas.emp_type} = 1"
      Else
           SSLOT = "and {vew_emp_mas.emp_type} = 3"
      End If
  End If
 If opt_cat_perm.Value = True Then
      If opt_staff.Value = True Then
           SSLOT = "and {vew_emp_mas.emp_type} = 0"
      Else
           SSLOT = "and {vew_emp_mas.emp_type} = 2"
      End If
  End If
      
  If opt_cat_retain.Value = True Then
           SSLOT = "and {vew_emp_mas.emp_cat} = 'R'"
  End If
      
   
   If opt_selective_slot.Value = True Then
      SSLOT = "and {vew_emp_mas.emp_salary_slot} = '" & cmb_slot.Text & "' "
   End If
   
   If opt_selective_range.Value = True Then
      SSLOT = SSLOT + " and {emp_salary.s_netpay}  >= " & Val(txt_from.Text) & " and {emp_salary.s_netpay} <= " & Val(txt_to.Text) & ""
   End If
''   '''''''''''''''''
''       pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_slotwise]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_bank_salary_statement_slotwise] "
''        paydb.Execute (pst_qry)
''
'' pst_qry = "create view vew_bank_salary_statement_slotwise as " _
''                  & "select emp_name,emp_code ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
''                  & "(select emp_name,emp_code,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_year = " & cmb_year.Text & " and emp_bank=1" _
''                  & " Union All " _
''                  & " select emp_name ,emp_code,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "   and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank=1" _
''                  & " Union All " _
''                  & " select emp_name ,emp_code,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_year = " & cmb_year.Text & " and emp_bank=1 " _
''                  & " Union All " _
''                  & " select emp_name ,emp_code,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & cmb_year.Text & " and emp_bank=1" _
''                  & "  )a group by emp_name,emp_code ,emp_cat ,emp_workplace, emp_bank_acno "
''        paydb.Execute (pst_qry)
''    '''''''''''''''''''''''''''''''
''   '''''''''''''''''
       pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_slotwise]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_bank_salary_statement_slotwise] "
        paydb.Execute (pst_qry)

 pst_qry = "create view vew_bank_salary_statement_slotwise as " _
                  & "select s_company,s_year ,s_month ,s_empcat,s_empcode ,sum(s_netpay) as amount from " _
                  & "(select s_company,s_year ,s_month ,s_empcat,s_empcode ,s_netpay  from emp_salary  where s_year= " & cmb_year.Text & " and s_month=" & cmb_month.ItemData(cmb_month.ListIndex) & " " _
                  & " Union All " _
                  & " select ot_company as s_company,ot_year as s_year,ot_month as s_month,ot_emp_cat as s_empcat,ot_emp_code as s_empcode,ot_amount as s_netpay from overtime_entry where ot_year=" & cmb_year.Text & " and ot_month=" & cmb_month.ItemData(cmb_month.ListIndex) & " " _
                  & " )a group by s_empcode ,s_company,s_year ,s_month ,s_empcat,s_netpay "
        paydb.Execute (pst_qry)
    '''''''''''''''''''''''''''''''
   
   
   
   If Opt_waccno.Value = True Then
        If opt_alldays.Value = True Then
            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\slotwise_salarystatement_bank_new.rpt"
        Else
            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\slotwise_salarystatement_bank_forlayoff_new.rpt"
        End If
   Else
        If opt_alldays.Value = True Then
            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\slotwise_salarystatement_new.rpt"
        Else
            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\slotwise_salarystatement_forlayoff_new.rpt"
        End If
   End If
''   If opt_selective_mill.Value = True Then
''        If opt_staff.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_company} = " & company_code & " and {vew_bank_salary_statement_slotwise.s_netpay} > 0  and {vew_bank_salary_statement_slotwise.s_empcat} = 'S' " & sslot)
''        ElseIf opt_worker.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_company} = " & company_code & " and {vew_bank_salary_statement_slotwise.s_netpay} > 0 and {vew_bank_salary_statement_slotwise.s_empcat} = 'W' " & sslot)
''        Else
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_netpay} > 0 and {vew_bank_salary_statement_slotwise.s_company} = " & company_code & " " & sslot)
''        End If
''   Else
''        If opt_staff.Value = True Then
''            pst_qry = "{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and ({vew_bank_salary_statement_slotwise.s_company} = 1 or {vew_bank_salary_statement_slotwise.s_company} = 2) and {vew_bank_salary_statement_slotwise.s_netpay} > 0  and {vew_bank_salary_statement_slotwise.s_empcat} = 'S' " & sslot
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_company} <> 90 and {vew_bank_salary_statement_slotwise.s_netpay} > 0  and {vew_bank_salary_statement_slotwise.s_empcat} = 'S' " & sslot)
''        ElseIf opt_worker.Value = True Then
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_company} <> 90 and {vew_bank_salary_statement_slotwise.s_netpay} > 0 and {vew_bank_salary_statement_slotwise.s_empcat} = 'W' " & sslot)
''        Else
''            cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement_slotwise.s_year} = " & Val(cmb_year.Text) & " and {vew_bank_salary_statement_slotwise.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {vew_bank_salary_statement_slotwise.s_netpay} > 0 and {vew_bank_salary_statement_slotwise.s_company} <> 90 " & sslot)
''        End If
''
''   End If
    If opt_selective_mill.Value = True Then
        If opt_staff.Value = True Then
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_netpay} > 0  and ({emp_salary.s_empcat} = 'S' or {emp_salary.s_empcat} = 'R') " & SSLOT)
        ElseIf opt_worker.Value = True Then
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_netpay} > 0 and {emp_salary.s_empcat} = 'W' " & SSLOT)
        Else
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_netpay} > 0 and {emp_salary.s_company} = " & company_code & " " & SSLOT)
        End If
   Else
        If opt_staff.Value = True Then
            pst_qry = "{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and ({emp_salary.s_company} = 1 or {emp_salary.s_company} = 2) and {emp_salary.s_netpay} > 0  and ({emp_salary.s_empcat} = 'S' or {emp_salary.s_empcat} = 'R')  " & SSLOT
            pst_qry = "{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & " and {vew_emp_mas.emp_type} <> 90 and {emp_salary.s_company} <> 90 and {emp_salary.s_netpay} > 0  and ({emp_salary.s_empcat} = 'S' or {emp_salary.s_empcat} = 'R') " & SSLOT
                                             
      ''      If opt_cat_all.Value = True Then
               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_company} <> 90 and {emp_salary.s_netpay} > 0  and ({emp_salary.s_empcat} = 'S' or {emp_salary.s_empcat} = 'R') " & SSLOT)
                                             
               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_company} <> 90 and {emp_salary.s_netpay} > 0  " & SSLOT)
                                             
                                             
''            ElseIf opt_cat_trainee.Value = True Then
''               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             " and {emp_salary.s_company} <> 90 and {emp_salary.s_netpay} > 0  and ({emp_salary.s_empcat} = 'S' or {emp_salary.s_empcat} = 'R') " & SSLOT)
''            End If
                                             
        ElseIf opt_worker.Value = True Then
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_company} <> 90 and {emp_salary.s_netpay} > 0 and {emp_salary.s_empcat} = 'W' " & SSLOT)
        Else
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month} = " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_salary.s_netpay} > 0 and {emp_salary.s_company} <> 90 " & SSLOT)
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




    


