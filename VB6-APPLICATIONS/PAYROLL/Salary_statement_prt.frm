VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Salary_statement_prt 
   Caption         =   "SALARY STATEMENT PREPARATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   5175
      Left            =   10200
      TabIndex        =   45
      Top             =   960
      Width           =   7335
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   2160
         TabIndex        =   49
         Top             =   120
         Width           =   5055
         Begin VB.ListBox lst_dept 
            Enabled         =   0   'False
            Height          =   4110
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   50
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton opt_alldept 
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_selective_dept 
            Caption         =   "Selective"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   3495
      Left            =   11520
      TabIndex        =   39
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton opt_wp 
         Caption         =   "WORKER"
         Height          =   285
         Left            =   960
         TabIndex        =   44
         Top             =   2520
         Width           =   2775
      End
      Begin VB.OptionButton opt_salary_abstract 
         Caption         =   "SALARY ABSTRACT"
         Height          =   405
         Left            =   960
         TabIndex        =   43
         Top             =   3000
         Width           =   3255
      End
      Begin VB.OptionButton opt_wt 
         Caption         =   "TRAINEE  WORKER"
         Height          =   300
         Left            =   960
         TabIndex        =   42
         Top             =   1320
         Width           =   2835
      End
      Begin VB.OptionButton opt_st 
         Caption         =   "TRAINEE STAFF"
         Height          =   420
         Left            =   600
         TabIndex        =   41
         Top             =   600
         Width           =   3075
      End
      Begin VB.OptionButton opt_retainer 
         Caption         =   "RETAINER"
         Height          =   405
         Left            =   720
         TabIndex        =   40
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.CheckBox chk_deptwise 
      Caption         =   "DEPARTMENTWISE"
      Height          =   255
      Left            =   10440
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Height          =   495
      Left            =   9480
      TabIndex        =   31
      Top             =   9480
      Visible         =   0   'False
      Width           =   495
      Begin VB.OptionButton chk_a4 
         Caption         =   "A4"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton chk_a4_check 
         Caption         =   "A4 - FOR CHECKING"
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   255
      Left            =   14400
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      Begin VB.OptionButton opt1 
         Caption         =   "Statement"
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Designationwise"
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   1080
      TabIndex        =   15
      Top             =   9120
      Visible         =   0   'False
      Width           =   4935
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131792897
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131792897
         CurrentDate     =   39359
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
         TabIndex        =   19
         Top             =   240
         Width           =   1935
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
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   -240
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   105
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
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
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   9240
      Begin VB.Frame Frame4 
         Caption         =   "SALARY FOR THE MONTH OF "
         Height          =   1215
         Left            =   480
         TabIndex        =   8
         Top             =   5520
         Width           =   8415
         Begin VB.Frame Frame8 
            Enabled         =   0   'False
            Height          =   375
            Left            =   7200
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   735
            Begin VB.OptionButton opt_salary_doj 
               Caption         =   "DOJ FROM "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2640
               TabIndex        =   26
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton opt_salary_all 
               Caption         =   "All"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker dt_doj_from 
               Height          =   375
               Left            =   5280
               TabIndex        =   27
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   131792897
               CurrentDate     =   39359
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
            Left            =   5280
            TabIndex        =   12
            Top             =   480
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
            Left            =   1320
            TabIndex        =   10
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4200
            TabIndex        =   11
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3480
         TabIndex        =   5
         Top             =   7080
         Width           =   1695
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "Salary_statement_prt.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "Salary_statement_prt.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.ListBox deduct_list 
         Height          =   285
         Left            =   7560
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT "
         Height          =   7065
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   8640
         Begin VB.OptionButton opt_voucher_salary_manager 
            Caption         =   "Voucher Salary - Managers"
            Height          =   405
            Left            =   480
            TabIndex        =   59
            Top             =   1920
            Width           =   3495
         End
         Begin VB.OptionButton opt_voucher_salary 
            Caption         =   "Voucher Salary - for 1/2 Absent"
            Height          =   405
            Left            =   480
            TabIndex        =   58
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Frame frame_bank 
            Height          =   1575
            Left            =   4920
            TabIndex        =   52
            Top             =   3000
            Visible         =   0   'False
            Width           =   3255
            Begin VB.OptionButton opt_Bank 
               Caption         =   "BANK SALARY"
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
               Height          =   195
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton opt_cash 
               Caption         =   "CASH SALARY"
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
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   600
               Width           =   2655
            End
            Begin VB.OptionButton opt_all 
               Caption         =   "All Employees"
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
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   960
               Width           =   2655
            End
         End
         Begin VB.OptionButton opt_dept_bank 
            Caption         =   "DEPT WISE BANK SALARY ABSTRACT"
            Height          =   405
            Left            =   480
            TabIndex        =   51
            Top             =   720
            Width           =   3495
         End
         Begin VB.Frame frame_salary 
            Height          =   2415
            Left            =   4800
            TabIndex        =   35
            Top             =   120
            Width           =   3015
            Begin VB.OptionButton opt_OtherBank 
               Caption         =   "Other Banks"
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
               Height          =   195
               Left            =   120
               TabIndex        =   57
               Top             =   1440
               Width           =   2655
            End
            Begin VB.OptionButton opt_Cashmember 
               Caption         =   "Cash"
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
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   960
               Width           =   2655
            End
            Begin VB.OptionButton opt_all_members 
               Caption         =   "All Employees"
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
               Height          =   195
               Left            =   240
               TabIndex        =   38
               Top             =   1920
               Width           =   2655
            End
            Begin VB.OptionButton opt_nonpf_member 
               Caption         =   "Non PF Members"
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
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   2655
            End
            Begin VB.OptionButton opt_pf_member 
               Caption         =   "PF Members"
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
               Height          =   195
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Value           =   -1  'True
               Width           =   2055
            End
         End
         Begin VB.Frame frame_retainer 
            Caption         =   "Retainer options"
            Height          =   135
            Left            =   5760
            TabIndex        =   20
            Top             =   5160
            Width           =   1335
            Begin VB.OptionButton optr_tds_No 
               Caption         =   "TDS - NO"
               Height          =   195
               Left            =   3120
               TabIndex        =   23
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton optr_tds_yes 
               Caption         =   "TDS - YES"
               Height          =   195
               Left            =   1560
               TabIndex        =   22
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton optr_all 
               Caption         =   "All"
               Height          =   195
               Left            =   360
               TabIndex        =   21
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.OptionButton opt_sp 
            Caption         =   "SALARY STATEMENT"
            Height          =   405
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.ListBox std_deduct_lst 
         Height          =   285
         Left            =   7680
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   6120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "SELECT DEDUCTION LIST "
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
         Left            =   6840
         TabIndex        =   4
         Top             =   6960
         Visible         =   0   'False
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Salary_statement_prt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim noofdays_in_month As Integer

Private Sub cmb_month_Click()
find_dates
End Sub


Private Sub cmb_year_Click()
find_dates
End Sub

Private Sub Command1_Click()
Set salary_statement_staff.rpt = rptprint
With salary_statement_staff.rpt
    .ExportOptions.DiskFileName = "xx.pdf"
    .ExportOptions.DestinationType = crEDTDiskFile
    .ExportOptions.FormatType = crEFTPortableDocFormat
    .ExportOptions.PDFExportAllPages = True
    .Export False
    .PrinterSetup 0
End With

End Sub

Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
 ''    opt_sp.Enabled = True
    sql = "select * from pdept_mas   order by dept_name"
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_dept.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("dept_name")
        payrs.MoveNext
    Wend
    
    dt_doj_from = Now
    frame_retainer.Visible = False
    
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
    
    If optchk = 3 Then
       opt_st.Visible = False
       opt_wp.Visible = False
       opt_wt.Visible = False
    Else
       opt_st.Visible = True
       opt_wp.Visible = True
       opt_wt.Visible = True
    End If
    
End Sub
''
''Private Sub opt_sp_Click()
''    sql = ("Select * from  pdedu_mas order by pdedu_name")
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    std_deduct_lst.Clear
''    deduct_list.Clear
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 2 Then
''           std_deduct_lst.AddItem payrs(1)
''           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
''        End If
''        If payrs.Fields("pdedu_type") = 4 Then
''           deduct_list.AddItem payrs(1)
''           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
''        End If
''        payrs.MoveNext
''    Wend
''End Sub
''
''Private Sub opt_wp_Click()
''    sql = ("Select * from  pdedu_mas order by pdedu_name")
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    std_deduct_lst.Clear
''    deduct_list.Clear
''
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 3 Then
''           std_deduct_lst.AddItem payrs(1)
''           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
''        End If
''        If payrs.Fields("pdedu_type") = 4 Then
''           deduct_list.AddItem payrs(1)
''           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
''        End If
''        payrs.MoveNext
''    Wend
''End Sub

Private Sub opt_dept_bank_Click()
    frame_salary.Visible = False
    frame_bank.Visible = True
End Sub

Private Sub opt_retainer_Click()
    frame_retainer.Visible = True
End Sub

Private Sub opt_selective_emp_Click()
    lst_dept.Enabled = True
End Sub

Private Sub opt_selective_dept_Click()
     lst_dept.Enabled = True
End Sub

Private Sub opt_sp_Click()
    frame_retainer.Visible = False
    frame_salary.Visible = True
    frame_bank.Visible = False
End Sub

Private Sub opt_st_Click()
frame_retainer.Visible = False
End Sub

Private Sub opt_wp_Click()
frame_retainer.Visible = False
End Sub

Private Sub opt_wt_Click()
frame_retainer.Visible = False
End Sub

Private Sub Option1_Click()

End Sub

Private Sub PROCESS_Click()

   ded1_code = " "
   ded2_code = " "
   ded3_code = " "
   ded4_code = " "
   ded5_code = " "
   ded6_code = " "
   ded7_code = " "
   ded8_code = " "
   ded9_code = " "
   ded10_code = " "
   
''
''   If opt_salary_abstract.Value = True Then
''        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_salary_abstract]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_salary_abstract] "
''        paydb.Execute (pst_qry)
''        pst_qry = "create view vew_salary_abstract as" _
''             & " select emp_cat,emp_fpcode,emp_name,s_netpay from emp_mas, emp_salary where s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & cmb_year.Text & "  and s_netpay > 0 " _
''             & " Union All " _
''             & " select 'S' as emp_cat,emp_fpcode,emp_name,s_netpay from emp_voupay_mast, emp_salary where s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & cmb_year.Text & "  and s_netpay > 0 " _
''             & " Union All " _
''             & " select 'C' as emp_cat,ca_fpcode as emp_fpcode,ca_empname as emp_name,c_netpay as s_netpay  from trn_casalary,mas_caemp where ca_empcode = c_empcode  and c_deptcode not in (28,36) and c_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and c_year = " & cmb_year.Text & "  and c_netpay > 0"
''       paydb.Execute (pst_qry)
''
''       cry_rep1.Formulas(1) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
''       gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''       cry_rep1.PrinterSelect
''
''
''       cry_rep1.Formulas(2) = ("millname= '" & compname & "'")
''       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salaray_abstract.rpt"
''       cry_rep1.WindowState = crptMaximized
''       cry_rep1.Connect = gst_repconnect
''       cry_rep1.Action = 1
''       Exit Sub
''   End If
''
   If opt_retainer.Value = True Then
      If optr_all.Value = False And optr_tds_yes.Value = False And optr_tds_No.Value = False Then
         MsgBox ("Select Retainer option...")
         Exit Sub
      End If
   End If
   If opt_sp.Value = True Then
      If opt_pf_member.Value = True Then
        disname = "SALARY STATEMENT FOR THE MONTH OF "
      ElseIf opt_nonpf_member.Value = True Then
        disname = "SALARY STATEMENT FOR THE MONTH OF "
      Else
        disname = "SALARY STATEMENT FOR THE MONTH OF "
      End If
   End If
''   If opt_st.Value = True Then disname = "TRAINEE STAFF SALARY STATEMENT FOR THE MONTH OF "
''
''   If opt_wp.Value = True Then
''      If opt_pf_member.Value = True Then
''        disname = "WORKER WAGES STATEMENT FOR THE MONTH OF "
''      ElseIf opt_nonpf_member.Value = True Then
''        disname = "WORKER EX-GRATIA STATEMENT FOR THE MONTH OF "
''      Else
''        disname = "WORKER WAGES STATEMENT FOR THE MONTH OF  "
''      End If
''   End If
''
''   If opt_wt.Value = True Then disname = "TRAINEE WORKER WAGES STATEMENT FOR THE MONTH OF "
''   If opt_retainer.Value = True Then disname = "RETAINER AMOUNT STATEMENT FOR THE MONTH OF "
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   Dim ds, doj As String
''   If optchk = 2 Then
''      ds = " and {emp_mas.emp_classification} = 'A'"
''   ElseIf optchk = 1 Then
''      ds = " and {emp_mas.emp_classification} = 'B'"
''   Else
''    ds = ""
''   End If


   dept = ""
   If opt_selective_dept.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                     dept = " and ( {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                     i = i + 1
                  Else
                     dept = dept + " or {pdept_mas.dept_name}= '" & lst_dept.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If dept <> "" Then dept = dept + ")"
   ds = ds + dept
   
   
   cry_rep1.Formulas(5) = ""
'''   If optchk = 1 Then
'''''      ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_classification} = 'B'"
'''      ds = " and {emp_mas.emp_workplace} = 'MIL'"
'''      cry_rep1.Formulas(5) = ("e_cat = ' MILL -  BELOW MANAGERS   '")
'''   ElseIf optchk = 2 Then
'''      ds = " and {emp_mas.emp_workplace} = 'CBE' "
'''      cry_rep1.Formulas(5) = ("e_cat = ' CBE - STAFF SALARY  '")
'''   ElseIf optchk = 3 Then
'''      ds = " and {emp_mas.emp_classification} = 'A'"
'''      cry_rep1.Formulas(5) = ("e_cat = ' ABOVE MANAGERS   '")
'''   Else
'''      ds = ""
'''   End If
   
   

   cry_rep1.Formulas(5) = ("e_cat = ''")
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   
   doj = "JOINED FROM " + CStr(Format(dt_doj_from.Value, "dd/MM/yyyy"))
   
   If company_code = 1 Then
      cry_rep1.Formulas(2) = millname
   End If
   
   cry_rep1.Formulas(2) = ("millname= '" & compname & "'")
   
   cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
''   If opt_salary_doj.Value = True Then
''      cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
''   Else
''      cry_rep1.Formulas(5) = ""
''   End If

   
'''   If opt_retainer.Value = True Then
'''        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_retainer.rpt"
'''   ElseIf chk_deptwise.Value = True And opt_sp.Value = True Then
'''        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\deptwise_salary_statement.rpt"
'''   ElseIf opt_sp.Value = True And chk_deptwise.Value = False Then
'''        If chk_a4.Value = True Or chk_a4_check.Value = True Then
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_staff.rpt"
'''        Else
'''            If opt_st.Value = True Then
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_staff.rpt"
'''            Else
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement.rpt"
'''            End If
'''        End If
'''   ElseIf opt_st.Value = True And chk_deptwise.Value = False Then
'''        If chk_a4.Value = True Or chk_a4_check.Value = True Then
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_staff_trainee_a4.rpt"
'''        Else
'''            If opt1.Value = True Then
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_staff_trainee.rpt"
'''            Else
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement.rpt"
'''            End If
'''        End If
'''        cry_rep1.Formulas(5) = ""
'''   ElseIf opt_wp.Value = True Or chk_deptwise.Value = False Then
'''        If chk_a4.Value = True Then
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_a4.rpt"
'''        ElseIf chk_a4_check.Value = True Then
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_a4_forchecking.rpt"
'''        Else
'''            If opt_st.Value = True Then
'''                cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker.rpt"
'''            Else
'''                cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_details.rpt"
'''            End If
'''        End If
'''
'''        cry_rep1.Formulas(5) = ""
'''   Else
'''        If chk_a4.Value = True Or chk_a4_check.Value = True Then
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_trainee_a4.rpt"
'''        Else
'''            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_worker_trainee.rpt"
'''        End If
'''         cry_rep1.Formulas(5) = ""
'''
'''   End If
'''   If opt_sp.Value = True Then
'''      pst_qry = ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'S' and {emp_salary.s_salarydays} > 0  " & ds & " and ({emp_mas.emp_status}  ='A' or ({emp_mas.emp_status}  ='R' and month({emp_mas.emp_resigneddate}) <=  " & cmb_month.ItemData(cmb_month.ListIndex) & " and year({emp_mas.emp_resigneddate}) <=  " & Val(cmb_year.Text) & "  ))")
'''      If opt_salary_all.Value = True Then
'''         If opt_all_members.Value = True Then
'''             cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'S' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''         ElseIf opt_pf_member.Value = True Then
'''             cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_pf} >0  and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'S' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''         Else
'''         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_pf} =0  and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'S' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''         End If
'''
'''      Else
'''         cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'S' and {emp_salary.s_salarydays} > 0  " & ds & "  and {emp_mas.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ")   ")
'''         cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
'''      End If
'''
'''''     cry_rep1.Formulas(5) = ("rtype =   ' JOINED FROM ' + '" & Trim(cmb_month2.Text) + cmb_year2.Text & "'")
'''''     cry_rep1.Formulas(5) = ("rtype =   ' JOINED FROM ' + '" & Trim(cmb_month2.Text) & "'")
'''''       cry_rep1.Formulas(5) = ("rtype = 'TEST'")
'''     ''cry_rep1.Formulas(5) = ("rtype = 'JOINED FROM' ")
'''   Else
'''      If opt_st.Value = True Then
'''         If opt_salary_all.Value = True Then
'''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                         "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1 and {emp_salary.s_salarydays} > 0  " & ds & "")
'''         Else
'''            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                         "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 1 and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_mas.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ") ")
'''
'''         cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
'''
'''         End If
'''      Else
'''         If opt_wp.Value = True Then
'''            If opt_salary_all.Value = True Then
'''               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'W' and {emp_salary.s_salarydays} > 0  " & ds & "")
'''            Else
'''               cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 'W' and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_mas.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ") ")
'''
'''            cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
'''
'''            End If
'''
'''         Else
'''            If opt_wt.Value = True Then
'''               If opt_salary_all.Value = True Then
'''                  If opt_salary_all.Value = True Then
'''                      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                       "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 and {emp_salary.s_salarydays} > 0  " & ds & "")
'''
'''                  Else
'''                      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                       "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_mas.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ") ")
'''
'''                     cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
'''
'''                  End If
'''               Else
'''                    cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                       "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_type} = 3 and {emp_salary.s_salarydays} > 0  " & ds & "  ")
'''
'''               End If
'''            Else
'''                If opt_retainer.Value = True Then
'''                   If optr_all.Value = True Then
'''                        If opt_salary_all.Value = True Then
'''                            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_company} = " & company_code & " and {emp_voupay_mast.emp_cat} = 'R' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''                        Else
'''                            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_company} = " & company_code & " and {emp_voupay_mast.emp_cat} = 'R' and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_voupay_mast.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ")  ")
'''
'''                            pst_qry = "{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        " and {emp_salary.s_company} = " & company_code & " and {emp_voupay_mast.emp_cat} = 'R' and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_voupay_mast.EMP_DOJ} >=   date(" & Format(dt_doj_from.Value, "yyyy,mm,dd") & ")  "
'''                            cry_rep1.Formulas(5) = ("rtype = '" & doj & "'")
'''
'''
'''                        End If
'''                   ElseIf optr_tds_yes.Value = True Then
'''                        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        "  and {emp_voupay_mast.emp_tds_per} > 0  and {emp_salary.s_company} = " & company_code & " and {emp_voupay_mast.emp_cat} = 'R' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''                   Else
'''                        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                        "  and {emp_voupay_mast.emp_tds_per} = 0  and {emp_salary.s_company} = " & company_code & " and {emp_voupay_mast.emp_cat} = 'R' and {emp_salary.s_salarydays} > 0  " & ds & " ")
'''                   End If
'''
'''                End If
'''            End If
'''         End If
'''      End If
'''   End If
   
   cry_rep1.Formulas(6) = ("report_year = " & Val(cmb_year.Text))
   

   If opt_voucher_salary.Value = True Then
   
         cry_rep1.Formulas(6) = ("noofdays_in_month = " & Val(noofdays_in_month))
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Voucher_Payment.rpt"
         cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_sft_hrs} > 8.50 and {bio_device_shiftlogs.ds_status} = '½P½A' and ({bio_device_shiftlogs.ds_fpcode} <>  1006 and {bio_device_shiftlogs.ds_fpcode} <>  1252  and {bio_device_shiftlogs.ds_fpcode} <>  1810 and {bio_device_shiftlogs.ds_fpcode} <>  1127 and {bio_device_shiftlogs.ds_fpcode} <>  3202 )")
         
   ElseIf opt_voucher_salary_manager.Value = True Then
         cry_rep1.Formulas(6) = ("noofdays_in_month = " & Val(noofdays_in_month))
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Voucher_Payment_manager.rpt"
         qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and {bio_attendlogs.a_salary_days} > 0  and ({bio_attendlogs.a_fp_code} = 1006 or {bio_attendlogs.a_fp_code} = 1252)"
' ' "3515" Or txt_empcode.Text = "1006" Or txt_empcode.Text = "1252" Or txt_empcode.Text = "1810" Or txt_empcode.Text = "1127" Or txt_empcode.Text = "3202" Or txt_empcode.Text = "1018" Then
         cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and {bio_attendlogs.a_salary_days} > 0  and ({bio_attendlogs.a_fpcode} = 1006 or {bio_attendlogs.a_fpcode} = 1252  or {bio_attendlogs.a_fpcode} = 1810 or {bio_attendlogs.a_fpcode} = 1127 or {bio_attendlogs.a_fpcode} = 3202 ) ")
         
             
   ElseIf opt_dept_bank.Value = True Then
   
      
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\bank_salary_statement_abstract.rpt"
       If opt_Bank.Value = True Then
           cry_rep1.Formulas(6) = ("salary_type = 'BANK'")
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                     " and {emp_salary.s_salary_bank} = 1 and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       ElseIf opt_cash.Value = True Then
            cry_rep1.Formulas(6) = ("salary_type = 'CASH'")
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                     " and {emp_salary.s_salary_bank} = 2 and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       Else
            cry_rep1.Formulas(6) = ("salary_type = 'ALL'")
            cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                     " and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       
       End If
   Else
''        If opt_nonpf_member.Value = True Then
''           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_nonpf.rpt"
''        Else
''           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement.rpt"
''        End If
        
       If opt_all_members.Value = True Then
                  cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement.rpt"

           cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       ElseIf opt_pf_member.Value = True Then
             cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement.rpt"
    ''       cry_rep1.ReplaceSelectionFormula (" {emp_mas.EMP_PFELIGIBLE} = 'Y' and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
      
           cry_rep1.ReplaceSelectionFormula (" {emp_salary.s_pf_eligible} = 'Y' and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_salary.s_salary_bank} = 1 ")
       ElseIf opt_nonpf_member.Value = True Then
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_nonpf.rpt"
           cry_rep1.ReplaceSelectionFormula (" {emp_salary.s_pf_eligible} = 'N' and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_salary.s_salary_bank} = 1 ")
       
       ElseIf opt_Cashmember.Value = True Then
       
            cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_nonpf.rpt"
           cry_rep1.ReplaceSelectionFormula (" {emp_salary.s_pf_eligible} = 'N' and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_salary.s_salary_bank} = 2 ")
       ElseIf opt_OtherBank.Value = True Then
       
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\salary_statement_nonpf.rpt"
           cry_rep1.ReplaceSelectionFormula (" {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_salary.s_salary_bank} > 2 ")
                                      
       
       Else
    ''       cry_rep1.ReplaceSelectionFormula (" {emp_mas.EMP_PFELIGIBLE} = 'N' and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       
           cry_rep1.ReplaceSelectionFormula (" {emp_salary.s_pf_eligible} = 'N'  and {emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                      "and {emp_salary.s_company} = " & company_code & " and {emp_salary.s_salarydays} > 0  " & ds & "")
       
       End If
       
       c = 1
       Dim item As String
       pin_fms = 5
       If std_deduct_lst.ListCount > 0 Then
          For pin_row = 0 To std_deduct_lst.ListCount - 1
    ''          If std_deduct_lst.Selected(pin_row) = True Then
                cry_rep1.Formulas(pin_fms) = "ded" & (c) & "_code = " & Val(std_deduct_lst.ItemData(pin_row))
                std_deduct_lst.ListIndex = Val(pin_row)
                cry_rep1.Formulas(pin_fms + 1) = "ded" & (c) & " = '" & Trim$(std_deduct_lst.Text) & "'"
                pin_sel_item = Val(pin_sel_item) + 1
                c = c + 1
                pin_fms = pin_fms + 2
    ''          End If
          Next pin_row
       End If
       If deduct_list.ListCount > 0 Then
          For pin_row = 0 To deduct_list.ListCount - 1
              If pin_fms > 21 Then Exit For
              If deduct_list.Selected(pin_row) = True Then
                 cry_rep1.Formulas(pin_fms) = "ded" & (c) & "_code = " & Val(deduct_list.ItemData(pin_row))
                 deduct_list.ListIndex = Val(pin_row)
                 cry_rep1.Formulas(pin_fms + 1) = "ded" & (c) & " = '" & Trim$(deduct_list.Text) & "'"
                 pin_sel_item = Val(pin_sel_item) + 1
                 c = c + 1
                 pin_fms = pin_fms + 2
              Else
                 cry_rep1.Formulas(pin_fms) = ""
                 cry_rep1.Formulas(pin_fms + 1) = ""
                 pin_fms = pin_fms + 2
              End If
            Next pin_row
       End If
   End If
''   cry_rep1.Formulas(5) = ("rtype = 'JOINED FROM '")
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
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
    
    noofdays_in_month = mdays
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
End Sub
