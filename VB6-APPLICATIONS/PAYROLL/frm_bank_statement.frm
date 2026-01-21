VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_bank_statement 
   Caption         =   "BANK STATEMENT"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame12 
      Caption         =   "Frame12"
      Height          =   4215
      Left            =   10440
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   1455
         Left            =   240
         TabIndex        =   49
         Top             =   2640
         Width           =   3615
         Begin VB.OptionButton opt_retainer 
            Caption         =   "RETAINERS"
            Height          =   285
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   2175
         End
         Begin VB.OptionButton opt_below_manager 
            Caption         =   "BELOW MANAGERS"
            Height          =   285
            Left            =   0
            TabIndex        =   51
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opt_above_manager 
            Caption         =   "MANAGERS && ABOVE"
            Height          =   285
            Left            =   0
            TabIndex        =   50
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "SELECT"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   2040
         TabIndex        =   45
         Top             =   1680
         Width           =   1815
         Begin VB.OptionButton opt_all_layoff 
            Caption         =   "ALL"
            Height          =   525
            Left            =   0
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_layoff 
            Caption         =   "EX-GRATIA"
            Height          =   405
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton opt_workeddays 
            Caption         =   "WORKED DAYS"
            Height          =   405
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "SELECT "
         Height          =   1065
         Left            =   480
         TabIndex        =   41
         Top             =   1920
         Width           =   1080
         Begin VB.OptionButton opt_all 
            Caption         =   "ALL"
            Height          =   525
            Left            =   480
            TabIndex        =   44
            Top             =   1080
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_vpt 
            Caption         =   "MILL"
            Height          =   405
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton opt_cbe 
            Caption         =   "OTHER"
            Height          =   405
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame Frame9 
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
         Left            =   600
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
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
            Height          =   615
            Left            =   1920
            TabIndex        =   35
            Top             =   480
            Width           =   975
            Begin VB.OptionButton opt_solvent 
               Caption         =   "SOLVENT"
               Height          =   375
               Left            =   4080
               TabIndex        =   40
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton opt_cogen 
               Caption         =   "COGEN"
               Height          =   375
               Left            =   3000
               TabIndex        =   39
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_vjpm 
               Caption         =   "VJPM"
               Height          =   375
               Left            =   2040
               TabIndex        =   38
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_slpb 
               Caption         =   "DPM-II"
               Height          =   375
               Left            =   1200
               TabIndex        =   37
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt_dpm 
               Caption         =   "DPM"
               Height          =   375
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.OptionButton opt_millselective 
            Caption         =   "SELECTIVE"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt_millall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   975
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
         Left            =   3840
         TabIndex        =   25
         Top             =   600
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txt_from 
            Height          =   495
            Left            =   600
            TabIndex        =   27
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txt_to 
            Height          =   495
            Left            =   600
            TabIndex        =   26
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   2400
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BANK STATEMENT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      Begin VB.Frame Frame7 
         Height          =   975
         Left            =   840
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   7695
         Begin VB.OptionButton opt_rep2 
            Caption         =   "Bank Statement (Breakup)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   240
            Width           =   3135
         End
         Begin VB.OptionButton opt_rep1 
            Caption         =   "Bank Statement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   720
         TabIndex        =   12
         Top             =   3120
         Width           =   8415
         Begin VB.Frame Frame10 
            Height          =   735
            Left            =   360
            TabIndex        =   20
            Top             =   120
            Width           =   7695
            Begin VB.OptionButton opt_all_banks_cash 
               Caption         =   "BANK && CASH"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton opt_selective_bank 
               Caption         =   "SELECTIVE BANK"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4560
               TabIndex        =   22
               Top             =   240
               Width           =   2655
            End
            Begin VB.OptionButton opt_all_banks 
               Caption         =   "ALL BANKS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2640
               TabIndex        =   21
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.CommandButton cmd_refresh 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7920
            TabIndex        =   19
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox cmb_bank 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label54 
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   960
         TabIndex        =   7
         Top             =   5280
         Width           =   7455
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
            Left            =   5160
            TabIndex        =   11
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
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4080
            TabIndex        =   10
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT "
         Height          =   1545
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   7680
         Begin VB.OptionButton opt_staff_worker 
            Caption         =   " ALL (Staff && Worker)"
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   " STAFF"
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3840
         TabIndex        =   1
         Top             =   6120
         Width           =   1695
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "frm_bank_statement.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "frm_bank_statement.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   720
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_bank_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Refresh_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    opt_all_layoff.Value = True
    sql = "select * from payroll_bank order by bank_name"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_bank.AddItem payrs("bank_name")
        cmb_bank.ItemData(cmb_bank.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    
End Sub

Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    opt_all_layoff.Value = True
    sql = "select * from payroll_bank order by bank_name"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_bank.AddItem payrs("bank_name")
        cmb_bank.ItemData(cmb_bank.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend

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
''    cmb_year.Text = "2015"
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
End Sub

Private Sub opt_sp_Click()
    sql = ("Select * from  pdedu_mas order by pdedu_name")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    std_deduct_lst.Clear
    deduct_list.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 2 Then
           std_deduct_lst.AddItem payrs(1)
           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
        End If
        If payrs.Fields("pdedu_type") = 4 Then
           deduct_list.AddItem payrs(1)
           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
        End If
        payrs.MoveNext
    Wend
End Sub
Private Sub opt_wp_Click()
    sql = ("Select * from  pdedu_mas order by pdedu_name")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    std_deduct_lst.Clear
    deduct_list.Clear
    
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        If payrs.Fields("pdedu_type") = 1 Or payrs.Fields("pdedu_type") = 3 Then
           std_deduct_lst.AddItem payrs(1)
           std_deduct_lst.ItemData(std_deduct_lst.NewIndex) = payrs(0)
        End If
        If payrs.Fields("pdedu_type") = 4 Then
           deduct_list.AddItem payrs(1)
           deduct_list.ItemData(deduct_list.NewIndex) = payrs(0)
        End If
        payrs.MoveNext
    Wend
End Sub


Private Sub opt_worker_Click()
 Frame8.Enabled = True
End Sub

Private Sub PROCESS_Click()
   Dim wp, qry1, qry2, qry3 As String
   If Trim(cmb_month.Text) = "" Then
      MsgBox ("Select the Reporting Month")
      Exit Sub
   End If
   MousePointer = vbDefault
   qry1 = ""
   qry2 = ""
   qry3 = ""
   If opt_selective_bank.Value = True Then
      If cmb_bank.ListIndex = -1 Then
         MsgBox ("Select Bank Name...")
         Exit Sub
      End If
   End If
   
   If opt_rep1.Value = True And opt_all_layoff.Value = True Then
        
        wp = "MIL "
        qry1 = ""
        If opt_staff.Value = True Then
           qry2 = "and emp_cat = 'S'"
           qry2 = "and (emp_cat = 'S' OR emp_cat = 'R') "
        ElseIf opt_worker.Value = True Then
           qry2 = " and emp_cat = 'W'"
        ElseIf opt_retainer.Value = True Then
           qry2 = " and emp_cat = 'R'"
        ElseIf opt_above_manager.Value = True Then
           qry2 = " and emp_cat = 'S' and emp_classification = 'A' "
        ElseIf opt_below_manager.Value = True Then
           qry2 = " and emp_cat = 'S' and emp_classification = 'B' "
        Else
           qry2 = ""
        End If
        
        
        cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
        cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
        cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
        cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
        cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
        cry_rep1.Formulas(6) = ""
        
        
        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_bank_salary_statement] "
        paydb.Execute (pst_qry)
''        pst_qry = "create view vew_bank_salary_statement as " _
''                  & "select emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
''                  & "(select emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "" _
''                  & "  )a group by emp_name ,emp_cat ,emp_workplace, emp_bank_acno "


''        If cmb_bank.ListIndex = -1 Then
'''        If opt_all_banks.Value = True Or opt_all_banks_cash = True Then
'''
'''            pst_qry = "create view vew_bank_salary_statement as " _
'''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc from " _
'''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc from  emp_voupay_mast a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " " & qry2 & " " _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " " & qry2 & "" _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount ,emp_bank,emp_bank_ifsc from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & "  " & qry2 & "" _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " " & qry2 & "" _
'''                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc "
'''
'''''            pst_qry = "create view vew_bank_salary_statement as " _
'''''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc from " _
'''''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount,emp_bank,emp_bank_ifsc from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " " & qry2 & " " _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " " & qry2 & "" _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount ,emp_bank,emp_bank_ifsc from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & "  " & qry2 & "" _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " " & qry2 & "" _
'''''                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc "
'''''
'''
'''        Else
'''''            pst_qry = "create view vew_bank_salary_statement as " _
'''''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
'''''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount from  emp_voupay_mast a, emp_salary b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''''                  & " Union All " _
'''''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "" _
'''''                   & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
'''
'''
'''            pst_qry = "create view vew_bank_salary_statement as " _
'''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc from " _
'''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc from  emp_voupay_mast a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount ,emp_bank,emp_bank_ifsc from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''                  & " Union All " _
'''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc "
'''
'''        End If
                  
''            pst_qry = "create view vew_bank_salary_statement as " _
''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "" _
''                   & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
''            End If
''

        If opt_all_banks.Value = True Or opt_all_banks_cash = True Then
        
            pst_qry = "create view vew_bank_salary_statement as " _
                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc, emp_dept,s_salary_bank from " _
                  & "(select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc, emp_dept ,s_salary_bank   from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " " & qry2 & "" _
                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc, emp_dept,s_salary_bank  "
         
''            pst_qry = "create view vew_bank_salary_statement as " _
''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc from " _
''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount,emp_bank,emp_bank_ifsc from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount ,emp_bank,emp_bank_ifsc from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & "  " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount,emp_bank,emp_bank_ifsc  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " " & qry2 & "" _
''                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc "
''
         
        Else
''            pst_qry = "create view vew_bank_salary_statement as " _
''                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
''                  & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount from  emp_voupay_mast a, emp_salary b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
''                  & " Union All " _
''                  & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "" _
''                   & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
                   
                   
            pst_qry = "create view vew_bank_salary_statement as " _
                  & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,emp_bank,emp_bank_ifsc , emp_dept ,s_salary_bank from " _
                  & "(select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount,emp_bank,emp_bank_ifsc , emp_dept ,s_salary_bank from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
                  & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno,emp_bank,emp_bank_ifsc, emp_dept ,s_salary_bank"
                   
        End If

        
        
        paydb.Execute (pst_qry)
        cry_rep1.PrinterSelect
        If cmb_bank.ListIndex = -1 Then
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\all_bank_salary_statement.rpt"
        Else
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\bank_statement.rpt"
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\all_bank_salary_statement.rpt"
        End If
        If opt_all_salary_range.Value = False Then
           If opt_all_banks.Value = True Then
               cry_rep1.ReplaceSelectionFormula ("{payroll_bank.bank_code} <> 0 and {vew_bank_salary_statement.amt} >= " & Val(txt_from.Text) & " and {vew_bank_salary_statement.amt} <= " & Val(txt_to.Text))
           Else
               cry_rep1.ReplaceSelectionFormula ("{vew_bank_salary_statement.amt} >= " & Val(txt_from.Text) & " and {vew_bank_salary_statement.amt} <= " & Val(txt_to.Text))
           End If
        Else
           If opt_all_banks.Value = True Then
               cry_rep1.Formulas(6) = ("reptype = 'ONLY BANK A/C'")
               cry_rep1.ReplaceSelectionFormula ("{payroll_bank.bank_code} <> 0 ")
        
           ElseIf cmb_bank.Text <> "CASH" Then
               cry_rep1.Formulas(6) = ("reptype = 'ONLY BANK A/C'")
               cry_rep1.ReplaceSelectionFormula ("{payroll_bank.bank_code} = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " ")
          
           ElseIf cmb_bank.Text = "CASH" Then
               cry_rep1.ReplaceSelectionFormula ("{payroll_bank.bank_code} =  2 ")
               cry_rep1.Formulas(6) = ("reptype = 'ONLY CASH '")
           
           Else
               cry_rep1.ReplaceSelectionFormula ("")
               cry_rep1.Formulas(6) = ("reptype = 'BANK & CASH '")
               
           End If
        
        End If

        
   ElseIf opt_rep2.Value = True Then
          wp = ""
          qry1 = ""

         If opt_staff.Value = True Then
               qry2 = "and emp_cat = 'S'"
         ElseIf opt_worker.Value = True Then
               qry2 = " and emp_cat = 'W'"
         ElseIf opt_retainer.Value = True Then
               qry2 = " and emp_cat = 'R'"
''         ElseIf opt_above_manager.Value = True Then
''               qry2 = " and emp_cat = 'S' and emp_classification = 'A' "
''         ElseIf opt_below_manager.Value = True Then
''               qry2 = " and emp_cat = 'S' and emp_classification = 'B' "
''         Else
               qry2 = ""
         End If
''         pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_bank_salary_statement] "
''         paydb.Execute (pst_qry)
         pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_breakup]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                      & " drop view [dbo].[vew_bank_salary_statement_breakup]"
         paydb.Execute (pst_qry)
''''         paydb.Execute (pst_qry)
'''        If opt_all_banks.Value = True Then
'''''            pst_qry = "create view vew_bank_salary_statement as " _
'''''                     & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
'''''                     & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " " & qry2 & " " _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & "  " & qry2 & "" _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " " & qry2 & "" _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " " & qry2 & "" _
'''''                     & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
'''''             paydb.Execute (pst_qry)
'''''
'''             pst_qry = "create view vew_bank_salary_statement_breakup as " _
'''                     & "select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,paytype,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
'''                     & "(select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,'Voucher' as paytype, emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount from  emp_voupay_mast a, emp_salary b , pdept_mas c where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & "  " & qry2 & " and emp_dept=dept_code " _
'''                     & " Union All " _
'''                     & " select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,'Salary' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b , pdept_mas c where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & "" & qry2 & " and emp_dept=dept_code " _
'''                     & " Union All " _
'''                     & " select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,'OT' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b , pdept_mas c  where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & "" & qry2 & " and emp_dept=dept_code " _
'''                     & " Union All " _
'''                     & " select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,'Additional' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b , pdept_mas c where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & "" & qry2 & " and emp_dept=dept_code " _
'''                     & "  )a group by emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,paytype, emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
'''             paydb.Execute (pst_qry)
'''         Else
'''''             pst_qry = "create view vew_bank_salary_statement as " _
'''''                     & "select emp_code,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
'''''                     & "(select emp_code,emp_name,emp_cat ,emp_workplace , emp_bank_acno, ot_amount as amount from  emp_voupay_mast a, emp_otherpayment_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and ot_company = " & company_code & " and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & "" _
'''''                     & " Union All " _
'''''                     & " select emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "" _
'''''                     & "  )a group by emp_code,emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
'''''                              paydb.Execute (pst_qry)
'''''
'''''             paydb.Execute (pst_qry)
'''             pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_breakup]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
'''                   & " drop view [dbo].[vew_bank_salary_statement_breakup]"
'''            paydb.Execute (pst_qry)
'''            pst_qry = "create view vew_bank_salary_statement_breakup as " _
'''                     & "select paytype,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
'''                     & "(select 'Voucher' as paytype, emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount from  emp_voupay_mast a, emp_salary b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''                     & " Union All " _
'''                     & " select 'Salary' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "  " _
'''                     & " Union All " _
'''                     & " select 'OT' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''                     & " Union All " _
'''                     & " select 'Additional' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
'''                     & "  )a group by paytype, emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
'''
'''
'''            paydb.Execute (pst_qry)
'''
'''         End If
         
        
        
''         pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_breakup]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_bank_salary_statement_breakup]"
''         paydb.Execute (pst_qry)
''         pst_qry = "create view vew_bank_salary_statement_breakup as " _
''                  & "select paytype,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt from " _
''                  & "(select 'Voucher' as paytype, emp_name,emp_cat ,emp_workplace , emp_bank_acno, s_netpay as amount from  emp_voupay_mast a, emp_salary b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_company = " & company_code & " and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select 'Salary' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount  from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "  " _
''                  & " Union All " _
''                  & " select 'OT' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, ot_amount as amount  from  emp_mas a, overtime_entry  b where emp_company = ot_company and emp_code = ot_emp_code and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and ot_company = " & company_code & "   and ot_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & " Union All " _
''                  & " select 'Additional' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, e_amount as amount  from  emp_mas a, employee_additional_amount b where emp_company = e_company and emp_code = e_emp_code and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = " & company_code & "  and e_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " " & qry2 & " " _
''                  & "  )a group by paytype, emp_name ,emp_cat ,emp_workplace, emp_bank_acno "
''
''
''         paydb.Execute (pst_qry)


        If opt_all_banks.Value = True Then
             pst_qry = "create view vew_bank_salary_statement_breakup as " _
                     & "select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,paytype,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,s_salary_bank from " _
                     & "(select emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,'Salary' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount ,s_salary_bank  from  emp_mas a, emp_salary  b , pdept_mas c where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & "" & qry2 & " and emp_dept=dept_code " _
                     & "  )a group by emp_company, emp_code,emp_bank, dept_name,emp_bank_ifsc,paytype, emp_name ,emp_cat ,emp_workplace, emp_bank_acno,s_salary_bank "
             paydb.Execute (pst_qry)
         Else
''             pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_bank_salary_statement_breakup]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_bank_salary_statement_breakup]"
''            paydb.Execute (pst_qry)
            pst_qry = "create view vew_bank_salary_statement_breakup as " _
                     & "select paytype,emp_name ,emp_cat ,emp_workplace , emp_bank_acno, sum(amount) as amt ,s_salary_bank from " _
                     & "(select 'Salary' as paytype,emp_name ,emp_cat ,emp_workplace, emp_bank_acno, s_netpay as amount ,s_salary_bank from  emp_mas a, emp_salary  b where emp_company = s_company and emp_code = s_empcode and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_company = " & company_code & " and s_netpay > 0  and s_year = " & cmb_year.Text & " and emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & "" & qry2 & "  " _
                     & "  )a group by paytype, emp_name ,emp_cat ,emp_workplace, emp_bank_acno,s_salary_bank "
                     
                     
            paydb.Execute (pst_qry)
         
         End If

         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\bank_statement_breakup.rpt"
        cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
        cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
        cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
        cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
        cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
        cry_rep1.Formulas(6) = ""

   Else
        gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
        cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Bank_statement_worker_forlayoff.rpt"
        cry_rep1.PrinterSelect
        If opt_layoff.Value = True Then
             cry_rep1.Formulas(8) = ("cond = 1")
        ElseIf opt_workeddays.Value = True Then
             cry_rep1.Formulas(8) = ("cond = 0")
        End If
        cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
        cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
        cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
        cry_rep1.Formulas(3) = ("sthead = '" & disname & "'")
        cry_rep1.Formulas(4) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
        cry_rep1.Formulas(6) = ""
        cry_rep1.Formulas(5) = ("bank = '" & cmb_bank.Text & "'")
        cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
             "and {emp_salary.s_company} = " & company_code & "  and {emp_mas.emp_bank}=" & cmb_bank.ItemData(cmb_bank.ListIndex) & " and {emp_salary.s_empcat} ='W' and {emp_salary.s_salarydays} > 0  " & ds & "")
   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.Action = 1
   Exit Sub
    
     If opt_staff.Value = True Then disname = wp + " - STAFF BANK STATEMENT "
     If opt_worker.Value = True Then disname = wp + " - WORKER BANK STATEMENT "
     If opt_staff_worker.Value = True Then disname = wp + " - STAFF & WORKER BANK STATEMENT "
     If opt_above_manager.Value = True Then disname = wp + " - MANAGER AND ABOVE - BANK STATEMENT "
     If opt_below_manager.Value = True Then disname = wp + " - BELOW MANAGER - BANK STATEMENT "
     
    
     pst_qry = ""
     gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
     cry_rep1.PrinterSelect
     cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
     cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
     cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
     cry_rep1.Formulas(3) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
     cry_rep1.Formulas(4) = ("sthead = '" & disname & "'")
     cry_rep1.Formulas(5) = ("bank = '" & cmb_bank.Text & "'")
     pst_qry = qry1
     cry_rep1.DiscardSavedData = True
     cry_rep1.ReplaceSelectionFormula (pst_qry)
     cry_rep1.WindowState = crptMaximized
     cry_rep1.Connect = gst_repconnect
     cry_rep1.Action = 1
     Exit Sub
 End Sub
