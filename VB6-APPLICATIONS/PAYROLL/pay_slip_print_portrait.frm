VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form pay_slip_print_portrait 
   Caption         =   "Payslip Printing Portrait"
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
      Left            =   1200
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3960
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   1080
         Picture         =   "pay_slip_print_portrait.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   120
         Picture         =   "pay_slip_print_portrait.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
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
      Height          =   5475
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   8640
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   6975
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
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
            Left            =   5400
            TabIndex        =   8
            Top             =   240
            Width           =   1335
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
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Year "
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
            Height          =   255
            Left            =   4440
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3735
         Left            =   840
         TabIndex        =   1
         Top             =   1320
         Width           =   7335
         Begin VB.Frame Frame6 
            Height          =   1935
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1935
            Begin VB.OptionButton opt_selective_emp 
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
               TabIndex        =   6
               Top             =   1080
               Width           =   1335
            End
            Begin VB.OptionButton opt_allemp 
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
               TabIndex        =   5
               Top             =   480
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Height          =   3255
            Left            =   2160
            TabIndex        =   2
            Top             =   120
            Width           =   5055
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   2760
               Left            =   120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   3
               Top             =   240
               Width           =   4815
            End
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   "PAYSLIP PRINTING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "pay_slip_print_portrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
''    If optchk = 2 Then
''       frame_group.Visible = False
''    Else
''       frame_group.Visible = True
''    End If
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
    get_emplist
End Sub

Private Sub opt_allemp_Click()
    lst_emp.Enabled = False
End Sub

Private Sub opt_selective_emp_Click()
    lst_emp.Enabled = True
End Sub

Private Sub PROCESS_Click()
    If cmb_month.Text = "" Then
      MsgBox ("Select Month ...")
      Exit Sub
   End If
   If cmb_year.Text = "" Then
      MsgBox ("Select Year ...")
      Exit Sub
   End If
   
   millname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD"
   
    MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip.rpt"
   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.PrinterSelect
   Dim ds, emp As String
   

   
''   If optchk = 1 Then
''      ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_classification} = 'B'"
''      ds = " and {emp_mas.emp_classification} = 'B'"
''   ElseIf optchk = 2 Then
''      ds = " and {emp_mas.emp_workplace} <> 'MILL' and {emp_mas.emp_classification} = 'B'"
''      ds = " and {emp_mas.emp_classification} = 'B'"
''   ElseIf optchk = 3 Then
''      ds = " and {emp_mas.emp_classification} = 'A'"
''   Else
''      ds = ""
''   End If
   emp = ""
   If opt_selective_emp.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_emp.ListCount > 0 Then
           For pin_row = 0 To lst_emp.ListCount - 1
               If lst_emp.Selected(pin_row) = True Then
                  If i = 0 Then
                     emp = " and ( {emp_mas.emp_name} = '" & lst_emp.List(pin_row) & "'"
                     i = i + 1
                  Else
                     emp = emp + " or {emp_mas.emp_name} = '" & lst_emp.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If emp <> "" Then emp = emp + ")"
   ds = ds + emp
   
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip.rpt"

      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             "and {emp_salary.s_company} = " & company_code & "  and {emp_salary.s_salarydays} > 0  " & ds & "")
                                             
   
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub
Public Sub get_emplist()
''
''       If optchk = 1 Then
''           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_classification = 'B'  order by emp_name"
''       Else
''           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_classification = 'A'  order by emp_name"
''       End If

    sql = "select * from emp_mas  where emp_company = '" & company_code & "' order by emp_name"

    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_emp.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_emp.AddItem payrs("emp_name")
        payrs.MoveNext
    Wend
End Sub

