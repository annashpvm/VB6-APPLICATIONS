VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form pay_slip_print_worker 
   Caption         =   "PAYSLIP PRINTING FOR WORKER"
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
      Caption         =   "PAY SLIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   465
      TabIndex        =   4
      Top             =   480
      Width           =   9240
      Begin VB.Frame frame_group 
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
         Height          =   1335
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   6885
         Begin VB.OptionButton opt_pdays 
            Caption         =   " WORKER - for PRESENT DAYS only"
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   4095
         End
         Begin VB.OptionButton opt_layoffdays 
            Caption         =   " WORKER  - for LAYOFF DAYS"
            Height          =   300
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   4395
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   960
         TabIndex        =   11
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Year "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3135
         Left            =   840
         TabIndex        =   5
         Top             =   2640
         Width           =   7335
         Begin VB.Frame Frame6 
            Height          =   2895
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1935
            Begin VB.OptionButton opt_selective_emp 
               Caption         =   "Selective"
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
               TabIndex        =   10
               Top             =   1560
               Width           =   1335
            End
            Begin VB.OptionButton opt_allemp 
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
               Height          =   375
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2775
            Left            =   2160
            TabIndex        =   6
            Top             =   240
            Width           =   5055
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   2310
               Left            =   120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   7
               Top             =   240
               Width           =   4815
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   1080
         Picture         =   "pay_slip_print_worker.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   120
         Picture         =   "pay_slip_print_worker.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   945
      End
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2145
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2865
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "pay_slip_print_worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''Private Sub ExportReportToPDF(ReportObject As CRAXDRT.report, ByVal FileName As String, ByVal ReportTitle As String)
''
''    Dim objExportOptions As CRAXDRT.ExportOptions
''    ReportObject.ReportTitle = ReportTitle
''
''    With ReportObject
''        .EnableParameterPrompting = False
''        .MorePrintEngineErrorMessages = True
''    End With
''
''    Set objExportOptions = ReportObject.ExportOptions
''
''    With objExportOptions
''        .DestinationType = crEDTDiskFile
''        .DiskFileName = FileName
''        .FormatType = crEFTPortableDocFormat
''        .PDFExportAllPages = True
''    End With
''    ReportObject.Export False
''End Sub
''

Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    opt_pdays.Value = True
    If optchk = 2 Then
       frame_group.Visible = False
    Else
       frame_group.Visible = True
    End If
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

Private Sub opt_staff_Click()
   get_emplist
End Sub

Private Sub opt_worker_Click()
   get_emplist
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
   
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip.rpt"
   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   cry_rep1.PrinterSelect
   Dim ds, emp As String
   If optchk = 1 Then
      ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_classification} = 'B'"
   ElseIf optchk = 2 Then
      ds = " and {emp_mas.emp_workplace} <> 'MILL' and {emp_mas.emp_classification} = 'B'"
   ElseIf optchk = 3 Then
      ds = " and {emp_mas.emp_classification} = 'A'"
   Else
      ds = ""
   End If
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
''   If opt_staff.Value = True Then
''      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip_staff.rpt"
''''      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''''                                             "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S'  and {emp_salary.s_salarydays} > 0  " & ds & " and {emp_mas.emp_status}  ='A'")
''''
''      cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'S'  and {emp_salary.s_salarydays} > 0  " & ds & "")
''
''
''   Else
        If opt_pdays = True Then
        cry_rep1.Formulas(5) = ("cond = 0")
''        cry_rep1.Formulas(6) = ("speriod = 'FOR WORKING DAYS'")
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip_worker_for_pdays.rpt"
         
       Else
       cry_rep1.Formulas(5) = ("cond = 1")
         cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip_worker_for_layoff.rpt"
       End If
       cry_rep1.ReplaceSelectionFormula ("{emp_salary.s_year} = " & Val(cmb_year.Text) & " and {emp_salary.s_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                         "and {emp_salary.s_company} = " & company_code & " and {emp_mas.emp_cat} = 'W'  and {emp_salary.s_salarydays} > 0  " & ds & "")
''   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub
Public Sub get_emplist()
''    If opt_staff.Value = True Then
''       If optchk = 1 Then
''           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_classification = 'B'  order by emp_name"
''       Else
''           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_classification = 'A'  order by emp_name"
''       End If
''    Else
       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W'  order by emp_name"
''    End If
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


