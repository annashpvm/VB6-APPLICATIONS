VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_rep_bio_metrics 
   Caption         =   "BIO-METRICS REPORTS"
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
      Caption         =   "BIO-METRIC ATTENDANCE REPORTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   465
      TabIndex        =   3
      Top             =   120
      Width           =   9240
      Begin VB.Frame Frame9 
         Caption         =   "Select Mill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   7335
         Begin VB.OptionButton opt_dpm1 
            Caption         =   "DPM-1"
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
            Height          =   195
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_dpm3 
            Caption         =   "DPM-2"
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
            Height          =   195
            Left            =   3120
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_all 
            Caption         =   "All"
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
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton opt_vjpm 
            Caption         =   "VJPM"
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
            Height          =   195
            Left            =   4680
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opt_cogen 
            Caption         =   "COGEN"
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
            Height          =   195
            Left            =   6000
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   840
         TabIndex        =   18
         Top             =   1560
         Width           =   7335
         Begin VB.ComboBox cmb_rep 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label3 
            Caption         =   "Select Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1665
         End
      End
      Begin VB.Frame frame_group 
         Caption         =   "DETAILS FOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   7365
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
            Height          =   345
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1545
         End
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
            Height          =   345
            Left            =   3720
            TabIndex        =   16
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   840
         TabIndex        =   10
         Top             =   2160
         Width           =   7335
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
            TabIndex        =   12
            Top             =   120
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
            TabIndex        =   11
            Top             =   120
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
            Height          =   210
            Left            =   240
            TabIndex        =   14
            Top             =   120
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
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   840
         TabIndex        =   4
         Top             =   2760
         Width           =   7335
         Begin VB.Frame Frame8 
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3480
            TabIndex        =   21
            Top             =   120
            Width           =   3135
            Begin VB.OptionButton opt_alldept 
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
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton opt_selective_dept 
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
               Height          =   255
               Left            =   1560
               TabIndex        =   22
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Employee  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   3135
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
               Height          =   255
               Left            =   1560
               TabIndex        =   9
               Top             =   240
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
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1935
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   6615
            Begin VB.ListBox lst_dept 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   600
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   24
               Top             =   120
               Width           =   5895
            End
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   6
               Top             =   120
               Width           =   5895
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2880
      TabIndex        =   0
      Top             =   5640
      Width           =   3495
      Begin VB.CommandButton cmd_mail 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Mail"
         Height          =   825
         Left            =   120
         Picture         =   "frm_rep_bio_metrics.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   2520
         Picture         =   "frm_rep_bio_metrics.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   1560
         Picture         =   "frm_rep_bio_metrics.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   945
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_bio_metrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mcode As Integer
Dim emp_type As String

Private Sub cmd_mail_Click()
Dim crapp As New CRAXDRT.Application
Dim report As New CRAXDRT.report
Dim crxExportOptions As CRAXDRT.ExportOptions

Set crapp = New CRAXDRT.Application
Set report = crapp.OpenReport("\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt", 1)

Set crxExportOptions = report.ExportOptions
crxExportOptions.DestinationType = crEDTDiskFile
crxExportOptions.DiskFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.pdf"
crxExportOptions.FormatType = crEFTPortableDocFormat
'crxExportOptions.PDFFirstPageNumber = 1
'crxExportOptions.PDFLastPageNumber = 1
'crxExportOptions.PDFExportAllPages = True
'report.Export False

''With report
''''.ExportOptions.FormatType = crEFTPortableDocFormat
''''.ExportOptions.DestinationType = crEDTDiskFile
''''.ExportOptions.DiskFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.pdf"
''''' location & the file name
''''
''''.ExportOptions.PDFExportAllPages .ExportOptions.PDFExportAllPages
''''.ExportOptions.PDFFirstPageNumber .ExportOptions.PDFFirstPageNumber
''''.ExportOptions.PDFLastPageNumber .ExportOptions.PDFLastPageNumber
''''
''''.Export (False)
''''
''''   .ExportOptions.FormatType = crEFTPortableDocFormat
''''   .ExportOptions.DestinationType = crEDTDiskFile
''''   .ExportOptions.DiskFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.pdf"
''''   .ExportOptions.PDFExportAllPages = True
''''   .Export (False)
''End With
'''no print or display, just export VB6 Crystal 10
'''super simple button click on form
'''report goes straight to disk
''Dim crCrystal As CRAXDRT.Application
''Dim crReport As CRAXDRT.Report
''
''Set crCrystal = New CRAXDRT.Application
''Set crReport = New CRAXDRT.Report
''
''If Dir("\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt") = "" Then
''MsgBox "Error - Report file Not Found! ", vbExclamation
''Exit Sub
''End If
''
''' set rpt = appname.openreport("app path and rpt name"), open method 1=view 2=print
''Set crReport = crCrystal.OpenReport("\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt")
''
''
'''crReport.SQLQueryString = " "
''crReport.DiscardSavedData
''crReport.ExportOptions.DiskFileName = App.path & "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.pdf"
''crReport.ExportOptions.DestinationType = crEDTDiskFile
''crReport.ExportOptions.FormatType = crEFTPortableDocFormat
''crReport.Export False ' do not prompt user = false
''
''Set crCrystal = Nothing
''Set crReport = Nothing
'    Dim objCrystal As CRAXDRT.Application
'    Dim objReport As CRAXDRT.report
'
'    Set objCrystal = New CRAXDRT.Application
'
'    Set objReport = objCrystal.OpenReport("shvpm\vbcryrep\payroll\monthly_attendance_status.rpt", 1)
'    '...code to set report parameters, login information etc...
'
'    ExportReportToPDF objReport, "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt", "Beds Held"

End Sub



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
Private Sub ExportReportToPDF(ReportObject As CRAXDRT.report, ByVal FileName As String, ByVal ReportTitle As String)
   
''    Dim objExportOptions As CRAXDRT.ExportOptions
''
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
''
''    ReportObject.Export False
 
End Sub
 
Private Sub Form_Load()
    mcode = 1
    cmb_rep.AddItem "Monthly Attendence Report"
    cmb_rep.AddItem "Monthly Attendence Report-Abstract"
    cmb_rep.AddItem "Monthly IN/OUT Report"
    cmb_rep.AddItem "Monthly Overtime Report"
    cmb_rep.AddItem "Monthly Late Attendance Report"
    opt_staff.Value = True
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
''    End With
''    cmb_year.Text = "2015"
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    sql = "Select * from  pdept_mas order by dept_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("dept_name")
        lst_dept.ItemData(lst_dept.NewIndex) = payrs("dept_code")
        payrs.MoveNext
    Wend
    payrs.Close
    lst_dept.Visible = False
    emp_type = "S"
    get_emplist
End Sub

Private Sub opt_alldept_Click()
    lst_dept.Enabled = False
    lst_dept.Visible = False
    lst_emp.Visible = True
End Sub

Private Sub opt_allemp_Click()
    lst_emp.Enabled = False
End Sub

Private Sub opt_selective_dept_Click()
    lst_dept.Enabled = True
    lst_dept.Visible = True
    lst_emp.Visible = False
End Sub

Private Sub opt_selective_emp_Click()
    lst_emp.Enabled = True
    lst_dept.Visible = False
    lst_emp.Visible = True
End Sub

Private Sub opt_staff_Click()
   emp_type = "S"
   get_emplist
End Sub


Private Sub opt_worker_Click()
   emp_type = "W"
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
''   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.Text)
   cry_rep1.Formulas(0) = ("report_month = '" & cmb_month.Text & "'")
   cry_rep1.Formulas(1) = ("report_year = '" & cmb_year.Text & "'")
   cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   If opt_staff.Value = True Then
      cry_rep1.Formulas(3) = ("sw= 'STAFF'")
   Else
      cry_rep1.Formulas(3) = ("sw= 'WORKER'")
   End If
   
   cry_rep1.PrinterSelect
   Dim ds, emp, dept As String
''   If optchk = 1 Then
''      ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_classification} = 'B'"
''   ElseIf optchk = 2 Then
''      ds = " and {emp_mas.emp_workplace} <> 'MILL' and {emp_mas.emp_classification} = 'B'"
''   ElseIf optchk = 3 Then
''      ds = " and {emp_mas.emp_classification} = 'A'"
''   Else
''      ds = ""
''   End If
   ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_status} = 'A' "
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
   
   
   dept = ""
   If opt_selective_dept.Value = True Then
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                     dept = " and ( {emp_mas.emp_dept} = " & lst_dept.ItemData(pin_row) & ""
                     i = i + 1
                  Else
                     dept = dept + " or {emp_mas.emp_dept} = " & lst_dept.ItemData(pin_row) & ""
                  End If
               End If
           Next pin_row
        End If
   End If
   If dept <> "" Then dept = dept + ")"
      
   
   
   ds = ds + emp + dept
   If cmb_rep.Text = "Monthly Attendence Report" Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt"
   ElseIf cmb_rep.Text = "Monthly IN/OUT Report" Then
       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_inout_status.rpt"
   End If
   If opt_All.Value = True Then
      cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & "")
      pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & ""
   Else
      cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                             " and {emp_mas.emp_company} = " & mcode & " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & "")
   
   End If
   
''   If opt_staff.Value = True Then
''      cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             "and {emp_mas.emp_company} = " & company_code & " and {emp_mas.emp_cat} = 'S'  " & ds & "")
''   Else
''      cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
''                                             "and {emp_mas.emp_company} = " & company_code & " and {emp_mas.emp_cat} = 'W'  " & ds & "")
''   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub
Public Sub get_emplist()
''    If opt_staff.Value = True Then
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    Else
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    End If
    If opt_All.Value = True Then
       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_cat = '" & emp_type & "' and emp_workplace  = 'MILL' order by emp_name "
    Else
       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A'  and emp_cat = '" & emp_type & "'  and emp_workplace  = 'MILL' order by emp_name  "
    End If
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

Private Sub opt_all_Click()
    mcode = 0
End Sub

Private Sub opt_cogen_Click()
   mcode = 5
End Sub

Private Sub opt_dpm1_Click()
   mcode = 1
End Sub

Private Sub opt_dpm2_Click()
  mcode = 4
End Sub

Private Sub opt_dpm3_Click()
  mcode = 2
End Sub

Private Sub opt_vjpm_Click()
   mcode = 3
End Sub

