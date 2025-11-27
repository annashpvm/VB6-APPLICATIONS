VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_bio_metrics 
   Caption         =   "BIO-METRICS REPORTS"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame12 
      Caption         =   "Frame12"
      Height          =   1095
      Left            =   10320
      TabIndex        =   41
      Top             =   4080
      Width           =   3855
      Begin VB.OptionButton opt_casual 
         Caption         =   "CASUAL"
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
         Left            =   600
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.OptionButton opt_vou 
         Caption         =   "VOUCHER"
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
         Left            =   720
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin VB.Frame Frame11 
      Height          =   2415
      Left            =   9960
      TabIndex        =   37
      Top             =   840
      Width           =   2175
      Begin VB.OptionButton opt_attn_all 
         Caption         =   "ALL ATTENDANCE DETAILS"
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
         Height          =   915
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton opt_attn_present 
         Caption         =   "ONLY PRESENT"
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
         Height          =   555
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame10 
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   6840
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   126746625
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   34
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   126746625
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
   End
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
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   9240
      Begin VB.Frame Frame13 
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
         TabIndex        =   44
         Top             =   840
         Width           =   6885
         Begin VB.OptionButton opt_pfno 
            Caption         =   "NON PF MEMBER"
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
            Left            =   2280
            TabIndex        =   47
            Top             =   240
            Width           =   2025
         End
         Begin VB.OptionButton opt_pfyes 
            Caption         =   "PF MEMBER"
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
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1905
         End
         Begin VB.OptionButton opt_pf_all 
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
            Height          =   345
            Left            =   5040
            TabIndex        =   45
            Top             =   240
            Value           =   -1  'True
            Width           =   1545
         End
      End
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
         Height          =   495
         Left            =   8160
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   15
         Begin VB.OptionButton opt_dpm1 
            Caption         =   "PM-1"
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
            Caption         =   "PM-2"
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
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opt_vjpm 
            Caption         =   "PM3"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   4485
         Begin VB.OptionButton opt_all_employees 
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
            Height          =   345
            Left            =   2760
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   1665
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
            Height          =   345
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1185
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
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   1425
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
               Left            =   1080
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
      Top             =   5760
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
         Visible         =   0   'False
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
Dim mname As String
Dim mcode As Integer
Dim emp_type As String

Private Sub cmb_month_Click()
find_dates
End Sub



Private Sub cmb_year_Click()
find_dates
End Sub

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
'    Set objReport = objCrystal.OpenReport("10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt", 1)
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
   
'    Dim objExportOptions As CRAXDRT.ExportOptions
'
'    ReportObject.ReportTitle = ReportTitle
'
'    With ReportObject
'        .EnableParameterPrompting = False
'        .MorePrintEngineErrorMessages = True
'    End With
'
'    Set objExportOptions = ReportObject.ExportOptions
'
'    With objExportOptions
'        .DestinationType = crEDTDiskFile
'        .DiskFileName = FileName
'        .FormatType = crEFTPortableDocFormat
'        .PDFExportAllPages = True
'    End With
'
    ReportObject.Export False
 
End Sub
 
Private Sub Form_Load()
   emp_type = "A"
   get_emplist
   If casual = "Y" Then
      opt_staff.Visible = False
      opt_worker.Visible = False
      opt_casual.Visible = True
      opt_casual.Value = True
   Else
      opt_staff.Visible = True
      opt_worker.Visible = True
      opt_casual.Visible = True
      opt_staff.Value = True
   End If
   
      opt_staff.Visible = True
      opt_worker.Visible = True
      opt_casual.Visible = False
      opt_casual.Value = False
   
   
    millname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITED "
    
    mcode = 1
    cmb_rep.AddItem "Monthly Attendence Report-Dept.wise"
    cmb_rep.AddItem "Monthly Attendence Report-EmpNo.wise"
    cmb_rep.AddItem "Monthly Attendence Report-Dept.wise-Full Attendance"
    cmb_rep.AddItem "Employee - Monthwise Attendance"
    cmb_rep.AddItem "Muster Roll"
    cmb_rep.AddItem "Worker Cost Report"
''    cmb_rep.AddItem "Monthly Attendence Abstract"
    cmb_rep.AddItem "Monthly IN/OUT Report"
''    cmb_rep.AddItem "Monthly IN/OUT Report-DEPARTMENTWISE"
''    cmb_rep.AddItem "Monthly Overtime Report"
''    cmb_rep.AddItem "Monthly Late Attendance Report"
''    cmb_rep.AddItem "Morethan 10 days Leave+Absent Report"
    cmb_rep.AddItem "MANUAL ATTENDANCE SHEET"
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
    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''        .AddItem "2016"
      .AddItem Left(fyear, 4)
      .AddItem Mid(fyear, 6, 4)
      If Year(Date) = Int(Left(fyear, 4)) Then
         cmb_year.Text = Left(fyear, 4)
      Else
          cmb_year.Text = Mid(fyear, 6, 4)
      End If
    
    End With
    cmb_month.ListIndex = Month(Date) - 1
''    cmb_year.Text = "2015"
    find_dates
    
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
    mname = "SHVPM"

    get_emplist
''    If casual = "Y" Then
''      opt_casual.Visible = True
''      opt_casual.Value = True
''   Else
''      opt_casual.Visible = False
''      opt_staff.Value = True
''   End If
End Sub

Private Sub opt_all_employees_Click()
   emp_type = "A"
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


Private Sub opt_vou_Click()
   emp_type = "R"
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
   
   If cmb_rep.Text = "" Then
      MsgBox ("Select Report ...")
      Exit Sub
   End If
      
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\payslip.rpt"
''   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.Text)
   Cry_rep1.Formulas(0) = ("report_month = '" & cmb_month.Text & "'")
   Cry_rep1.Formulas(1) = ("report_year = '" & cmb_year.Text & "'")
   Cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
   If opt_staff.Value = True Then
      Cry_rep1.Formulas(3) = ("sw= 'STAFF'")
   ElseIf opt_worker.Value = True Then
      Cry_rep1.Formulas(3) = ("sw= 'WORKER'")
   Else
      Cry_rep1.Formulas(3) = ("sw= ''")
   End If
   Cry_rep1.Formulas(4) = ""
   Cry_rep1.Formulas(5) = ""

   
   Cry_rep1.PrinterSelect
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
   
''   If opt_attn_present.Value = True And cmb_rep.Text <> "MANUAL ATTENDANCE SHEET" Then
''        If opt_vou.Value = True Then
''           ds = " and {emp_voupay_mast.emp_workplace} = 'MILL' and {emp_voupay_mast.emp_status} = 'A' and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 "
''        Else
''           If opt_all_employees.Value = True Then
''               ds = " and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0"
''           Else
''               ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_status} = 'A'  and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0"
''           End If
''        End If
''   Else
''        If opt_vou.Value = True Then
''           ds = " and {emp_voupay_mast.emp_workplace} = 'MILL' and {emp_voupay_mast.emp_status} = 'A' "
''        Else
''          ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_status} = 'A' "
''        End If
''   End If
   
   
   If opt_attn_present.Value = True And cmb_rep.Text <> "MANUAL ATTENDANCE SHEET" Then
''        If opt_vou.Value = True Then
''           ds = " and {emp_voupay_mast.emp_workplace} = 'MILL' and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 "
''        Else
''           If opt_all_employees.Value = True Then
''               ds = " and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0"
''           Else
''               ds = " and {emp_mas.emp_workplace} = 'MILL' and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0"
''           End If
''        End If
''
''   Else
''        If opt_vou.Value = True Then
''           ds = " and {emp_voupay_mast.emp_workplace} = 'MILL' "
''        Else
''          ds = " and {emp_mas.emp_workplace} = 'MILL' "
''        End If
       ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0"
   Else
     ds = " and {emp_mas.emp_code} > 110 and {emp_mas.emp_code} < 20000"
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
   
   
   dept = ""
   If opt_all_employees.Value = True Then
        If opt_selective_dept.Value = True Then
        
             i = 0
             If lst_dept.ListCount > 0 Then
                For pin_row = 0 To lst_dept.ListCount - 1
                    If lst_dept.Selected(pin_row) = True Then
                       If i = 0 Then
                          dept = " and ( {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                          i = i + 1
                       Else
                          dept = dept + " or {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                       End If
                    End If
                Next pin_row
             End If
             
''             i = 0
''             If lst_dept.ListCount > 0 Then
''                For pin_row = 0 To lst_dept.ListCount - 1
''                    If lst_dept.Selected(pin_row) = True Then
''                       If i = 0 Then
''                          dept = " and ( {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
''                          i = i + 1
''                       Else
''                          dept = dept + " or {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
''                       End If
''                    End If
''                Next pin_row
''             End If
             
             
             
        End If
   Else
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
   
   End If
   If dept <> "" Then dept = dept + ")"
      
   
   ds = ds + emp + dept
   
   If cmb_rep.Text = "Monthly Attendence Report-EmpNo.wise" Then
      If opt_casual.Value = True Then
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_cs.rpt"
      Else
         If opt_vou.Value = True Then
            Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_vou.rpt"
         Else
            Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status.rpt"
         End If
      End If
   ElseIf cmb_rep.Text = "Monthly Attendence Abstract" Then
           Cry_rep1.Formulas(4) = ("month_end_date = " & Day(end_date.Value) & "")
           If opt_staff.Value = True Then
              Cry_rep1.Formulas(5) = ("formtype= 'FORM 25'")
           ElseIf opt_worker.Value = True Then
               Cry_rep1.Formulas(5) = ("formtype= 'FORM 25B'")
           End If

      If opt_vou.Value = True Then
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_abstract_vou.rpt"
      Else
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_abstract.rpt"
      End If
   ElseIf cmb_rep.Text = "Muster Roll" Then
''           If (opt_pfno.Value = True) Then
''              ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 and {emp_salary.s_pf} = 0"
''           Else
''              ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 and {emp_salary.s_pf} > 0"
''           End If

           If (opt_pfno.Value = True) Then
              ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 "
           Else
              ds = " and {emp_mas.emp_code} > 110  and {emp_mas.emp_code} < 20000   and ({bio_attendlogs.a_present}+{bio_attendlogs.a_layoff}+{bio_attendlogs.a_ml}+{bio_attendlogs.a_ch}+{bio_attendlogs.a_hpe}) > 0 "
           End If
           
           Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
           Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
           
           Cry_rep1.Formulas(2) = ("report_month = '" & cmb_month.Text & "'")
           Cry_rep1.Formulas(3) = ("report_year = '" & cmb_year.Text & "'")
   
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\muster_roll.rpt"
      
   ElseIf cmb_rep.Text = "Monthly Attendence Report-Dept.wise" Then
        If opt_casual.Value = True Then
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_cs.rpt"
        Else
           If opt_staff.Value = True Then
              Cry_rep1.Formulas(5) = ("formtype= 'FORM 25'")
           ElseIf opt_worker.Value = True Then
               Cry_rep1.Formulas(5) = ("formtype= 'FORM 25B'")
           Else
               Cry_rep1.Formulas(5) = ("formtype= ''")
           End If

''            If opt_vou.Value = True Then
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_vou.rpt"
''            ElseIf opt_staff.Value = True Or opt_worker.Value = True Then
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise.rpt"
''            Else
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_all.rpt"
''            End If
               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise.rpt"
        End If
   ElseIf cmb_rep.Text = "Monthly Attendence Report-Dept.wise-Full Attendance" Then
        If opt_casual.Value = True Then
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_cs.rpt"
        Else
           If opt_staff.Value = True Then
              Cry_rep1.Formulas(5) = ("formtype= 'FORM 25'")
           ElseIf opt_worker.Value = True Then
               Cry_rep1.Formulas(5) = ("formtype= 'FORM 25B'")
           Else
               Cry_rep1.Formulas(5) = ("formtype= ''")
           End If

''            If opt_vou.Value = True Then
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_vou.rpt"
''            ElseIf opt_staff.Value = True Or opt_worker.Value = True Then
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise.rpt"
''            Else
''               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_all.rpt"
''            End If
               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise_full.rpt"
        End If
   ElseIf cmb_rep.Text = "Employee - Monthwise Attendance" Then
               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_empwise_abstract.rpt"
   ElseIf cmb_rep.Text = "Monthly IN/OUT Report" Then
           Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
           Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
           Cry_rep1.Formulas(2) = ("millname= '" & millname & "'")
           Cry_rep1.Formulas(3) = ("repname= 'DAILY ATTENDANCE SHEET FROM'")
         

        Cry_rep1.Formulas(4) = ""
        Cry_rep1.Formulas(5) = ""
       
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_Daywise.rpt"
   ElseIf cmb_rep.Text = "MANUAL ATTENDANCE SHEET" Then
       
        Cry_rep1.Formulas(0) = ""
        Cry_rep1.Formulas(1) = ""
        Cry_rep1.Formulas(2) = ""
        If opt_staff.Value = True Then
           Cry_rep1.Formulas(3) = ("sw= 'STAFF'")
        ElseIf opt_worker.Value = True Then
           Cry_rep1.Formulas(3) = ("sw= 'WORKER'")
        ElseIf opt_vou.Value = True Then
           Cry_rep1.Formulas(3) = ("sw= 'RETAINER'")
        Else
           Cry_rep1.Formulas(3) = ("sw= 'CASUAL'")
        End If
        Cry_rep1.Formulas(4) = ""
        Cry_rep1.Formulas(5) = ""
        
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_format.rpt"
        
''        If opt_staff.Value = True Or opt_worker.Value = True Then
''           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_format.rpt"
''        ElseIf opt_vou.Value = True Then
''           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_format_retainer.rpt"
''        ElseIf opt_casual.Value = True Then
''           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_format_casual.rpt"
''        End If
        
   
   ElseIf cmb_rep.Text = "Monthly IN/OUT Report-DEPARTMENTWISE" Then
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_inout_status_deptwise.rpt"
       
   ElseIf cmb_rep.Text = "Morethan 10 days Leave+Absent Report" Then
           If opt_vou.Value = True Then
               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_vou_absent.rpt"
           Else
               Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\monthly_attendance_status_deptwise_absent.rpt"
           End If
   ElseIf cmb_rep.Text = "Worker Cost Report" Then
        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_emp_daily_cost_worker]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_emp_daily_cost_worker] "
        paydb.Execute (pst_qry)
   
        pst_qry = "create view vew_emp_daily_cost_worker as " _
                & " select  dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,sum(w_tot_ot_hrs) as othrs ,s_grosspay/240 as hrwages ,s_netpay  from bio_worker_daily_pihrs a , emp_salary b , " _
                & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = 1 and s_year = 2022 and s_empcode = w_emp_fpcode and s_empcat = w_cat and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "'" _
                & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay"
                
''        pst_qry = "create view vew_emp_daily_cost_worker as " _
''        & " select  dept_name,s_deptcode,emp_name,s_empcode,sum(s_grosspay) as s_grosspay ,sum(othrs) as othrs ,sum(s_netpay) as s_netpay , sum(s_eligible_grosspay) as s_eligible_grosspay from (" _
''        & " select  dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_ot_days *8  as othrs,s_netpay,s_eligible_grosspay  from emp_salary ,  " _
''        & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_year = " & Val(cmb_year.Text) & "  and s_empcat = 'W'  " _
''        & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_ot_days,s_netpay,s_eligible_grosspay  " _
''        & " Union All  " _
''        & " select  dept_name,s_deptcode,emp_name,s_empcode,0 as s_grosspay,sum(w_tot_ot_hrs) as othrs ,0 as s_netpay , 0 as s_eligible_grosspay   from bio_worker_daily_pihrs a , emp_salary b ,  " _
''        & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_year = " & Val(cmb_year.Text) & "  and s_empcode = w_emp_fpcode and s_empcat = w_cat and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "'  " _
''        & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay ) a group by dept_name,s_deptcode,emp_name,s_empcode"
        pst_qry = "create view vew_emp_daily_cost_worker as " _
        & " select  dept_name,s_deptcode,emp_name,s_empcode,sum(s_grosspay) as s_grosspay ,sum(othrs) as othrs ,sum(s_netpay) as s_netpay , sum(s_eligible_grosspay) as s_eligible_grosspay from (" _
        & " select  dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_ot_days *8  as othrs,s_netpay,s_eligible_grosspay  from emp_salary ,  " _
        & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_year = " & Val(cmb_year.Text) & "  and s_empcat = 'W'  " _
        & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_ot_days,s_netpay,s_eligible_grosspay  " _
        & " Union All  " _
        & " select  dept_name,s_deptcode,emp_name,s_empcode,0 as s_grosspay,sum(w_accepted_hrs+w_holiday_ot_hrs) as othrs ,0 as s_netpay , 0 as s_eligible_grosspay   from bio_worker_daily_pihrs a , emp_salary b ,  " _
        & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and s_year = " & Val(cmb_year.Text) & "  and s_empcode = w_emp_fpcode and s_empcat = w_cat and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "'  " _
        & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay ) a group by dept_name,s_deptcode,emp_name,s_empcode"

        paydb.Execute (pst_qry)
           
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_worker_cost.rpt"
       Cry_rep1.ReplaceSelectionFormula ("")
       
        Cry_rep1.Formulas(2) = ""
        Cry_rep1.Formulas(3) = ""
        Cry_rep1.Formulas(4) = ""
        Cry_rep1.Formulas(5) = ""
        Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
        Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1
   Exit Sub
   
   Else
       Exit Sub
   End If
   
   
'''   If cmb_rep.Text = "MANUAL ATTENDANCE SHEET" Then
'''''      If opt_staff.Value = True Or opt_worker.Value = True Then
'''''         Cry_rep1.ReplaceSelectionFormula (" {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & "")
'''''
'''''      ElseIf opt_vou.Value = True Then
'''''          Cry_rep1.ReplaceSelectionFormula (" {emp_voupay_mast.emp_cat} = '" & emp_type & "'  " & ds & "")
'''''      Else
'''''          Cry_rep1.ReplaceSelectionFormula (" {mas_caemp.ca_status} = 'A'")
'''''      End If
'''        Cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_code} > 110")
'''   Else
'''        If opt_casual.Value = True Then
'''             If opt_all.Value = True Then
'''                Cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PVT LTD'")
'''                Cry_rep1.Formulas(2) = ("millname= 'ALL'")
'''                Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & "")
'''                pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & ""
'''             Else
'''                Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {mas_caemp.ca_compcode} = " & mcode & "")
'''                pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & " and {bio_empmas.bioemp_company} = '" & mname & "' and {bio_empmas.bioemp_team} like '%CASUAL%'"
'''             End If
'''        ElseIf opt_vou.Value = True Then
'''             If opt_all.Value = True Then
'''                Cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PVT LTD'")
'''                Cry_rep1.Formulas(2) = ("millname= 'ALL'")
'''                Cry_rep1.ReplaceSelectionFormula (" {emp_voupay_mast.EMP_CLASSIFICATION} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & "")
'''
'''                Cry_rep1.ReplaceSelectionFormula (" {emp_voupay_mast.EMP_WORKPLACE} = 'MILL' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & "")
'''
'''                pst_qry = " {emp_voupay_mast.EMP_WORKPLACE} = 'MILL' and {emp_voupay_mast.EMP_CLASSIFICATION} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & ""
'''             Else
'''                Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {emp_voupay_mast.emp_company} = " & mcode & " " & ds & "")
'''                pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {emp_voupay_mast.emp_company} = " & mcode & " " & ds & ""
'''             End If
'''        ElseIf opt_all.Value = True Then
'''                Cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PVT LTD'")
'''                If opt_all_employees.Value = True Then
'''                    Cry_rep1.ReplaceSelectionFormula (" {bio_empmas.bioemp_status} = 'Working' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                        " and ({bio_attendlogs.a_el}+{bio_attendlogs.a_pl}+{bio_attendlogs.a_absent}) > 1 " & ds & "")
'''
'''                    pst_qry = "{bio_empmas.bioemp_status} = 'Working' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & ds
'''
'''                Else
'''                    Cry_rep1.ReplaceSelectionFormula (" {emp_mas.EMP_CLASSIFICATION} = 'A' and {bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                        " and ({bio_attendlogs.a_el}+{bio_attendlogs.a_pl}+{bio_attendlogs.a_absent}) > 1  and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & "")
'''                     pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & ""
'''
'''                End If
'''
'''                Cry_rep1.ReplaceSelectionFormula pst_qry
'''
'''
'''             Else
'''                If cmb_rep.Text = "Morethan 10 days Leave+Absent Report" Then
'''                   Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and ({bio_attendlogs.a_el}+{bio_attendlogs.a_pl}+{bio_attendlogs.a_absent}) > 9 and {emp_mas.emp_company} = " & mcode & " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & "")
'''                Else
'''                    If opt_pf_all.Value = True Then
'''                       Cry_rep1.Formulas(4) = ("PF= 'ALL'")
'''                    ElseIf opt_pfyes.Value = True Then
'''                       Cry_rep1.Formulas(4) = ("PF= 'PF MEMBERS'")
'''                    ElseIf opt_pfno.Value = True Then
'''                       Cry_rep1.Formulas(4) = ("PF= 'NON PF MEMBERS'")
'''                    End If
'''
'''                   If opt_pf_all.Value = True Then
'''                       Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " ")
'''                   ElseIf opt_pfyes.Value = True Then
'''                       Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " and {emp_mas.EMP_PFELIGIBLE} = 'Y'")
'''
'''                   Else
'''                       Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " and {emp_mas.EMP_PFELIGIBLE} = 'N'")
'''
'''
'''                   End If
'''
'''                        pst_qry = "{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
'''                                                       " and {emp_mas.emp_company} = " & mcode & " and {emp_mas.emp_cat} = '" & emp_type & "'  " & ds & ""
'''                End If
'''             End If

 '''       End If
        
   If opt_pfyes.Value = True Then
      Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " and {emp_mas.EMP_PFELIGIBLE} = 'Y' ")
   
   ElseIf opt_pfno.Value = True Then
      Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & "and {emp_mas.EMP_PFELIGIBLE} = 'N' ")

   Else
     If cmb_rep.Text = "Monthly IN/OUT Report" Then
         Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_sft_hrs} > 0 ")
     ElseIf cmb_rep.Text = "Employee - Monthwise Attendance" Then
         Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {emp_mas.emp_company} = " & mcode & "   " & ds & " ")
     Else
         Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Val(cmb_year.Text) & " and {bio_attendlogs.a_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                         " and {emp_mas.emp_company} = " & mcode & "   " & ds & " ")
     End If
   End If
   


   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1
End Sub
Public Sub get_emplist()
     Dim payrs As New ADODB.Recordset
 
   If emp_type = "A" Then
        lst_dept.Clear
        sql = "select bioemp_dept from bio_empmas group by bioemp_dept order by bioemp_dept"
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF()
              lst_dept.AddItem payrs("bioemp_dept")
              payrs.MoveNext
        Wend
        payrs.Close
   Else
        lst_dept.Clear
        sql = "Select * from  pdept_mas order by dept_name"
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF()
              lst_dept.AddItem payrs("dept_name")
              lst_dept.ItemData(lst_dept.NewIndex) = payrs("dept_code")
              payrs.MoveNext
        Wend
        payrs.Close
    End If
    
    
    If opt_all_employees.Value = True Then
       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_cat = '" & emp_type & "' order by emp_name "
       sql = "Select * from  emp_mas where emp_status = 'A'   order by emp_name "
    Else
       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A'  and emp_cat = '" & emp_type & "'  order by emp_name  "
       sql = "Select * from  emp_mas where emp_company = " & mcode & " and emp_cat = '" & emp_type & "'  order by emp_name  "
       
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
    payrs.Close
End Sub

Private Sub opt_all_Click()
    mcode = 0
    get_emplist
End Sub

Private Sub opt_cogen_Click()
   mcode = 5
   millname = "IIC"
   mname = "IIC"
    get_emplist
End Sub

Private Sub opt_dpm1_Click()
   mcode = 1
   millname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITED "
   mname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITED "
    get_emplist

End Sub

Private Sub opt_dpm2_Click()
  mcode = 4
End Sub

Private Sub opt_dpm3_Click()
  mcode = 2
  millname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITED"
  mname = "SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITED"
   get_emplist
End Sub

Private Sub opt_vjpm_Click()
   mcode = 3
   millname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD"
   mname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD"
    get_emplist
End Sub

Public Sub find_dates()
    If cmb_year.Text = "" Then Exit Sub
    If cmb_month.ListIndex = -1 Then Exit Sub
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
