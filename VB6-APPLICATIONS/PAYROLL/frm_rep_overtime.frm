VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_overtime 
   Caption         =   "OVER TIME REPORTS"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   13125
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   10200
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   154075137
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   154075137
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "OVER TIME REPORT"
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
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   9240
      Begin VB.OptionButton opt_abstract 
         Caption         =   "MONTH ABSTRACT OF  OVER TIME DETAILS"
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
         Left            =   840
         TabIndex        =   10
         Top             =   2520
         Width           =   7095
      End
      Begin VB.OptionButton opt_datewise 
         Caption         =   "DAY WISE OVER TIME DETAILS ABSTRACT"
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
         Left            =   840
         TabIndex        =   9
         Top             =   1920
         Value           =   -1  'True
         Width           =   5895
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3840
         TabIndex        =   6
         Top             =   4320
         Width           =   1695
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "frm_rep_overtime.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "frm_rep_overtime.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   720
         TabIndex        =   1
         Top             =   840
         Width           =   7455
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
            TabIndex        =   3
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
            Left            =   5160
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "MONTH"
            Height          =   330
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "YEAR"
            Height          =   285
            Left            =   4080
            TabIndex        =   4
            Top             =   240
            Width           =   885
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   240
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_month_Change()
find_dates
End Sub

Private Sub cmb_month_Click()
find_dates
End Sub

Private Sub cmb_year_Change()
find_dates
End Sub

Private Sub cmb_year_Click()
find_dates
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
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
End Sub

Private Sub PROCESS_Click()
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
''   cry_rep1.Formulas(0) = ("report_month = " & cmb_month.ItemData(cmb_month.ListIndex))
''   cry_rep1.Formulas(1) = ("report_year = " & Val(cmb_year.Text))
   cry_rep1.Formulas(1) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(1) = ("millname= '" & compname & "'")
   cry_rep1.Formulas(0) = ("rmonth = '" & cmb_month.Text & "-" & cmb_year.Text & "'")
   
   If opt_abstract.Value = True Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\overtime_wages.rpt"
      pst_qry = "{overtime_entry.ot_year} = " & Val(cmb_year.Text) & " and {overtime_entry.ot_month}= " & cmb_month.ItemData(cmb_month.ListIndex) & _
                                               " and {overtime_entry.ot_company} = " & company_code & " "
   Else
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\daily_pihrs.rpt"
      pst_qry = "{bio_worker_daily_pihrs.w_date} >= " & " date(" & Format$(st_date, "yyyy,mm,dd") & ")  and {bio_worker_daily_pihrs.w_date} <= " & " date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_worker_daily_pihrs.w_company} = " & company_code & " "
   
''
''            pst_qry = "{emp_voupay_mast.emp_resigneddate} >= " & " date(" & Format$(dt_from, "yyyy,mm,dd") & ") AND " _
''                  & " {emp_voupay_mast.emp_resigneddate} <= " & " date(" & Format$(dt_to, "yyyy,mm,dd") & ") "
''
   End If
   cry_rep1.ReplaceSelectionFormula (pst_qry)
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
   Exit Sub

End Sub


Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
    If cmb_year.Text = "" Then Exit Sub
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
    st_date = end_date - Day(end_date - 1)
End Sub


