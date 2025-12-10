VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rep_bio_metrics_daily 
   Caption         =   "BIOMETRIC DAILY REPORTS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3840
      TabIndex        =   23
      Top             =   7320
      Width           =   1935
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   0
         Picture         =   "frm_rep_bio_metrics_daily.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   960
         Picture         =   "frm_rep_bio_metrics_daily.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BIOMETRIC DAILY REPORTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.Frame Frame10 
         Height          =   1455
         Left            =   5880
         TabIndex        =   32
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opt_shiftG 
            Caption         =   "G Shift"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_shiftC 
            Caption         =   "C Shift"
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
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton opt_shiftB 
            Caption         =   "B Shift"
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
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton opt_shiftA 
            Caption         =   "A Shift"
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
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opt_all_shift 
            Caption         =   "All Shift"
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
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   360
         TabIndex        =   26
         Top             =   5760
         Width           =   7335
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61800449
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   5520
            TabIndex        =   28
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61800449
            CurrentDate     =   39359
         End
         Begin VB.Label Label2 
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
            Left            =   4200
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   7335
         Begin VB.Frame Frame7 
            Height          =   1935
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   6615
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   22
               Top             =   120
               Width           =   5895
            End
            Begin VB.ListBox lst_dept 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   600
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   21
               Top             =   120
               Width           =   5895
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
            TabIndex        =   17
            Top             =   120
            Width           =   3135
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
               TabIndex        =   19
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
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
               TabIndex        =   18
               Top             =   240
               Width           =   1335
            End
         End
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
            TabIndex        =   14
            Top             =   120
            Width           =   3135
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
               TabIndex        =   16
               Top             =   240
               Width           =   1335
            End
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
               TabIndex        =   15
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   360
         TabIndex        =   10
         Top             =   2040
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
            TabIndex        =   11
            Top             =   240
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
            TabIndex        =   12
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
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   5445
         Begin VB.OptionButton opt_allcat 
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
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   945
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
            Left            =   3120
            TabIndex        =   9
            Top             =   240
            Width           =   1545
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
            Left            =   1560
            TabIndex        =   8
            Top             =   240
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
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   5415
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
            Left            =   4320
            TabIndex        =   6
            Top             =   240
            Width           =   975
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
            Left            =   3240
            TabIndex        =   5
            Top             =   240
            Width           =   855
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
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   615
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
            Left            =   2040
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
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
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_rep_bio_metrics_daily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mname As String
Dim mcode, sft, mill As String
Dim emp_type As String
Private Sub exit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    mill = ""
    mcode = ""
    mname = "DPM/VJPM/COGEN"
    cmb_rep.AddItem "Daily Status Report"
    opt_allcat.Value = True
    
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    sql = "select emp_dept  from bio_empmas group by emp_dept order by emp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("emp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    lst_dept.Visible = False
''    emp_type = "S"
    get_emplist
    st_date.Value = Now
    end_date.Value = Now
    cmb_rep.Text = "Daily Status Report"
End Sub
Public Sub get_emplist()
''    If opt_staff.Value = True Then
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    Else
''       sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_mas.emp_workplace = 'MILL'  order by emp_name"
''    End If
''If opt_allcat.Value = True Then
''    If opt_all.Value = True Then
''       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_workplace  = 'MILL' order by emp_name "
''    Else
''       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A' and emp_workplace  = 'MILL' order by emp_name  "
''    End If
''Else
''    If opt_all.Value = True Then
''       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_cat = '" & emp_type & "' and emp_workplace  = 'MILL' order by emp_name "
''    Else
''       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A'  and emp_cat = '" & emp_type & "'  and emp_workplace  = 'MILL' order by emp_name  "
''    End If
''End If
    
   If opt_allcat.Value = True Then
        If opt_All.Value = True Then
           sql = "Select * from  bio_empmas where emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where emp_company = '" & mill & "' and emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        End If
   ElseIf opt_staff.Value = True Then
        If opt_All.Value = True Then
           sql = "Select * from  bio_empmas where emp_team = 'STAFF' and emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where emp_team =  'STAFF' and emp_company = '" & mill & "' and emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        End If
   Else
        If opt_All.Value = True Then
           sql = "Select * from  bio_empmas where emp_team <> 'STAFF' and emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where emp_team <>  'STAFF' and emp_company = '" & mill & "' and emp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        End If
   
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

Private Sub opt_all_shift_Click()
  sft = ""
End Sub

''Private Sub opt_allcat_Click()
'' emp_type = "S or {emp_mas.emp_cat} = W "
''''   get_emplist
''End Sub

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

Private Sub opt_shiftA_Click()
  sft = " and {bio_device_shiftlogs.ds_shift} = 'A SHIFT' "
End Sub

Private Sub opt_shiftB_Click()
  sft = " and {bio_device_shiftlogs.ds_shift} = 'B SHIFT' "
End Sub

Private Sub opt_shiftC_Click()
  sft = " and {bio_device_shiftlogs.ds_shift} = 'C SHIFT' "

End Sub

Private Sub opt_shiftG_Click()
  sft = " and {bio_device_shiftlogs.ds_shift} = 'GS'"
End Sub

Private Sub opt_staff_Click()
    emp_type = "STAFF"
    get_emplist
End Sub

Private Sub opt_worker_Click()
    emp_type = "W"
   get_emplist
End Sub
Private Sub opt_all_Click()
    mill = ""
    mcode = ""
    mname = "DPM/VJPM/COGEN"
End Sub

Private Sub opt_cogen_Click()
   mill = "COGEN"
   mcode = " and {bio_empmas.emp_company} = 'COGEN' "
   mname = "COGEN"

End Sub

Private Sub opt_dpm1_Click()
   mill = "DPM"
   mcode = " and {bio_empmas.emp_company} = 'DPM' "
   mname = "DPM"
End Sub

Private Sub opt_dpm2_Click()
  mcode = 4
End Sub

Private Sub opt_dpm3_Click()
  mill = "DPM 2"
  mcode = " and {bio_empmas.emp_company} = 'DPM 2'"
  mname = "DPM 2"

End Sub

Private Sub opt_vjpm_Click()
   mill = "VJPM"
   mcode = " and {bio_empmas.emp_company} = 'VJPM' "
   mname = "VJPM"
   
End Sub


Private Sub Option1_Click()
  sft = ""

End Sub

Private Sub PROCESS_Click()
   Dim date1 As Date
   date1 = 1 & "/" & 1 & " /" & 1900
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   
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
''   ds = " and {emp_mas.emp_workplace} = 'MILL' and {emp_mas.emp_status} = 'A' "
   emp = ""
   If opt_selective_emp.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_emp.ListCount > 0 Then
           For pin_row = 0 To lst_emp.ListCount - 1
               If lst_emp.Selected(pin_row) = True Then
                  If i = 0 Then
                     emp = " and ( {bio_empmas.emp_name} = '" & lst_emp.List(pin_row) & "'"
                     i = i + 1
                  Else
                     emp = emp + " or {bio_empmas.emp_name} = '" & lst_emp.List(pin_row) & "'"
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
                     dept = " and ( {bio_empmas.emp_dept} = " & lst_dept.List(pin_row) & ""
                     i = i + 1
                  Else
                     dept = dept + " or {bio_empmas.emp_dept} = " & lst_dept.List(pin_row) & ""
                  End If
               End If
           Next pin_row
        End If
   End If
   If dept <> "" Then dept = dept + ")"
   ds = ds + emp + dept
   If cmb_rep.Text = "Daily Status Report" Then
      cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_report.rpt"
   End If

If opt_allcat.Value = True Then
     If opt_All.Value = True Then
        cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_status}<> 'P(OD)'  " & ds & " ")
        pst_qry = "{bio_attendlogs_daily.a_date} >= date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & " and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")"
     Else
        cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.emp_company} = " & mcode & " ")

     End If
Else
    If opt_All.Value = True Then
    cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.emp_cat} = '" & emp_type & "'")
        pst_qry = "{bio_attendlogs_daily.a_date} >= date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & " and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")"
    Else
     cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {emp_mas.emp_company} = " & mcode & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.emp_cat} = '" & emp_type & "' ")
    End If

End If
   If opt_allcat.Value = True Then
      cry_rep1.ReplaceSelectionFormula (" isnull({bio_device_shiftlogs.ds_shift_in}) = false and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") " & sft & mcode)
   ElseIf opt_staff.Value = True Then
      cry_rep1.ReplaceSelectionFormula (" isnull({bio_device_shiftlogs.ds_shift_in}) = false and  {bio_empmas.emp_team} = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") " & sft & mcode)
   Else
      cry_rep1.ReplaceSelectionFormula (" isnull({bio_device_shiftlogs.ds_shift_in}) = false and  {bio_empmas.emp_team} <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") " & sft & mcode)
   End If
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub
