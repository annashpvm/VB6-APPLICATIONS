VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_bio_metrics_daily 
   Caption         =   "BIOMETRIC DAILY REPORTS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   1335
      Left            =   8880
      TabIndex        =   29
      Top             =   3600
      Width           =   2175
      Begin VB.OptionButton opt_de_select_all 
         Caption         =   "De-Select All"
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
         TabIndex        =   31
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton opt_select_all 
         Caption         =   "Select All"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
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
      TabIndex        =   13
      Top             =   7200
      Width           =   1935
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   0
         Picture         =   "frm_rep_bio_metrics_daily.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   14
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8295
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   360
         TabIndex        =   32
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
            TabIndex        =   39
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
               TabIndex        =   41
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
               TabIndex        =   40
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
            TabIndex        =   36
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
               TabIndex        =   38
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
               TabIndex        =   37
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1935
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   6615
            Begin VB.ListBox lst_dept 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   600
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   35
               Top             =   120
               Width           =   5895
            End
            Begin VB.ListBox lst_emp 
               Enabled         =   0   'False
               Height          =   1635
               Left            =   120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   34
               Top             =   120
               Width           =   5895
            End
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1815
         Left            =   5880
         TabIndex        =   22
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton opt_shiftA_B_C 
            Caption         =   "A + B + C Shift"
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
            TabIndex        =   28
            Top             =   1440
            Width           =   1815
         End
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
            TabIndex        =   27
            Top             =   1200
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
            TabIndex        =   26
            Top             =   960
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
            TabIndex        =   25
            Top             =   720
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
            TabIndex        =   24
            Top             =   480
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
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   360
         TabIndex        =   16
         Top             =   5760
         Width           =   7335
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   2400
            TabIndex        =   17
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   131334145
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   5520
            TabIndex        =   18
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   131334145
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   360
            Width           =   1935
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
            TabIndex        =   11
            Text            =   "cmb_rep"
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
            TabIndex        =   21
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
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   255
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
    mname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD"
    ''cmb_rep.AddItem "Daily Status Report"
    cmb_rep.AddItem "Daily IN / OUT punch"
    cmb_rep.AddItem "Daily IN / OUT punch - Employee/Date wise"
    cmb_rep.AddItem "Daily Status Report-Deptwise"
    cmb_rep.AddItem "Daily Status Report-Deptwise-Abstract"
''    cmb_rep.AddItem "Daily Attendance Status Report"
    cmb_rep.AddItem "Daily Absent Report"
''    cmb_rep.AddItem "Attendance Status for the period"

    cmb_rep.AddItem "Daily IN/OUT > 12 HOURS WORKED & < 3 HOURS WORKED"
    cmb_rep.AddItem "Daily IN/OUT > 24 HOURS WORKED"
    cmb_rep.AddItem "Manual Punch Report"
    cmb_rep.AddItem "Miss IN / OUT punch Report"
    cmb_rep.AddItem "Daily Cost Report"
    cmb_rep.AddItem "Worker Cost Report"
    cmb_rep.AddItem "Leave / Absent Abstract"
    cmb_rep.AddItem "Late Attendance Report"
    cmb_rep.AddItem "Absent Details"
    cmb_rep.AddItem "Doubtful Punches"
    opt_allcat.Value = True
    
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    sql = "select bioemp_dept  from bio_empmas group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    lst_dept.Visible = False
''    emp_type = "S"
    get_emplist
    st_date.Value = Now
    end_date.Value = Now
    end_date.MaxDate = Now
    st_date.MaxDate = Now
    cmb_rep.Text = "Daily IN / OUT punch"
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
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where bioemp_company = '" & mill & "' and bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
        End If
   ElseIf opt_staff.Value = True Then
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_team = 'STAFF' and bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where bioemp_team =  'STAFF' and bioemp_company = '" & mill & "' and bioemp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        End If
   Else
        If opt_all.Value = True Then
           sql = "Select * from  bio_empmas where bioemp_team <> 'STAFF' and bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
        Else
           sql = "Select * from  bio_empmas where bioemp_team <>  'STAFF' and bioemp_company = '" & mill & "' and bioemp_status = 'Working'  and emp_dept = '" & lst_dept.Text & "'"
        End If
   
   End If
 sql = "Select * from  bio_empmas where bioemp_status = 'Working'"
   
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_emp.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_emp.AddItem payrs("bioemp_name")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub opt_all_shift_Click()
  sft = ""
End Sub

''Private Sub opt_allcat_Click()
'' emp_type = "S or {emp_mas.emp_cat} = W "
''''   get_emplist
''End Sub

Private Sub opt_alldept_Click()
    lst_dept.Enabled = True
    lst_dept.Visible = True
    lst_emp.Visible = False
End Sub

Private Sub opt_allemp_Click()
    lst_emp.Enabled = False
End Sub

Private Sub opt_de_select_all_Click()
     For i = 1 To lst_dept.ListCount() - 1
        lst_dept.Selected(i) = False
    Next
End Sub

Private Sub opt_select_all_Click()
     For i = 0 To lst_dept.ListCount() - 1
        lst_dept.Selected(i) = True
    Next
    
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

Private Sub opt_shiftA_B_C_Click()
  sft = " and {bio_device_shiftlogs.ds_shift_actual} = 'A+B+C'"
End Sub

Private Sub opt_shiftA_Click()
  sft = " and {bio_device_shiftlogs.ds_shift_actual} = 'A SHIFT' "
End Sub

Private Sub opt_shiftB_Click()
  sft = " and {bio_device_shiftlogs.ds_shift_actual} = 'B SHIFT' "
End Sub

Private Sub opt_shiftC_Click()
  sft = " and {bio_device_shiftlogs.ds_shift_actual} = 'C SHIFT' "

End Sub

Private Sub opt_shiftG_Click()
  sft = " and {bio_device_shiftlogs.ds_shift_actual} = 'GS' "
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
    mname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD'"
End Sub




Private Sub Option1_Click()
  sft = ""

End Sub

Private Sub PROCESS_Click()
''
''   Dim prt As Printer
''   For Each prt In Printers
''       If prt.DeviceName = "doPDF v7" Then
''          Set Printer = prt
''          Exit For
''       End If
''       If prt.DeviceName = "PDFCreator" Then
''          Set Printer = prt
''          Exit For
''       End If
''   Next
''
''    Printer.FontSize = 10
''    Printer.FontName = "Courier New"
''    Printer.PaperSize = vbPRPSA4 ' vbPRPSLegal  vbPRPS11x17 vbPRPSA3
''    Printer.Orientation = vbPRORPortrait 'vbPRORLandscape

   Dim date1 As Date
   date1 = 1 & "/" & 1 & " /" & 1900
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   Cry_rep1.Formulas(3) = ""
   
   
   Cry_rep1.ParameterFields(0) = vbNullString
   Cry_rep1.ParameterFields(1) = vbNullString

   Cry_rep1.PrinterSelect
   Dim ds, emp, dept As String

   ds = " and {bio_empmas.bioemp_fpcode} > 110 and {bio_empmas.bioemp_fpcode} < 20000  "


   emp = ""
   If opt_selective_emp.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_emp.ListCount > 0 Then
           For pin_row = 0 To lst_emp.ListCount - 1
               If lst_emp.Selected(pin_row) = True Then
                  If i = 0 Then
                     emp = " and ( {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
                     i = i + 1
                  Else
                     emp = emp + " or {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If emp <> "" Then emp = emp + ")"
   
   
   dept = ""
   If cmb_rep.Text = "Leave / Absent Abstract" Or cmb_rep.Text = "Absent Details" Then
        ds = " {vew_attn.fpcode}> 110 and  {vew_attn.fpcode} < 20000  "
        If opt_selective_dept.Value = True Then
             i = 0
             If lst_dept.ListCount > 0 Then
                For pin_row = 0 To lst_dept.ListCount - 1
                    If lst_dept.Selected(pin_row) = True Then
                          If i = 0 Then
                            dept = " and ( {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                          Else
                            dept = dept + " or {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                          End If
                    End If
                Next pin_row
             End If
        End If
   Else
        If opt_selective_dept.Value = True Then
             i = 0
             If lst_dept.ListCount > 0 Then
                For pin_row = 0 To lst_dept.ListCount - 1
                    If lst_dept.Selected(pin_row) = True Then
                       If i = 0 Then
                          If cmb_rep.Text = "Miss IN / OUT punch Report" Then
                              dept = " and ( {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                          Else
                              dept = " and ( {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                          End If
                          i = i + 1
                       Else
                          If cmb_rep.Text = "Miss IN / OUT punch Report" Then
                              dept = dept + " or {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                          Else
                              dept = dept + " or {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                          End If
                       End If
                    End If
                Next pin_row
             End If
        End If
   
   End If
   
   
   
   If dept <> "" Then dept = dept + ")"
   ds = ds + emp + dept
   
''   If cmb_rep.Text = "Daily Status Report" Then
''      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_report.rpt"
''   ElseIf cmb_rep.Text = "Daily Absent Report" Then
''      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_absent.rpt"
''   ElseIf cmb_rep.Text = "Attendance Status for the period" Then
''      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_status_for_period.rpt"
''
''   End If

''If opt_allcat.Value = True Then
''     If opt_all.Value = True Then
''        Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status}<> 'P(OD)' AND {bio_device_shiftlogs.ds_status}<> 'P'  " & ds & "")
''        Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status} = 'P' " & ds & "")
''        pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status}<> 'P(OD)'  "
''     Else
''        Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status}<> 'P(OD)'  and {emp_mas.bioemp_company} = " & mcode & " ")
''''        Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.bioemp_company} = " & mcode & " ")
''
''     End If
''Else
''    If opt_all.Value = True Then
''''       Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.emp_cat} = '" & emp_type & "'")
''       Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status}= 'P'  and {emp_mas.emp_cat} = '" & emp_type & "'")
''    Else
''       Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs_daily.a_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_attendlogs_daily.a_date_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " and {emp_mas.emp_company} = " & mcode & " and {bio_attendlogs_daily.a_status}<> 'P(OD)' and {emp_mas.emp_cat} = '" & emp_type & "' ")
''       Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status}<> 'P(OD)'  and {emp_mas.emp_cat} = '" & emp_type & "' and {emp_mas.bioemp_company} = " & mcode & " ")
''    End If
''
''End If
   If cmb_rep.Text = "Daily Status Report" Or cmb_rep.Text = "Daily Status Report-Deptwise" Or cmb_rep.Text = "Daily Status Report-Deptwise-Abstract" Then
        If cmb_rep.Text = "Daily Status Report" Then
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_report.rpt"
        ElseIf cmb_rep.Text = "Daily Status Report-Deptwise" Then
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_report_deptwise.rpt"
        Else
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_report_deptwise_abs.rpt"
        End If
        sft = ""
''        If opt_shiftA.Value = True Then
''           sft = " and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '05:40'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '06:10' "
''        ElseIf opt_shiftB.Value = True Then
''           sft = " and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '13:40'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '14:10' "
''        ElseIf opt_shiftC.Value = True Then
''           sft = " and ds_date = '11/16/2018' and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '21:00'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '22:20' "
''        End If
        
        If opt_shiftA.Value = True Then
           sft = "and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"") >= '05:00'  and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"")  <= '06:20' "
           Cry_rep1.Formulas(3) = ("shift_name = 'A SHIFT'")
        ElseIf opt_shiftB.Value = True Then
           sft = " and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '13:40'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '14:10' "
           sft = "and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"") >= '13:00'  and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"")  <= '14:20' "
           Cry_rep1.Formulas(3) = ("shift_name = 'B SHIFT'")
        ElseIf opt_shiftC.Value = True Then
        
''           sft = " and ds_date = '11/16/2018' and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '21:00'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '22:20' "
''           sft = "and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"") >= '21:00'  and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"")  <= '22:10' "
''           sft = "and ({bio_device_shiftlogs.ds_shift}= '6.00PM-6.00AM' or {bio_device_shiftlogs.ds_shift}= '8.00 PM to 8.00 AM' or {bio_device_shiftlogs.ds_shift}= 'C Shift') "
           sft = "and ({bio_device_shiftlogs.ds_shift_actual}= '08.00PM-08.00AM' or {bio_device_shiftlogs.ds_shift_actual}= '06.00PM-06.00AM' or  {bio_device_shiftlogs.ds_shift_actual}= '07.00PM-07.00AM' or {bio_device_shiftlogs.ds_shift_actual}= '08.00 PM to 08.00 AM' or {bio_device_shiftlogs.ds_shift_actual}= 'C Shift') "
           Cry_rep1.Formulas(3) = ("shift_name = 'C SHIFT , 6 PM - 6 AM , 7 PM - 7 AM  & 8 PM to 8 AM '")
        ElseIf opt_shiftG.Value = True Then
           sft = " and ds_date = '11/16/2018' and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) >= '21:00'  and CONVERT(VARCHAR(8),{bio_device_shiftlogs.ds_shift_in},114) <= '22:20' "
           sft = "and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"") >= '07:00'  and totext({bio_device_shiftlogs.ds_shift_in},""HH:mm"")  <= '08:50' "
           Cry_rep1.Formulas(3) = ("shift_name = 'GENERAL SHIFT'")
        ElseIf opt_shiftA_B_C.Value = True Then
           Cry_rep1.Formulas(3) = ("shift_name = 'A+B+C SHIFT'")
            sft = "and ({bio_device_shiftlogs.ds_shift}= 'A+B+C')"
        Else
        
        End If
        
        If opt_allcat.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode)
           
           If cmb_rep.Text = "Daily Status Report-Deptwise-Abstract" Then
              Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and ({bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) or {bio_device_shiftlogs.ds_status} = 'L') " & sft & mcode & ds)
           Else
              Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode & ds)
           End If

           pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode & ds
        ElseIf opt_staff.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("left({bio_empmas.bioemp_team},5) = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode)
        Else
           Cry_rep1.ReplaceSelectionFormula (" left({bio_empmas.bioemp_team},5) <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode)
        End If
        
   
   ElseIf cmb_rep.Text = "Daily Absent Report" Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_absent.rpt"
           
        If opt_allcat.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and  {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} = Date (1900,01,01) and {bio_device_shiftlogs.ds_status} <> 'WO' " & sft & mcode & ds)
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and  {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status} = 'A' " & sft & mcode & ds)
        ElseIf opt_staff.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and {bio_empmas.bioemp_team} = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} = Date (1900,01,01)  and {bio_device_shiftlogs.ds_status} <> 'WO'" & sft & mcode & ds)
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and {bio_empmas.bioemp_team} = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_status} <> 'A'" & sft & mcode & ds)
          
        Else
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and {bio_empmas.bioemp_team} <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} = Date (1900,01,01) and {bio_device_shiftlogs.ds_status} <> 'WO' " & sft & mcode & ds)
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_status} = 'Working' and {bio_empmas.bioemp_team} <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_status} <> 'A'" & sft & mcode & ds)
        
        End If

   ElseIf cmb_rep.Text = "Attendance Status for the period" Then
        Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_status_for_period.rpt"
        
''        Cry_rep1.ParameterFields(0) = "fdate;" & Format(st_date.Value, "MM/dd/yyyy") & ";true"
''        Cry_rep1.ParameterFields(1) = "sdate;" & Format(end_date.Value, "MM/dd/yyyy") & ";true"
   

        ''Cry_rep1.ParameterFields(0) = "fdate;date(" & Year(st_date) & "," & Month(st_date) & "," & Day(st_date) & ")";true"
'   ''     Cry_rep1.ParameterFields(1) = "sdate;" & Format(end_date.Value, "MM/dd/yyyy") & ";true"


        If opt_allcat.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.a_year} = " & Year(st_date.Value) & " and {bio_attendlogs.a_month}= " & Month(st_date.Value) & "")

''           Cry_rep1.ReplaceSelectionFormula ("{bio_attendlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") " & sft & mcode)
        ElseIf opt_staff.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_team} = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & sft & mcode)
        Else
           Cry_rep1.ReplaceSelectionFormula (" {bio_empmas.bioemp_team} <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")" & sft & mcode)
        End If


   ElseIf cmb_rep.Text = "Daily IN / OUT punch" Or cmb_rep.Text = "Daily IN / OUT punch - Employee/Date wise" Or cmb_rep.Text = "Daily IN/OUT > 12 HOURS WORKED & < 3 HOURS WORKED" Or cmb_rep.Text = "Daily IN/OUT > 24 HOURS WORKED" Then
        Dim io_opt As String
        
        Cry_rep1.Formulas(3) = ("repname = '" & cmb_rep.Text & "'")
        If cmb_rep.Text = "Daily IN / OUT punch - Employee/Date wise" Then
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_in_out_punch_empwise.rpt"
        Else
           Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_in_out_punch_report.rpt"
        End If
''        If opt_allcat.Value = True Then
''           Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode)
''           pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode
''        ElseIf opt_staff.Value = True Then
''           Cry_rep1.ReplaceSelectionFormula ("left({bio_empmas.bioemp_team},5) = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode)
''        Else
''           Cry_rep1.ReplaceSelectionFormula (" left({bio_empmas.bioemp_team},5) <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode)
''        End If
        io_opt = ""
        If cmb_rep.Text = "Daily IN/OUT > 12 HOURS WORKED & < 3 HOURS WORKED" Then
            io_opt = "and ({bio_device_shiftlogs.ds_sft_hrs} > 12 or ({bio_device_shiftlogs.ds_sft_hrs} < 3 and {bio_device_shiftlogs.ds_shift_in} <> DATE(1900,01,01))) "
        End If
        If cmb_rep.Text = "Daily IN/OUT > 24 HOURS WORKED" Then
            io_opt = "and ({bio_device_shiftlogs.ds_sft_hrs} > 24 and {bio_device_shiftlogs.ds_shift_in} <> DATE(1900,01,01)) "
        End If
        
        
        If opt_allcat.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode & ds & io_opt)
           Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and ({bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) or {bio_device_shiftlogs.ds_shift_out} <> Date (1900,01,01)) " & sft & mcode & ds & io_opt)
           pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode & ds & io_opt
        ElseIf opt_staff.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("left({bio_empmas.bioemp_team},5) = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode & ds & io_opt)
        Else
           Cry_rep1.ReplaceSelectionFormula (" left({bio_empmas.bioemp_team},5) <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode & ds & io_opt)
        End If
        
   ElseIf cmb_rep.Text = "Miss IN / OUT punch Report" Then
        Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_missout.rpt"
        If opt_allcat.Value = True Then
           ''Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) " & ds)
           ''Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in2} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out2} = Date (1900,01,01)) or ({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) " & ds)
           
           ''pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) " & ds
           Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and  (({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or ({bio_device_shiftlogs.ds_shift_in} = Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} > Date (1900,01,01)) or ({bio_device_shiftlogs.ds_shift_in2} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out2} = Date (1900,01,01)) or ({bio_device_shiftlogs.ds_shift_in2} = Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out2} > Date (1900,01,01))) " & ds)
           pst_qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in2} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out2} = Date (1900,01,01)) or ({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) " & ds
           
        ElseIf opt_staff.Value = True Then
           Cry_rep1.ReplaceSelectionFormula ("left({bio_empmas.bioemp_team},5) = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01) ")
           Cry_rep1.ReplaceSelectionFormula ("left({bio_empmas.bioemp_team},5) = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) ")
        Else
           Cry_rep1.ReplaceSelectionFormula (" left({bio_empmas.bioemp_team},5) <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01) ")
           Cry_rep1.ReplaceSelectionFormula (" left({bio_empmas.bioemp_team},5) <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and (({bio_device_shiftlogs.ds_shift_in} > Date (1900,01,01) and {bio_device_shiftlogs.ds_shift_out} = Date (1900,01,01)) or (({bio_device_shiftlogs.ds_shift_in} = {bio_device_shiftlogs.ds_shift_out}) and {bio_device_shiftlogs.ds_shift_in} > date(1900,01,01))) ")
        End If
   ElseIf cmb_rep.Text = "Manual Punch Report" Then
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_manualpunch.rpt"
       Cry_rep1.ReplaceSelectionFormula ("{bio_devicelogs.ad_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_devicelogs.ad_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_devicelogs.ad_auto} = 'M'")
   ElseIf cmb_rep.Text = "Daily Cost Report" Then
        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_emp_daily_cost]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_emp_daily_cost] "
        paydb.Execute (pst_qry)

   
''        pst_qry = "create view vew_emp_daily_cost as " _
''                  & "select  fpcode,rdate, sum(ds_sft_hrs) as ds_sft_hrs, sum(woot) as woot,sum(ot) as ot,sum(hot) as hot   from " _
''                  & "(select ds_fpcode as fpcode,ds_date as rdate,ds_status ,ds_sft_hrs,0  as woot,0 as ot , 0 as hot from bio_device_shiftlogs where ds_date between '" & Format$(st_date, "mm/dd/yyyy") & "' and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " Union All " _
''                  & " select w_emp_fpcode as fpcode,w_date as rdate,'' as ds_status,0 as ds_sft_hrs, w_wo_ot_hrs as woot,w_accepted_hrs as ot ,w_holiday_ot_hrs as hot from bio_worker_daily_pihrs where w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " )a group by fpcode,rdate "
''        paydb.Execute (pst_qry)

''        pst_qry = "create view vew_emp_daily_cost as " _
''                  & "select  fpcode,rdate, sum(ds_sft_hrs) as ds_sft_hrs, sum(woot) as woot,sum(ot) as ot,sum(hot) as hot , ds_shift_actual  from " _
''                  & "(select ds_fpcode as fpcode,ds_date as rdate,ds_status ,ds_sft_hrs,0  as woot,0 as ot , 0 as hot , ds_shift_actual from bio_device_shiftlogs where ds_date between '" & Format$(st_date, "mm/dd/yyyy") & "' and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " Union All " _
''                  & " select w_emp_fpcode as fpcode,w_date as rdate,'' as ds_status,0 as ds_sft_hrs, w_wo_ot_hrs as woot,w_accepted_hrs as ot ,w_holiday_ot_hrs as hot , ds_shift_actual from bio_worker_daily_pihrs  , bio_device_shiftlogs   where  w_emp_fpcode  = ds_fpcode and w_date = ds_date and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " )a group by fpcode,rdate, ds_shift_actual "
''        paydb.Execute (pst_qry)




''        pst_qry = "create view vew_emp_daily_cost as " _
''                  & "select  fpcode,rdate, sum(ds_sft_hrs) as ds_sft_hrs, sum(woot) as woot,sum(ot) as ot,sum(hot) as hot , ds_shift_actual ,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from " _
''                  & "(select ds_fpcode as fpcode,ds_date as rdate,ds_status ,ds_sft_hrs,0  as woot,0 as ot , 0 as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_device_shiftlogs where ds_date between '" & Format$(st_date, "mm/dd/yyyy") & "' and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " Union All " _
''                  & " select w_emp_fpcode as fpcode,w_date as rdate,'' as ds_status,0 as ds_sft_hrs, w_wo_ot_hrs + w_woot_days_hrs as woot,w_accepted_hrs as ot ,w_holiday_ot_hrs as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_worker_daily_pihrs  , bio_device_shiftlogs   where  w_emp_fpcode  = ds_fpcode and w_date = ds_date and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
''                  & " )a group by fpcode,rdate, ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 "
''        paydb.Execute (pst_qry)
''

        pst_qry = "create view vew_emp_daily_cost as " _
                  & "select  fpcode,rdate, sum(ds_sft_hrs) as ds_sft_hrs, sum(woot) as woot,sum(ot) as ot,sum(hot) as hot , ds_shift_actual ,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from " _
                  & "(select ds_fpcode as fpcode,ds_date as rdate,ds_status ,case when ds_sft_hrs > 0 then ds_sft_hrs else case when ds_status in ('ML','C.H','WOP(OD)','P(OD)','WO') then 9  else 0  end end as ds_sft_hrs,0  as woot,0 as ot , 0 as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_device_shiftlogs where ds_date between '" & Format$(st_date, "mm/dd/yyyy") & "' and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
                  & " Union All " _
                  & " select w_emp_fpcode as fpcode,w_date as rdate,'' as ds_status,0 as ds_sft_hrs, w_wo_ot_hrs + w_woot_days_hrs as woot,w_accepted_hrs as ot ,w_holiday_ot_hrs as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_worker_daily_pihrs  , bio_device_shiftlogs   where  w_emp_fpcode  = ds_fpcode and w_date = ds_date and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
                  & " )a group by fpcode,rdate, ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 "
        paydb.Execute (pst_qry)
        
        
        
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_dailycost.rpt"
''       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_dailycost_temp.rpt"
       Cry_rep1.ReplaceSelectionFormula (" {bio_empmas.bioemp_fpcode} > 1000 and  {vew_emp_daily_cost.rdate} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {vew_emp_daily_cost.rdate} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {vew_emp_daily_cost.ds_sft_hrs}+{vew_emp_daily_cost.woot}+{vew_emp_daily_cost.ot} > 0")
   
''   ElseIf cmb_rep.Text = "Worker Cost Report" Then
''        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_emp_daily_cost_worker]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
''                   & " drop view [dbo].[vew_emp_daily_cost_worker] "
''        paydb.Execute (pst_qry)
''
''        pst_qry = "create view vew_emp_daily_cost_worker as " _
''                & " select  dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,sum(w_tot_ot_hrs) as othrs ,s_grosspay/240 as hrwages ,s_netpay  from bio_worker_daily_pihrs a , emp_salary b , " _
''                & " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = 1 and s_year = 2022 and s_empcode = w_emp_fpcode and s_empcat = w_cat and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "'" _
''                & " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay"
''
''pst_qry = "create view vew_emp_daily_cost_worker as " _
''& " select  dept_name,s_deptcode,emp_name,s_empcode,sum(s_grosspay) as s_grosspay ,sum(othrs) as othrs ,sum(s_netpay) as s_netpay from (" _
''& " select  dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,0 as othrs,s_netpay  from emp_salary ,  " _
''& " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = 1 and s_year =2022 and s_empcat = 'W'  " _
''& " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay  " _
''& " Union All  " _
''& " select  dept_name,s_deptcode,emp_name,s_empcode,0 as s_grosspay,sum(w_tot_ot_hrs) as othrs ,0 as s_netpay  from bio_worker_daily_pihrs a , emp_salary b ,  " _
''& " pdept_mas c , emp_mas d where dept_code = s_deptcode and s_empcode = emp_code and s_month = 1 and s_year =2022 and s_empcode = w_emp_fpcode and s_empcat = w_cat and w_date between '01/01/2022' and '01/31/2022'  " _
''& " group by dept_name,s_deptcode,emp_name,s_empcode,s_grosspay,s_netpay ) a group by dept_name,s_deptcode,emp_name,s_empcode"
''
''        paydb.Execute (pst_qry)
''
''       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_worker_cost.rpt"
''       Cry_rep1.ReplaceSelectionFormula ("")
   
   ElseIf cmb_rep.Text = "Leave / Absent Abstract" Then
        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_attn ]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_attn ] "
        paydb.Execute (pst_qry)
        
        pst_qry = "create view vew_attn as " _
                  & "select  fpcode,rdate, sum(ds_sft_hrs) as ds_sft_hrs, sum(woot) as woot,sum(ot) as ot,sum(hot) as hot , ds_shift_actual ,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from " _
                  & "(select ds_fpcode as fpcode,ds_date as rdate,ds_status ,case when ds_sft_hrs > 0 then ds_sft_hrs else case when ds_status in ('ML','C.H','WOP(OD)','P(OD)','WO') then 9  else 0  end end as ds_sft_hrs,0  as woot,0 as ot , 0 as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_device_shiftlogs where ds_date between '" & Format$(st_date, "mm/dd/yyyy") & "' and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
                  & " Union All " _
                  & " select w_emp_fpcode as fpcode,w_date as rdate,'' as ds_status,0 as ds_sft_hrs, w_wo_ot_hrs + w_woot_days_hrs as woot,w_accepted_hrs as ot ,w_holiday_ot_hrs as hot , ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 from bio_worker_daily_pihrs  , bio_device_shiftlogs   where  w_emp_fpcode  = ds_fpcode and w_date = ds_date and w_date between '" & Format$(st_date, "mm/dd/yyyy") & "'  and '" & Format$(end_date, "mm/dd/yyyy") & "' " _
                  & " )a group by fpcode,rdate, ds_shift_actual,ds_shift_in,ds_shift_out,ds_shift_in2,ds_shift_out2 "
        
         pst_qry = "create view vew_attn as " _
                   & " select ds_fpcode as fpcode,ds_date as rdate,ds_status as rstatus , '' as emp_reason   from bio_device_shiftlogs where ds_date >= '" & Format$(st_date, "mm/dd/yyyy") & "' and ds_date <= '" & Format$(end_date, "mm/dd/yyyy") & "'  and ds_status in ('A') and ds_fpcode > 1000 " _
                   & " Union All " _
                   & " select emp_fpcode as fpcode,emp_leave_date as rdate,emp_leave_type as rstatus ,emp_reason  from bio_empleave where emp_leave_date >= '" & Format$(st_date, "mm/dd/yyyy") & "' and  emp_leave_date <= '" & Format$(end_date, "mm/dd/yyyy") & "'"
        
        paydb.Execute (pst_qry)
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_employee_daily_attn_status.rpt"
       If ds = "" Then
          ds = "({emp_mas.EMP_STATUS} <> 'R')  or   ({emp_mas.EMP_STATUS} = 'R' and {emp_mas.EMP_RESIGNEDDATE} > date(" & Format$(end_date, "yyyy,mm,dd") & ")  )"
          
       Else
           ds = ds + " and ({emp_mas.EMP_STATUS} <> 'R')  or   ({emp_mas.EMP_STATUS} = 'R' and {emp_mas.EMP_RESIGNEDDATE} > date(" & Format$(end_date, "yyyy,mm,dd") & ")  )"
       End If
       
       Cry_rep1.ReplaceSelectionFormula (ds)
   ElseIf cmb_rep.Text = "Absent Details" Then
        pst_qry = "if exists (select * from sysobjects where id = object_id(N'[dbo].[vew_attn ]') and OBJECTPROPERTY(id, N'IsView') = 1)" _
                   & " drop view [dbo].[vew_attn ] "
        paydb.Execute (pst_qry)
        

         pst_qry = "create view vew_attn as " _
                   & " select ds_fpcode as fpcode,ds_date as rdate,ds_status as rstatus , '' as emp_reason   from bio_device_shiftlogs where ds_date >= '" & Format$(st_date, "mm/dd/yyyy") & "' and ds_date <= '" & Format$(end_date, "mm/dd/yyyy") & "'  and ds_status in ('A') and ds_fpcode > 1000 "
        
        paydb.Execute (pst_qry)
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_employee_daily_absent.rpt"
       If ds = "" Then
          ds = "({emp_mas.EMP_STATUS} <> 'R')  or   ({emp_mas.EMP_STATUS} = 'R' and {emp_mas.EMP_RESIGNEDDATE} > date(" & Format$(end_date, "yyyy,mm,dd") & ")  )"
          
       Else
           ds = ds + " and ({emp_mas.EMP_STATUS} <> 'R')  or   ({emp_mas.EMP_STATUS} = 'R' and {emp_mas.EMP_RESIGNEDDATE} > date(" & Format$(end_date, "yyyy,mm,dd") & ")  )"
       End If
       
       Cry_rep1.ReplaceSelectionFormula (ds)
           
   ElseIf cmb_rep.Text = "Late Attendance Report" Then
       
       Cry_rep1.Formulas(3) = ("repname = '" & cmb_rep.Text & "' + ' for the period from' ")
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_late_attendance.rpt"
       Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")    and {bio_device_shiftlogs.ds_fpcode} > 1000   and {bio_device_shiftlogs.ds_status} <> 'P(OD)' and {bio_device_shiftlogs.ds_per_hrs} = 0 and {bio_empmas.bioemp_fpcode} > 110 and {bio_empmas.bioemp_fpcode} < 20000 " & sft & mcode)
       ''qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_status} <> 'P(OD)'  "
   
   ElseIf cmb_rep.Text = "Doubtful Punches" Then
       
   ''    Cry_rep1.Formulas(3) = ("repname = '" & cmb_rep.Text & "'")
       Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_doubtful_punches.rpt"
       Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")    and {bio_device_shiftlogs.ds_fpcode} > 1000   and ({bio_device_shiftlogs.ds_status} = 'A' or {bio_device_shiftlogs.ds_status} = '폩폗'  or {bio_device_shiftlogs.ds_status} = '폩L') ")
       Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")    and {bio_device_shiftlogs.ds_fpcode} > 1000   and ({bio_device_shiftlogs.ds_status} = '폩폗'  or {bio_device_shiftlogs.ds_status} = '폩L'  or {bio_device_shiftlogs.ds_status} = 'WOL') ")
       '' qry = "{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")    and {bio_device_shiftlogs.ds_fpcode} > 1000   and {bio_device_shiftlogs.ds_status} = 'A' or {bio_device_shiftlogs.ds_status} = '폩폗'  or {bio_device_shiftlogs.ds_status} = '폩L' "
      
   End If
   
   
''
''   If opt_allcat.Value = True Then
''      Cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & sft & mcode)
''   ElseIf opt_staff.Value = True Then
''      Cry_rep1.ReplaceSelectionFormula ("{bio_empmas.bioemp_team} = 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01) " & sft & mcode)
''
''   Else
''      Cry_rep1.ReplaceSelectionFormula (" isnull({bio_device_shiftlogs.ds_shift_in}) = false and  {bio_empmas.bioemp_team} <> 'STAFF' and {bio_device_shiftlogs.ds_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.ds_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ") " & sft & mcode)
''
''   End If
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1
End Sub
