VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form emp_mas_slot_entry 
   Caption         =   "SALARY SLOT ENTRY"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.Frame Frame_emp 
         Caption         =   "Employee Name"
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
         Height          =   855
         Left            =   5760
         TabIndex        =   29
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmb_employee 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   4440
         TabIndex        =   26
         Top             =   720
         Width           =   3855
         Begin VB.OptionButton opt_empwise 
            Caption         =   "Employee wise"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   1680
            TabIndex        =   28
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton opt_empall 
            Caption         =   "Emp.Group"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "WORK PLACE"
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
         Height          =   1695
         Left            =   10080
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton opt_all 
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton opt_vpt 
            Caption         =   "VPT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opt_cbe 
            Caption         =   "CBE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame frame_type 
         Caption         =   "Employee Type"
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
         Height          =   855
         Left            =   840
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox cmb_designation 
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
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   19
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   1920
         TabIndex        =   9
         Top             =   9240
         Visible         =   0   'False
         Width           =   3735
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   152633345
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   1920
            TabIndex        =   11
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   152633345
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   2880
         TabIndex        =   3
         Top             =   7200
         Width           =   3855
         Begin VB.CommandButton exit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Exit"
            Height          =   705
            Left            =   3000
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_slot_entry.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Refresh 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Refresh"
            Height          =   705
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_slot_entry.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Save"
            Height          =   705
            Left            =   1560
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_slot_entry.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton edit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Edit"
            Height          =   705
            Left            =   840
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_slot_entry.frx":1116
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton NEW 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&New"
            Height          =   705
            Left            =   120
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_slot_entry.frx":1780
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   735
         End
      End
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7800
         TabIndex        =   1
         Top             =   2400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   3450
         Left            =   840
         TabIndex        =   14
         Top             =   3120
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   6085
         _Version        =   393216
         Cols            =   4
         FixedCols       =   3
         BackColorFixed  =   16776960
         BackColorSel    =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl_emp 
         Alignment       =   2  'Center
         Caption         =   "SALARY SLOT ENTRY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   120
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         Height          =   6615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   9735
      End
      Begin VB.Label Label2 
         Caption         =   "SLOT"
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
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   2520
         Width           =   885
      End
   End
End
Attribute VB_Name = "emp_mas_slot_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paydb As New ADODB.Connection
Dim payrs As New ADODB.Recordset

Private Sub emp_designation_Click()
    
End Sub

Private Sub cmb_designation_Click()
    If cmb_designation.ListIndex = -1 Then Exit Sub
    fillgrid
    filldata_staff
End Sub

Private Sub cmb_employee_Click()
    fillgrid
    ''Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    If opt_vpt.Value = True Then
       loc = "and emp_workplace = 'MILL'"
    ElseIf opt_cbe.Value = True Then
       loc = "and emp_workplace = 'CBE'"
    End If
    
    sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_name = '" & cmb_employee.Text & "' and emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With flx_data
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_salary_slot")
             
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub cmd_update_Click()
  If cmb_slot.Text = "" Then Exit Sub
  Dim i As Integer
    For i = 1 To flx_data.Rows - 1
       flx_data.TextMatrix(i, 3) = cmb_slot.Text
    Next
End Sub

Private Sub exit_Click()
    paydb.Close
    Unload Me
End Sub

Private Sub Form_Load()
    sql = ("select pdesi_code,pdesi_name from emp_mas , pdesi_mas where emp_design = pdesi_code and emp_cat = 'S' group by pdesi_code,pdesi_name")
    cmb_slot.AddItem "SLOT1"
    cmb_slot.AddItem "SLOT2"
    cmb_slot.AddItem "SLOT3"
    cmb_slot.AddItem "SLOT4"
    cmb_slot.AddItem "SLOT5"
    cmb_slot.AddItem "SLOT6"
    
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    cmb_designation.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_designation.AddItem payrs(1)
        cmb_designation.ItemData(cmb_designation.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
    fillgrid
End Sub

Function fillgrid()
    With flx_data
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "SLOT"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        
    End With
End Function

Function filldata_staff()
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    
    sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_design = " & cmb_designation.ItemData(cmb_designation.ListIndex) & " and emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With flx_data
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_salary_slot")
             
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    
    sql = ("Select emp_code as ecode,* from  emp_mas where emp_design = " & cmb_designation.ItemData(cmb_designation.ListIndex) & " and emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & " order by EMP_CODE")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With flx_data
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_salary_slot")
             
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    
End Function

Function filldata_worker()
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    
    sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With flx_data
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_salary_slot")
             
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    sql = ("Select emp_code as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A'  and EMP_CODE  like '%A'")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With flx_data
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_salary_slot")
             
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    
End Function

Private Sub opt_empall_Click()
    Frame_emp.Visible = False
    frame_type.Visible = True
    cmb_designation.Clear
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    payrs.MoveFirst
    While Not payrs.EOF
        cmb_designation.AddItem payrs(1)
        cmb_designation.ItemData(cmb_designation.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub opt_empwise_Click()
    If opt_staff.Value = True Then
       Frame_emp.Visible = True
       frame_type.Visible = False
    End If
End Sub

Private Sub opt_staff_Click()
    fillgrid
    filldata_emp
    frame_type.Visible = True
    If cmb_designation.ListIndex = -1 Then Exit Sub
    filldata_staff
End Sub

Private Sub opt_worker_Click()
    filldata_emp
    frame_type.Visible = False
    Frame_emp.Visible = False
    fillgrid
    filldata_worker
End Sub

Private Sub Option1_Click()

End Sub

Private Sub SAVE_Click()
On Error GoTo err_handler
  If flx_data.Rows < 2 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
  Me.MousePointer = 11
  For i = 1 To flx_data.Rows - 1
      sql = "update emp_mas set emp_salary_slot = '" & flx_data.TextMatrix(i, 3) & "'  where emp_company = " & company_code & " and  emp_code = '" & flx_data.TextMatrix(i, 1) & "'"
      paydb.Execute sql
  Next
  MsgBox ("Records are saved")
  
  fillgrid
  Me.MousePointer = 1
  Exit Sub
err_handler:

    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume

End Sub


 
Function filldata_emp()
    cmb_employee.Clear
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    
    sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_employee.AddItem payrs!emp_name
        payrs.MoveNext
    Wend
    payrs.Close
    
    sql = ("Select emp_code as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & " order by EMP_CODE")
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_employee.AddItem payrs!emp_name
        payrs.MoveNext
        payrs.MoveNext
    Wend
    payrs.Close
    
End Function

