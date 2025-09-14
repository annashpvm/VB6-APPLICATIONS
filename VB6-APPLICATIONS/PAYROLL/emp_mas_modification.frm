VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form emp_mas_modification 
   Caption         =   "EMPLOYEE MASTER MODIFICATIONS"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   16755
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3120
         TabIndex        =   11
         Top             =   4920
         Width           =   3855
         Begin VB.CommandButton NEW 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&New"
            Height          =   705
            Left            =   120
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_modification.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton edit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Edit"
            Height          =   705
            Left            =   840
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_modification.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Save"
            Height          =   705
            Left            =   1560
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_modification.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Refresh 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Refresh"
            Height          =   705
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_modification.frx":133E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton exit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Exit"
            Height          =   705
            Left            =   3000
            MaskColor       =   &H000000FF&
            Picture         =   "emp_mas_modification.frx":19A8
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   1920
         TabIndex        =   6
         Top             =   9240
         Visible         =   0   'False
         Width           =   3735
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   120324097
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   120324097
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   3600
         TabIndex        =   3
         Top             =   720
         Width           =   3255
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
            TabIndex        =   5
            Top             =   120
            Width           =   1215
         End
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
            TabIndex        =   4
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame_emp 
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
         Height          =   2895
         Left            =   1200
         TabIndex        =   1
         Top             =   1320
         Width           =   7815
         Begin VB.TextBox txt_dept 
            Height          =   495
            Left            =   2400
            TabIndex        =   22
            Top             =   1200
            Width           =   4575
         End
         Begin VB.ComboBox cmb_weekoff 
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2160
            Width           =   4695
         End
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
            Left            =   2400
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label3 
            Caption         =   "Department"
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
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Week Off"
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
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Label lbl_emp 
         Alignment       =   2  'Center
         Caption         =   "Employee Master Modification"
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
         TabIndex        =   17
         Top             =   120
         Width           =   4575
      End
   End
End
Attribute VB_Name = "emp_mas_modification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_employee_Click()
    If opt_staff.Value = False And opt_worker.Value = False Then Exit Sub
    Set payrs = New ADODB.Recordset
 
 
 
    sql = "Select * from  emp_mas where emp_company = '" & company_code & "' and emp_name = '" & cmb_employee.Text & "'"
    If opt_staff.Value = True Then
       sql = "Select * from  emp_mas , pdept_mas where dept_code = emp_dept and emp_cat = 'S' and  emp_company = '" & company_code & "'  and emp_name = '" & cmb_employee.Text & "'"
       sql = "Select * from  emp_mas , pdept_mas where dept_code = emp_dept and emp_cat = 'S' and  emp_company = '" & company_code & "'  and emp_code = " & cmb_employee.ItemData(cmb_employee.ListIndex) & ""
    Else
       sql = "Select * from  emp_mas , pdept_mas where dept_code = emp_dept and emp_cat = 'W' and  emp_company = '" & company_code & "'  and emp_name = '" & cmb_employee.Text & "'"
       sql = "Select * from  emp_mas , pdept_mas where dept_code = emp_dept and emp_cat = 'W' and  emp_company = '" & company_code & "'  and emp_code = " & cmb_employee.ItemData(cmb_employee.ListIndex) & ""
  
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_weekoff.Text = payrs!emp_holiday
        txt_dept.Text = payrs!dept_name
        payrs.MoveNext
    Wend
    payrs.Close
    

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmb_weekoff.AddItem "Default"
    cmb_weekoff.AddItem "SUNDAY"
    cmb_weekoff.AddItem "MONDAY"
    cmb_weekoff.AddItem "TUESDAY"
    cmb_weekoff.AddItem "WEDNESDAY"
    cmb_weekoff.AddItem "THURSDAY"
    cmb_weekoff.AddItem "FRIDAY"
    cmb_weekoff.AddItem "SATURDAY"
    get_employee
End Sub



Private Sub opt_staff_Click()
get_employee
End Sub

Private Sub opt_worker_Click()
get_employee
End Sub

Private Sub SAVE_Click()
    sql = "update emp_mas set emp_holiday = '" & cmb_weekoff.Text & "' where emp_company = '" & company_code & "' and emp_name = '" & cmb_employee.Text & "'"
    paydb.Execute sql
    MsgBox ("Updated...")
End Sub

Public Sub get_employee()
    cmb_employee.Clear
    Set payrs = New ADODB.Recordset
''
''    If opt_staff.Value = True Then
''       sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_status = 'A' order by  EMP_NAME")
''    Else
''       sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_status = 'A' order by  EMP_NAME")
''    End If
    
    If opt_staff.Value = True Then
       sql = ("Select emp_code  as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and emp_status = 'A' order by  EMP_NAME")
    Else
       sql = ("Select emp_code  as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_status = 'A' order by  EMP_NAME")
    End If
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_employee.AddItem payrs!emp_name
        cmb_employee.ItemData(cmb_employee.NewIndex) = payrs!ecode
        payrs.MoveNext
    Wend
    payrs.Close
    

End Sub
