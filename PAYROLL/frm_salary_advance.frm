VERSION 5.00
Begin VB.Form frm_salary_advance 
   Caption         =   "Salary Advance Entry"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   9135
         Begin VB.TextBox txt_idcode 
            Enabled         =   0   'False
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
            Left            =   3000
            TabIndex        =   10
            Top             =   1200
            Width           =   1845
         End
         Begin VB.TextBox txt_dept 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3000
            TabIndex        =   9
            Top             =   720
            Width           =   3540
         End
         Begin VB.TextBox txt_fpcode 
            Enabled         =   0   'False
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
            Left            =   6600
            TabIndex        =   8
            Top             =   1200
            Width           =   2040
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
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   3000
            TabIndex        =   6
            Top             =   240
            Width           =   6030
         End
         Begin VB.Label Label5 
            Caption         =   "Emp. Code"
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
            Height          =   360
            Left            =   480
            TabIndex        =   13
            Top             =   1320
            Width           =   1305
         End
         Begin VB.Label Label6 
            Caption         =   "Department"
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
            Height          =   390
            Left            =   480
            TabIndex        =   12
            Top             =   840
            Width           =   2370
         End
         Begin VB.Label Label7 
            Caption         =   "F.P. Code"
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
            Height          =   270
            Left            =   5040
            TabIndex        =   11
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Employee Name"
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
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2280
         TabIndex        =   1
         Top             =   120
         Width           =   6615
         Begin VB.OptionButton opt_retainer 
            Caption         =   "Retainer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton opt_worker 
            Caption         =   "Workers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   600
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frm_salary_advance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub getdata()
    cmb_employee.Clear
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
     If opt_staff.Value = True Then
        sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat in ('S','M') and emp_status = 'A' order by emp_name"
     ElseIf opt_worker.Value = True Then
        sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat in ('W') and emp_status = 'A' order by emp_name"
     End If
     paydb.Open pay
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF Then
            payrs.MoveFirst
            cmb_employee.Clear
            While Not payrs.EOF
                cmb_employee.AddItem payrs("emp_name")
                payrs.MoveNext
            Wend
        End If
    
End Sub

Private Sub cmb_employee_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
        MsgBox "Out Of Financial Year", vbInformation, "Message"
        Exit Sub
    End If
        
    If Trim(cmb_month.Text) = "" Then
       MsgBox ("Select Deduction month")
       Exit Sub
    End If
    '' Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    paydb.Open pay
    If emptype_chk = 0 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('S','M') ")
    ElseIf emptype_chk = 1 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat = 'W' ")
    ElseIf emptype_chk = 2 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('M') ")
    ElseIf emptype_chk = 3 Then
       sql = ("select * from emp_voupay_mast where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('R') ")
    
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
    Else
       emp_idcode = payrs.Fields("emp_code")
       find_deptname (payrs.Fields("emp_dept"))
       dept.Text = dname
       find_desiname (payrs.Fields("emp_design"))
       DESI.Text = dname
       find_etypename (payrs.Fields("emp_type"))
       emptype.Text = dname
       emptypecode = payrs.Fields("emp_type")
    End If

End Sub

Private Sub opt_staff_Click()
    getdata
End Sub

Private Sub opt_worker_Click()
    getdata
End Sub
