VERSION 5.00
Begin VB.Form emp_mas_position 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   3600
         TabIndex        =   20
         Top             =   2280
         Width           =   5655
         Begin VB.OptionButton opt_select 
            Caption         =   "Select"
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
            Left            =   2640
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton Opt_Head 
            Caption         =   " Head"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txt_appointedby 
         Height          =   495
         Left            =   3600
         TabIndex        =   19
         Text            =   " "
         Top             =   5400
         Width           =   5415
      End
      Begin VB.TextBox txt_preinterviewby 
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Text            =   " "
         Top             =   4560
         Width           =   5415
      End
      Begin VB.TextBox txt_interviewername 
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Text            =   " "
         Top             =   3720
         Width           =   5415
      End
      Begin VB.ComboBox cmb_position 
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
         Left            =   3600
         TabIndex        =   13
         Top             =   3000
         Width           =   5535
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Height          =   855
         Left            =   3360
         TabIndex        =   8
         Top             =   6720
         Width           =   3615
         Begin VB.CommandButton NEW 
            Caption         =   "&New"
            Height          =   735
            Left            =   120
            Picture         =   "gm_mas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton emp_edit 
            Caption         =   "&Edit"
            Height          =   735
            Left            =   960
            Picture         =   "gm_mas.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton emp_save 
            Caption         =   "&Save "
            Height          =   735
            Left            =   1800
            Picture         =   "gm_mas.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Exit 
            Caption         =   "&Exit"
            Height          =   735
            Left            =   2640
            Picture         =   "gm_mas.frx":1116
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.ComboBox cmb_dept 
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
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   6150
      End
      Begin VB.TextBox txt_empcode 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cmb_empname 
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
         Left            =   3600
         TabIndex        =   1
         Top             =   960
         Width           =   6150
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Appointed by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   18
         Top             =   5400
         Width           =   2685
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Preliminary Interview by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   16
         Top             =   4560
         Width           =   2685
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Interviewed by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   14
         Top             =   3720
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Superior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Department Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   3165
      End
      Begin VB.Label empcode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1605
      End
      Begin VB.Label empname 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   3165
      End
   End
End
Attribute VB_Name = "emp_mas_position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fst_save As String

Private Sub cmb_dept_Click()
cmb_empname.Clear
txt_empcode.Text = ""
Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    sql = ("Select * from  emp_mas where emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " and emp_company not in (50,90) and emp_status='A'")
    sql = ("select * from (select emp_company,emp_code,emp_name,emp_dept,emp_status,emp_ctc from  emp_mas where emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " and emp_company not in (50,90) and emp_status='A'  union all select emp_company,emp_code,emp_name,emp_dept,emp_status,emp_ctc from emp_voupay_mast where emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " and emp_company not in (50,90) and emp_status='A' ) a order by emp_ctc desc")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_empname.AddItem payrs(2)
        ''cmb_empname.ItemData(cmb_empname.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    
    
End Sub



Private Sub cmb_empname_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select emp_code from  emp_mas where emp_name = '" & cmb_empname.Text & "' union all select emp_code from emp_voupay_mast where  emp_name = '" & cmb_empname.Text & "'")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    If payrs.EOF = False Then
''        cmb_empname.AddItem payrs(5)
        ''cmb_empname.ItemData(cmb_empname.NewIndex) = payrs(0)
''        payrs.MoveNext
''    Wend
        txt_empcode.Text = payrs!emp_code
    End If
    
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    sql = ("Select * from  emp_mas where emp_name <> '" & cmb_empname.Text & "' and emp_company not in (50,90) and emp_status='A' and emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " ")
    sql = ("select * from (select emp_company,emp_code,emp_name,emp_dept,emp_status,emp_ctc from  emp_mas where emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " and emp_name <> '" & cmb_empname.Text & "' and emp_company not in (50,90) and emp_status='A'  union all select emp_company,emp_code,emp_name,emp_dept,emp_status,emp_ctc from emp_voupay_mast where emp_dept=" & cmb_dept.ItemData(cmb_dept.ListIndex) & " and emp_name <> '" & cmb_empname.Text & "' and emp_company not in (50,90) and emp_status='A' ) a order by emp_ctc desc")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_position.AddItem payrs(2)
        cmb_position.ItemData(cmb_position.NewIndex) = payrs(1)
        payrs.MoveNext
    Wend
End Sub

Private Sub cmb_position_Change()
If cmb_empname.Text = "" Then
    MsgBox "Select Employee Name First to continue"
    Exit Sub
End If


End Sub

''Private Sub emp_edit_Click()
''fst_save = "EDIT"
''End Sub

Private Sub emp_save_Click()
paydb.BeginTrans
On Error GoTo err_handler
Dim pst_respo As String
Dim pst_qry As String
Dim no As Integer
If fst_save = "NEW" Then
    If cmb_dept.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Select Department Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_dept.SetFocus
        Exit Sub
    End If
    If cmb_empname.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Select employee Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_empname.SetFocus
        Exit Sub
    End If
    
    If opt_select.Value = True And cmb_position.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Select Superior Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_position.SetFocus
        Exit Sub
    End If
    
    pst_respo = MsgBox("Do You want to save the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        paydb.RollbackTrans
        MousePointer = vbDefault
        Exit Sub
    End If
    Dim rs_pay As New ADODB.Recordset
    pst_qry = "select * from emp_mas_position where p_empcode='" & txt_empcode.Text & "'"
    rs_pay.Open pst_qry, paydb, 1, 2
        If Not (rs_pay.EOF) Then
            rs_pay.Close
            paydb.RollbackTrans
            MsgBox " Details Already Exists ..", vbOKOnly + vbExclamation, "Message"
            MousePointer = vbDefault
            Exit Sub
        End If
    rs_pay.Close
''    pst_qry = "select max(bank_code)+1 as endno from payroll_bank"
''    rs_pay.Open pst_qry, paydb, 1, 2
''    no = 1
''    If Not IsNull(rs_pay!endno) Then
''        If Not rs_pay.EOF Then
''             no = rs_pay!endno
''        End If
''    End If
''    rs_pay.Close
    pst_qry = "insert into emp_mas_position values('" & txt_empcode.Text & "'," & cmb_dept.ItemData(cmb_dept.ListIndex) & "," & cmb_position.ItemData(cmb_position.ListIndex) & ",'" & txt_interviewername.Text & "','" & txt_preinterviewby.Text & "','" & txt_appointedby.Text & "')"
    paydb.Execute pst_qry
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
''    cmb_dept.Text = ""
    Exit Sub
''ElseIf fst_save = "EDIT" Then
''    If cmb_bank.Text = "" Then
''        paydb.RollbackTrans
''        MsgBox " Select Bank Name ", vbOKOnly + vbExclamation, "vbInformation "
''        cmb_bank.SetFocus
''        Exit Sub
''    End If
''    If txt_bank.Text = "" Then
''        paydb.RollbackTrans
''        MsgBox " Enter Bank Name to be Change as ", vbOKOnly + vbExclamation, "vbInformation "
''        txt_bank.SetFocus
''        Exit Sub
''      End If
''    pst_respo = MsgBox("Do You want to Modify the record", vbYesNo + vbInformation, "Information")
''    If pst_respo = vbNo Then
''        paydb.RollbackTrans
''        MousePointer = vbDefault
''        Exit Sub
''    End If
''
''    pst_qry = "update payroll_bank set bank_name='" & txt_bank.Text & "' where bank_name='" & cmb_bank.Text & "'"
''    paydb.Execute pst_qry
''    MsgBox "Records Updated", vbOKOnly + vbInformation, "vbInformation"
''    cmb_bank.Text = ""
''    txt_bank.Text = ""
''    paydb.CommitTrans
End If
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub exit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
fst_save = "NEW"
opt_select.Value = True
Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pdept_mas order by dept_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        cmb_dept.AddItem payrs(1)
        cmb_dept.ItemData(cmb_dept.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    
''    For i = 1 To 10
''    cmb_position.AddItem (i)
''    Next
End Sub

Private Sub NEW_Click()
fst_save = "NEW"
End Sub

Private Sub Opt_Head_Click()
 cmb_position.Enabled = False
End Sub

Private Sub opt_select_Click()
 cmb_position.Enabled = True
End Sub
