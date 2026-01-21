VERSION 5.00
Begin VB.Form frm_control_lock 
   Caption         =   "CONTROL UPDATION"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.Frame frame 
      Caption         =   "Enter Password for Release"
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
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton cmd_ok 
         Caption         =   "OK"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt_pw 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
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
      Left            =   7920
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   3840
      Width           =   4575
      Begin VB.CommandButton cmd_release 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LOCK RELEASE"
         Height          =   975
         Left            =   1560
         Picture         =   "frm_control_lock.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton pay_process 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&LOCK"
         Height          =   975
         Left            =   120
         Picture         =   "frm_control_lock.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   975
         Left            =   3000
         Picture         =   "frm_control_lock.frx":2EB4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      Height          =   1890
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1785
      Width           =   8595
   End
   Begin VB.Label pay_label 
      Caption         =   "LOCKING FOR  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   6795
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1920
      TabIndex        =   6
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      Top             =   2520
      Width           =   885
   End
End
Attribute VB_Name = "frm_control_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim employee_code As Double
Dim employee_type As Integer

Private Sub cmd_ok_Click()
    Dim payrs As New ADODB.Recordset
    If txt_pw.Text = "" Then
       MsgBox ("Enter Password ...")
       Exit Sub
    End If
''    sql = "select * from payroll_password"
    sql = "select * from mas_users where usr_code = 47"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       If payrs("usr_pwd") <> Trim(txt_pw.Text) Then
          MsgBox ("Password Error.. ")
          payrs.Close
          Exit Sub
       End If
    End If
    payrs.Close
    If control_lock = 1 Then
       sql2 = "update payroll_lock  set pay_attn_lock = 'N' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    ElseIf control_lock = 2 Then
        sql2 = "update payroll_lock  set pay_dedu_lock = 'N' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    Else
        sql2 = "update payroll_lock  set pay_salary_lock = 'N' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    End If
    paydb.Execute sql2
    txt_pw.Text = ""
End Sub

Private Sub cmd_release_Click()
    Dim payrs As New ADODB.Recordset
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       If payrs("pay_sp_lock") = "Y" Or payrs("pay_st_lock") = "Y" Or payrs("pay_wp_lock") = "Y" Or payrs("pay_wt_lock") = "Y" Then
          MsgBox ("Salary details are update in the Accounts system. Can't Modify details")
          frame.Visible = False
          payrs.Close
          Exit Sub
       End If
    End If
    
    frame.Visible = True
End Sub

Private Sub exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    If control_lock = 1 Then
       pay_label.Caption = pay_label + " ATTENDANCE "
    ElseIf control_lock = 2 Then
       pay_label.Caption = pay_label + " DUDUCTIONS"
    Else
       pay_label.Caption = pay_label + " SALARY PROCESS"
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
End Sub

Private Sub pay_process_Click()
    Dim payrs As New ADODB.Recordset
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    If control_lock = 1 Then
       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & ""
       payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If payrs.EOF Then
          MsgBox ("Attendance Details are Not Available....")
          payrs.Close
          Exit Sub
       End If
       payrs.Close
    ElseIf control_lock = 2 Then
       sql = "Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_company = '" & company_code & "'"
       payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If payrs.EOF Then
          MsgBox ("Deduction Details are Not Available....")
          payrs.Close
          Exit Sub
       End If
       payrs.Close
    Else
       sql = "select * from emp_salary where s_company = " & company_code & " and s_finyear = " & finyear & " and s_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and s_year = " & Val(cmb_year.Text) & ""
       payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If payrs.EOF Then
          MsgBox ("Salary Details are Not Available....")
          payrs.Close
          Exit Sub
       End If
       payrs.Close
    End If
    sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       payrs.Close
       If control_lock = 1 Then
          sql2 = "insert into payroll_lock values ( " & company_code & ", " & finyear & " ," & cmb_month.ItemData(cmb_month.ListIndex) & " , " & cmb_year.Text & " ,'Y','N','N','N','N','N','N','N','N')"
       ElseIf control_lock = 2 Then
          sql2 = "insert into payroll_lock values ( " & company_code & ", " & finyear & " ," & cmb_month.ItemData(cmb_month.ListIndex) & " , " & cmb_year.Text & " ,'N','Y','N','N','N','N','N','N','N')"
       Else
          sql2 = "insert into payroll_lock values ( " & company_code & ", " & finyear & " ," & cmb_month.ItemData(cmb_month.ListIndex) & " , " & cmb_year.Text & " ,'N','N','Y','N','N','N','N','N','N')"
       End If
       paydb.Execute sql2
    Else
       payrs.Close
       If control_lock = 1 Then
          sql2 = "update payroll_lock  set pay_attn_lock = 'Y' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
       ElseIf control_lock = 2 Then
          sql2 = "update payroll_lock  set pay_dedu_lock = 'Y' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
       Else
          sql2 = "update payroll_lock  set pay_salary_lock = 'Y' where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
       End If
       paydb.Execute sql2
    End If
    MsgBox ("Locked..")
End Sub

