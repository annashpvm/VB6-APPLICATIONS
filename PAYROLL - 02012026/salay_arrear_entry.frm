VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form salary_arrear_entry 
   Caption         =   "ARREARS DETAILS ENTRY "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3600
      TabIndex        =   9
      Top             =   6840
      Width           =   4815
      Begin VB.CommandButton NEW 
         Caption         =   "&New"
         Height          =   825
         Left            =   0
         Picture         =   "salay_arrear_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton exit 
         Caption         =   "&Exit"
         Height          =   825
         Left            =   3840
         Picture         =   "salay_arrear_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Refresh 
         Caption         =   "&Refresh"
         Height          =   825
         Left            =   2880
         Picture         =   "salay_arrear_entry.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton edit 
         Caption         =   "&Edit"
         Height          =   825
         Left            =   960
         Picture         =   "salay_arrear_entry.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton save 
         Caption         =   "&Save"
         Height          =   825
         Left            =   1920
         Picture         =   "salay_arrear_entry.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   870
      Left            =   3555
      TabIndex        =   0
      Top             =   225
      Width           =   4410
      Begin VB.OptionButton opt_worker 
         Caption         =   "WORKER"
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
         Height          =   615
         Left            =   2340
         TabIndex        =   2
         Top             =   165
         Width           =   1710
      End
      Begin VB.OptionButton opt_staff 
         Caption         =   "STAFF"
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
         Height          =   615
         Left            =   450
         TabIndex        =   1
         Top             =   195
         Width           =   1530
      End
   End
   Begin VB.ComboBox cmb_year 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   8265
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1695
      Width           =   2265
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
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   3720
      TabIndex        =   4
      Top             =   1650
      Width           =   2790
   End
   Begin MSFlexGridLib.MSFlexGrid arrear_flex 
      Height          =   3945
      Left            =   885
      TabIndex        =   6
      Top             =   2835
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   6959
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
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   2100
      TabIndex        =   3
      Top             =   1335
      Width           =   8985
      Begin VB.Label Label2 
         Caption         =   "YEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   5130
         TabIndex        =   8
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
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
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   435
         Width           =   1455
      End
   End
End
Attribute VB_Name = "salary_arrear_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Integer
Dim rec_chk As Integer
Dim arrear_amt As Integer
Function fillgrid()
    With arrear_flex
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Amount"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColWidth(2) = 5000
        .ColWidth(3) = 1500
    End With
End Function

Private Sub edit_Click()
    rec_chk = 0
''    If endrow = 0 Then
''       MsgBox (" Details not available ")
''       Exit Sub
''    End If
    If cmb_month.Text = "" Then
       MsgBox ("Select the Month")
       Exit Sub
    End If
''    If opt_staff.Value = True Then
''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and e_year = " & Val(cmb_year.Text) & " and (emp_type = 0 or emp_type = 1)order by emp_dept")
''    Else
''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and e_year = " & Val(cmb_year.Text) & " and (emp_type = 2 or emp_type = 3)order by emp_dept")
''    End If
    If opt_staff.Value = True Then
       sql = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 0 or e_emptype = 1)")
    Else
       sql = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & "and e_company = '" & company_code & "' and (e_emptype = 2 or e_emptype = 3)")
    End If
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          For i = 1 To arrear_flex.Rows - 1
              If Trim(arrear_flex.TextMatrix(i, 1)) <> "" Then
                 If arrear_flex.TextMatrix(i, 1) = payrs.Fields("e_empcode") Then
                    arrear_flex.TextMatrix(i, 3) = payrs.Fields("e_amount")
                 End If
              End If
          Next
         payrs.MoveNext
    Wend



End Sub

Private Sub exit_Click()
   Unload Me
End Sub
 
Private Sub Form_Load()
    opt_staff.Value = True
    refresh_Click
End Sub

Private Sub arrear_flex_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler
 Dim fin_selrow%, fin_selcol%
 fin_selrow = arrear_flex.Row
 fin_selcol = arrear_flex.Col
 With arrear_flex
 Select Case fin_selcol
        Case 3
            If KeyAscii <> 13 Then
               KeyAscii = Numeric_Chk(KeyAscii, arrear_flex.TextMatrix(fin_selrow, fin_selcol), 9, 6, 2)
            End If
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
               arrear_flex.TextMatrix(fin_selrow, fin_selcol) = arrear_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
               If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                  KeyAscii = 0
            End If
    End Select
 End With
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub


Private Sub opt_staff_Click()
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1)order by emp_dept")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    endrow = 0
    fillgrid
    While Not payrs.EOF
        With arrear_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs(0)
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
End Sub

Private Sub opt_worker_Click()
    endrow = 0
    fillgrid
    sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3)order by emp_dept")
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''     payrs.MoveFirst
    While Not payrs.EOF
        With arrear_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs(0)
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
End Sub

Private Sub refresh_Click()
    cmb_month.Clear
    cmb_year.Clear
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
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    
    fillgrid
    If opt_staff.Value = True Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1)order by emp_dept")
    Else
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3)order by emp_dept")
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    endrow = 0
    While Not payrs.EOF
        With arrear_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs(0)
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
End Sub

Private Sub SAVE_Click()
  rec_chk = 0
  If endrow = 0 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
  If cmb_month.Text = "" Then
     MsgBox ("Select the Month")
     Exit Sub
  End If
  On Error GoTo err_handler
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  paydb.BeginTrans
  
  If opt_staff.Value = True Then
''     sql = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 0 or e_emptype = 1)")
     sql = ("delete from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 0 or e_emptype = 1)")
  Else
''     sql = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & "and e_company = '" & company_code & "' and (e_emptype = 2 or e_emptype = 3)")
     sql = ("delete from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 2 or e_emptype = 3)")
  End If
  
  paydb.Execute sql
  sql = "select * from arrear_entry where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 1 To endrow
      If Trim(arrear_flex.TextMatrix(i, 2)) <> "" Then
            payrs.AddNew
            payrs.Fields("e_company") = company_code
            payrs.Fields("e_finyear") = finyear
            payrs.Fields("e_empcode") = arrear_flex.TextMatrix(i, 1)
            find_empdetails (arrear_flex.TextMatrix(i, 1))
            payrs.Fields("e_emptype") = emptypecode
            payrs.Fields("e_month") = cmb_month.ItemData(cmb_month.ListIndex)
            payrs.Fields("e_year") = Val(cmb_year.Text)
            If arrear_flex.TextMatrix(i, 3) <> "" Then
                payrs.Fields("e_amount") = arrear_flex.TextMatrix(i, 3)
            Else
                payrs.Fields("e_amount") = 0
            End If
            rec_chk = 1
            payrs.Update
      End If
  Next
  endrow = 0
  If rec_chk = 1 Then
     MsgBox ("Records are saved")
     fillgrid
     cmb_month.Text = ""
  Else
     MsgBox ("Details not available for save")
  End If
  paydb.CommitTrans
  refresh_Click
  Exit Sub
err_handler:
    paydb.RollbackTrans
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
  
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
Dim ret%, pst_row%, pst_rawname$
Static PrevIndex%
    Select Case KeyAscii
        Case 8
            If Trim(fst_item) <> "" Then fst_item = Mid(fst_item, 1, Len(fst_item) - 1)
        Case 13
             pst_row = arrear_flex.Row
        Case Else
            fst_item = txt & Chr(KeyAscii)
            If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    End Select
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_handler
    If KeyCode = 40 Then
        txt.SetFocus
    End If
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
      
Private Sub arrear_flex_EnterCell()
On Error GoTo err_handler
    Select Case arrear_flex.Col
        Case 1, 2, 3
            txt.Visible = False
    End Select
    Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

End Sub


Function find_arear_amt(empcode As Integer)
    Set paydb = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
'    If opt_staff.Enabled = True Then
    If opt_staff.Value = True Then
       sql2 = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 0 or e_emptype = 1) and e_empcode = " & empcode)
    Else
       sql2 = ("Select * from  arrear_entry where e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_company = '" & company_code & "' and (e_emptype = 2 or e_emptype = 3) and e_empcode = " & empcode)
    End If
    paydb.Open pay
    payrs2.Open sql2, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       arrear_amt = payrs2.Fields("e_amount")
    Else
       arrear_amt = 0
    End If
End Function
