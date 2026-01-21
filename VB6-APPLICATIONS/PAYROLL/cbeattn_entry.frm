VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cbeattn_entry 
   Caption         =   "COIMBATORE STAFF ATTENDANCE"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton NEW 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&New"
      Height          =   825
      Left            =   360
      MaskColor       =   &H000000FF&
      Picture         =   "cbeattn_entry.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton exit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Exit"
      Height          =   825
      Left            =   4200
      MaskColor       =   &H000000FF&
      Picture         =   "cbeattn_entry.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Refresh 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Refresh"
      Height          =   825
      Left            =   3240
      MaskColor       =   &H000000FF&
      Picture         =   "cbeattn_entry.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton edit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Edit"
      Height          =   825
      Left            =   1320
      MaskColor       =   &H000000FF&
      Picture         =   "cbeattn_entry.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      Height          =   825
      Left            =   2280
      MaskColor       =   &H000000FF&
      Picture         =   "cbeattn_entry.frx":1780
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox lst_name 
      Height          =   1425
      Left            =   4020
      TabIndex        =   3
      Top             =   4950
      Width           =   2415
   End
   Begin VB.TextBox txt_itemname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   225
      TabIndex        =   2
      Top             =   5715
      Width           =   1620
   End
   Begin VB.ListBox lst_code 
      Height          =   1425
      Left            =   1395
      TabIndex        =   1
      Top             =   4950
      Width           =   2580
   End
   Begin VB.TextBox txt 
      Height          =   465
      Left            =   285
      TabIndex        =   0
      Top             =   4935
      Width           =   1170
   End
   Begin MSComCtl2.DTPicker attn_dt 
      Height          =   420
      Left            =   8865
      TabIndex        =   9
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60751873
      CurrentDate     =   37519
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   5985
      Left            =   345
      TabIndex        =   10
      Top             =   720
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   10557
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
   Begin VB.Label DATE 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DATE"
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
      Height          =   270
      Left            =   8025
      TabIndex        =   11
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "cbeattn_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Byte
Dim fst_item$
Dim endrow As Byte
Function fillgrid()
    With att_flex
        .Clear
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Atten. Status"
        .TextMatrix(0, 4) = "Extra hours"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 3500
        .ColWidth(4) = 1100
    End With
End Function

Private Sub attn_dt_Change()
      sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(attn_dt, "mm/dd/yyyy") & "'"
      Set paydb = New ADODB.Connection
      Set payrs = New ADODB.Recordset
      paydb.Open pay
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If Not payrs.EOF Then
         attstatus = payrs(1)
      Else
         attstatus = "EARNED LEAVE (FULL DAY)"
      End If
      endrow = 0
      fillgrid
      lst_code.Visible = False
      lst_name.Visible = False
      txt_itemname.Visible = False
      txt.Visible = False
      Set paydb = New ADODB.Connection
      Set payrs = New ADODB.Recordset
      If emptype_chk = 0 Then
         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
      Else
         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
      End If
      paydb.Open pay
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If payrs.EOF Then
         MsgBox ("Pervious date details are missing. First enter for previous date & continue")
         attn_dt = attn_dt - 1
      End If
      i = 1
      If emptype_chk = 0 Then
         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
      Else
         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
      End If
      Set paydb = New ADODB.Connection
      Set payrs = New ADODB.Recordset
      paydb.Open pay
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If Not payrs.EOF Then
         While Not payrs.EOF
              With att_flex
                   .Rows = .Rows + 1
                   find_empdetails (payrs.Fields("attn_empcode"))
                   find_attnstatus (payrs.Fields("attn_status"))
                   .TextMatrix(i, 0) = dept_name
                   .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
                   .TextMatrix(i, 2) = employee_name
                    att_dat = attn_dt
                    find_present_status (payrs(0))
                   .TextMatrix(i, 3) = attstatus
                   .TextMatrix(i, 4) = payrs(5)
                   i = i + 1
                   endrow = endrow + 1
              End With
              payrs.MoveNext
         Wend
      Else
         Set paydb = New ADODB.Connection
         Set payrs = New ADODB.Recordset
         If emptype_chk = 0 Then
            sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1) order by emp_name")
         Else
            sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3) order by emp_name")
         End If
         paydb.Open pay
         payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
         payrs.MoveFirst
         While Not payrs.EOF
               With att_flex
                   .Rows = .Rows + 1
                    find_deptname (payrs.Fields("emp_dept"))
                   .TextMatrix(.Rows - 1, 0) = dname
                   .TextMatrix(.Rows - 1, 1) = payrs(0)
                   .TextMatrix(.Rows - 1, 2) = payrs(5)
                   If Trim((payrs.Fields("emp_holiday"))) = UCase(RTrim(Format(attn_dt, "dddd"))) Then
                      .TextMatrix(.Rows - 1, 3) = "WEEKLY OFF"
                   End If
                    att_dat = attn_dt
                    find_present_status (payrs(0))
                   .TextMatrix(.Rows - 1, 3) = attstatus
                   endrow = endrow + 1
               End With
               payrs.MoveNext
          Wend
      End If
      lst_code.Clear
      Set paydb = New ADODB.Connection
      Set payrs = New ADODB.Recordset
      sql = ("Select * from attn_status_mas")
      paydb.Open pay
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      While Not payrs.EOF
            lst_code.AddItem payrs(1)
            lst_code.ItemData(lst_code.ListCount - 1) = payrs(0)
            payrs.MoveNext
      Wend
      lst_code.ListIndex = -1
End Sub

Private Sub attn_dt_Click()
   Set paydb = New ADODB.Connection
   Set payrs = New ADODB.Recordset
   If emptype_chk = 0 Then
      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
   Else
      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
   End If
   paydb.Open pay
   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   If payrs.EOF Then
      MsgBox ("Pervious date details are missing. First enter for previous date & continue")
      attn_dt = attn_dt - 1
   End If
   If emptype_chk = 0 Then
      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
   Else
      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
   End If
   Set paydb = New ADODB.Connection
   Set payrs = New ADODB.Recordset
   paydb.Open pay
   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   If Not payrs.EOF Then
      While Not payrs.EOF
            With att_flex
                 .Rows = .Rows + 1
                 find_empdetails (payrs.Fields("attn_empcode"))
                 find_attnstatus (payrs.Fields("attn_status"))
                 .TextMatrix(i, 0) = dept_name
                 .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
                 .TextMatrix(i, 2) = employee_name
                  att_dat = attn_dt
                  find_present_status (payrs(0))
                 .TextMatrix(i, 3) = attstatus
                 .TextMatrix(i, 4) = payrs(5)
                 i = i + 1
                 endrow = endrow + 1
            End With
            payrs.MoveNext
      Wend
    Else
       MsgBox ("Details not available for the date ")
    End If
    new_entry_chk = 0
End Sub

Private Sub NEW_Click()
  new_entry_chk = 1
  attn_dt.SetFocus
End Sub

Private Sub edit_Click()
    endrow = 0
    fillgrid
    i = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Then
       sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
    Else
       sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
        While Not payrs.EOF
              With att_flex
                  .Rows = .Rows + 1
                  find_empdetails (payrs.Fields("attn_empcode"))
                  find_attnstatus (payrs.Fields("attn_status"))
                  .TextMatrix(i, 0) = dept_name
                  .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
                  .TextMatrix(i, 2) = employee_name
                  .TextMatrix(i, 3) = attendance_staus
                  .TextMatrix(i, 4) = payrs(5)
                  i = i + 1
                  endrow = endrow + 1
             End With
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
End Sub

Private Sub exit_Click()
   Unload Me
End Sub
 
Private Sub Form_Load()

  If emptype_chk = 0 Then
     cbeattn_entry.Caption = "Daily Attendacne Entry for STAFF"
  Else
     cbeattn_entry.Caption = "Daily Attendacne Entry for WORKER"
  End If
  new_entry_chk = 0
  attn_dt = Format(Now, "dd/mm/yyyy")
''  pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
  sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(attn_dt, "mm/dd/yyyy") & "'"
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  If Not payrs.EOF Then
     attstatus = payrs(1)
  Else
     attstatus = "PRESENT"
  End If
  endrow = 0
  fillgrid
  lst_code.Visible = False
  lst_name.Visible = False
  txt_itemname.Visible = False
  txt.Visible = False
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  If emptype_chk = 0 Then
     sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1) and emp_work = 'COIMBATORE' order by emp_name")
  Else
     sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3) and emp_work = 'COIMBATORE' order by emp_name")
  End If
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  payrs.MoveFirst
  While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs(0)
             .TextMatrix(.Rows - 1, 2) = payrs(5)
             If Trim((payrs.Fields("emp_holiday"))) = UCase(RTrim(Format(Now, "dddd"))) Then
                .TextMatrix(.Rows - 1, 3) = "WEEKLY OFF"
             End If
             att_dat = attn_dt
             find_present_status (payrs(0))
'
'             find_present_status (payrs(0),"01/01/2007")
'             attnstatus =
             'End If
             endrow = endrow + 1
        End With
        payrs.MoveNext
  Wend
  lst_code.Clear
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  sql = ("Select * from attn_status_mas")
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
' payrs.MoveFirst
  While Not payrs.EOF
        lst_code.AddItem payrs(1)
        lst_code.ItemData(lst_code.ListCount - 1) = payrs(0)
        payrs.MoveNext
  Wend
  lst_code.ListIndex = -1
End Sub

Private Sub att_flex_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler
 Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
 Select Case fin_selcol
        Case 4
            If KeyAscii <> 13 Then
               If emptype_chk = 0 Then Exit Sub
               KeyAscii = hours_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
            End If
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
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

Private Sub refresh_Click()
 Form_Load
End Sub

Private Sub SAVE_Click()
  If endrow = 0 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  If emptype_chk = 0 Then
     sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
  Else
     sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
  End If
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  While Not payrs.EOF
     payrs.Delete
     payrs.Update
     payrs.MoveNext
  Wend
  For i = 1 To endrow
      If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
            payrs.AddNew
            payrs.Fields("attn_company") = company_code
            payrs.Fields("attn_date") = Format(attn_dt, "dd/mm/yyyy")
            payrs.Fields("attn_empcode") = att_flex.TextMatrix(i, 1)
            find_empdetails (att_flex.TextMatrix(i, 1))
            payrs.Fields("attn_emptype") = emptypecode
            find_attn_status_code (att_flex.TextMatrix(i, 3))
            payrs.Fields("attn_status") = attendance_staus
            payrs.Fields("attn_exhr") = Val(att_flex.TextMatrix(i, 4))
            payrs.Fields("attn_pr_month") = Month(attn_dt) - 1
            payrs.Fields("attn_pr_year") = Year(attn_dt)
            payrs.Update
      End If
  Next
  MsgBox ("Records are saved")
  Form_Load
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
Dim ret%, pst_row%, pst_rawname$
 Static PrevIndex%
lst_code.Tag = "Keypress"
    Select Case KeyAscii
        Case 8
            If Trim(fst_item) <> "" Then fst_item = Mid(fst_item, 1, Len(fst_item) - 1)
        Case 13
             pbl_status = True
             If lst_code.ListIndex <> -1 Then pst_rawname = lst_code.Text
             For pin_cnt = 1 To att_flex.Rows - 1
                If pin_cnt <> att_flex.Row Then If LCase(att_flex.TextMatrix(pin_cnt, 3)) = LCase(pst_rawname) Then pbl_status = False
             Next
                pst_row = att_flex.Row
                If lst_code.ListIndex <> -1 Then
                    att_flex.TextMatrix(pst_row, 3) = lst_code.Text
                 att_flex.Col = 1
                 att_flex.SetFocus
                 lst_code.Tag = ""
                 Exit Sub
               End If
        Case Else
            fst_item = txt & Chr(KeyAscii)
            If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    End Select
    
    PrevIndex = lst_code.ListIndex
    If Trim(fst_item) = "" Then
        lst_code.ListIndex = -1
    Else
        ret = SendMessage(lst_code.hwnd, LB_FINDSTRING, -1, fst_item)
        If ret = -1 Then
            lst_code.ListIndex = PrevIndex
        Else
            lst_code.ListIndex = ret
        End If
    End If
    lst_code.Tag = ""
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
        lst_code.ListIndex = IIf(lst_code.ListIndex + 1 = lst_code.ListCount, lst_code.ListIndex, lst_code.ListIndex + 1)
        lst_code.SetFocus
    End If
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
      
Private Sub att_flex_EnterCell()
On Error GoTo err_handler
    Select Case att_flex.Col
        Case 3
            txt.Left = att_flex.Left + att_flex.CellLeft
            txt.Top = att_flex.Top + att_flex.CellTop
            txt.Width = att_flex.CellWidth - 15
            txt.Visible = True
            lst_code.Left = att_flex.Left + att_flex.CellLeft
            lst_code.Top = txt.Top + txt.Height
            lst_code.Width = att_flex.CellWidth
            lst_code.ListIndex = -1
            txt = att_flex.Text
            lst_code.Visible = True
            txt_itemname.Visible = False
            lst_name.Visible = False
            txt.SetFocus
        Case 4, 1, 2
            txt.Visible = False
            lst_code.Visible = False
    End Select
    Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub

Private Sub lst_code_DblClick()
On Error GoTo err_handler
     If lst_code.ListIndex <> -1 And lst_code.Tag = "" Then txt_KeyPress 13
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
Private Sub lst_code_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
     If KeyAscii = 13 Then lst_code_DblClick
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
