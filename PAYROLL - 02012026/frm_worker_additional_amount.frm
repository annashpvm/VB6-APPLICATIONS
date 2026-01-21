VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_worker_additional_amount 
   Caption         =   "WORKER ADDITIONAL AMOUNT PAYMENT DETAILS"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2760
      TabIndex        =   13
      Top             =   7680
      Width           =   4935
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   825
         Left            =   3900
         MaskColor       =   &H000000FF&
         Picture         =   "frm_worker_additional_amount.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   825
         Left            =   2910
         MaskColor       =   &H000000FF&
         Picture         =   "frm_worker_additional_amount.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   825
         Left            =   1920
         MaskColor       =   &H000000FF&
         Picture         =   "frm_worker_additional_amount.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   825
         Left            =   960
         MaskColor       =   &H000000FF&
         Picture         =   "frm_worker_additional_amount.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   825
         Left            =   0
         MaskColor       =   &H000000FF&
         Picture         =   "frm_worker_additional_amount.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   8655
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
         Left            =   6120
         TabIndex        =   9
         Top             =   240
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
         Left            =   2280
         TabIndex        =   8
         Top             =   300
         Width           =   2655
      End
      Begin VB.TextBox txt_amount 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmd_apply 
         Caption         =   "APPLY FOR ALL"
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
         Left            =   4920
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
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
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "YEAR"
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
         Height          =   285
         Left            =   5445
         TabIndex        =   11
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "AMOUNT"
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
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59047937
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   4770
      Left            =   2040
      TabIndex        =   19
      Top             =   2280
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   8414
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   2
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
      Caption         =   "EMPLOYEE ADDITIONAL AMOUNT ENTRY"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   0
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      Height          =   5055
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   8655
   End
End
Attribute VB_Name = "frm_worker_additional_amount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Byte
Dim emp_cat As String
Function fillgrid()
    With att_flex
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Addn. Amt"
                
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        
    End With
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Then
''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1)order by emp_name")
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A'  order by emp_doj")
       emp_cat = "S"
    Else
''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3)order by emp_name")
        sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' order by emp_doj")
       emp_cat = "W"
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
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend

End Function

Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
    find_dates
End Sub

Private Sub cmd_apply_Click()
    Dim i As Integer
    For i = 1 To att_flex.Rows - 1
       att_flex.TextMatrix(i, 3) = txt_amount.Text
    Next
End Sub

Private Sub NEW_Click()
  new_entry_chk = 0
  fillgrid
End Sub

Private Sub edit_Click()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Then
       sql = "select * from employee_additional_amount where e_company = " & company_code & " and e_finyear = " & finyear & " and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_emp_cat = 'S'"
    Else
       sql = "select * from employee_additional_amount where e_company = " & company_code & " and e_finyear = " & finyear & " and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_emp_cat = 'W'"
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                      If att_flex.TextMatrix(i, 1) = payrs.Fields("e_emp_code") Then
                         att_flex.TextMatrix(i, 3) = payrs.Fields("e_amount")
                      End If
                End If
             Next
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
   new_entry_chk = 0
  If emptype_chk = 0 Then
''     millattn_entry.Caption = "Individual deduction for STAFF"
     lbl_emp.Caption = "STAFF ADDITIONAL AMOUNT ENTRY"
  Else
  ''   millattn_entry.Caption = "Individual deduction for WORKER"
     lbl_emp.Caption = "WORKER ADDITIONA AMOUNT ENTRY"
  End If
''  new_entry_chk = 0
''  attn_dt = Format(Now, "dd/mm/yyyy")
''  sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(attn_dt, "mm/dd/yyyy") & "'"
''  Set paydb = New ADODB.Connection
''  Set payrs = New ADODB.Recordset
''  paydb.Open pay
''  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''  If Not payrs.EOF Then
''     attstatus = payrs(1)
''  Else
''     attstatus = "PRESENT"
''  End If
''  endrow = 0
  fillgrid
''  lst_code.Visible = False
''  lst_name.Visible = False
''  txt_itemname.Visible = False
''txt.Visible = False
End Sub

Private Sub att_flex_KeyPress(KeyAscii As Integer)
 If cmb_month.Text = "" Or cmb_year.Text = "" Then
    MsgBox ("Select Month / Year....")
    Exit Sub
 End If
 On Error GoTo err_handler
 Dim layoffdays As Double
 Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
 Select Case fin_selcol
        Case 3
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 8, 5, 2)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              KeyAscii = 0
            End If
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
    fillgrid
    new_entry_chk = 0
End Sub
Private Sub SAVE_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
       MsgBox "Out Of Financial Year", vbInformation, "Message"
       Exit Sub
    End If
    Dim attn_Days As Double
    attn_Days = 0
    On Error GoTo err_handler
    If endrow = 0 Then
         MsgBox (" Details not available ")
         Exit Sub
    End If
    If cmb_month.Text = "" Then
         MsgBox (" Select Month ")
         Exit Sub
    End If
    If cmb_year.Text = "" Then
         MsgBox (" Select Year ")
         Exit Sub
    End If
    Me.MousePointer = 11
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    paydb.BeginTrans
    paydb.Execute sql
    If emptype_chk = 0 Then
       sql2 = "delete from employee_additional_amount where e_company = " & company_code & " and e_finyear = " & finyear & " and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_emp_cat = 'S'"
    Else
       sql2 = "delete from employee_additional_amount where e_company = " & company_code & " and e_finyear = " & finyear & " and e_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_year = " & Val(cmb_year.Text) & " and e_emp_cat = 'W'"
    End If
    paydb.Execute sql2
    
    For i = 1 To att_flex.Rows - 1
          If Val(att_flex.TextMatrix(i, 3)) > 0 Then
             If emptype_chk = 0 Then
                sql2 = "insert into employee_additional_amount values ( " & company_code & ", " & finyear & ", '" & att_flex.TextMatrix(i, 1) & "' ,'S', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & Val(att_flex.TextMatrix(i, 3)) & ")"
             Else
                sql2 = "insert into employee_additional_amount values ( " & company_code & ", " & finyear & ", '" & att_flex.TextMatrix(i, 1) & "' ,'W', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & Val(att_flex.TextMatrix(i, 3)) & ")"
             End If
             paydb.Execute sql2
          End If
    Next
    MsgBox ("Records are saved")
    paydb.CommitTrans
    fillgrid
    Me.MousePointer = 1
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
''            txt.Left = att_flex.Left + att_flex.CellLeft
''            txt.Top = att_flex.Top + att_flex.CellTop
''            txt.Width = att_flex.CellWidth - 15
''            txt.Visible = True
''            lst_code.Left = att_flex.Left + att_flex.CellLeft
''            lst_code.Top = txt.Top + txt.Height
''            lst_code.Width = att_flex.CellWidth
''            lst_code.ListIndex = -1
''            txt = att_flex.Text
''            lst_code.Visible = True
''            txt_itemname.Visible = False
''            lst_name.Visible = False
''            txt.SetFocus
        Case 4, 1, 2
'            txt.Visible = False
'            lst_code.Visible = False
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




Private Sub txt_amount_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii txt_amount, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Public Sub find_dates()
    If cmb_year.Text = "" Then Exit Sub
    Dim mdays, diff As Integer
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
    st_date = end_date - Day(end_date) + 1
End Sub



