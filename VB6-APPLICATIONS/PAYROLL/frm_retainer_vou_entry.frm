VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_retainer_vou_entry 
   Caption         =   "RETAINER / VOCCHER PAYMENT ENTRY"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   8655
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
         TabIndex        =   7
         Top             =   240
         Width           =   1335
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
         TabIndex        =   10
         Top             =   315
         Width           =   885
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
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   7440
      Width           =   4935
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   825
         Left            =   0
         MaskColor       =   &H000000FF&
         Picture         =   "frm_retainer_vou_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   825
         Left            =   960
         MaskColor       =   &H000000FF&
         Picture         =   "frm_retainer_vou_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   825
         Left            =   1920
         MaskColor       =   &H000000FF&
         Picture         =   "frm_retainer_vou_entry.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   825
         Left            =   2910
         MaskColor       =   &H000000FF&
         Picture         =   "frm_retainer_vou_entry.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   825
         Left            =   3900
         MaskColor       =   &H000000FF&
         Picture         =   "frm_retainer_vou_entry.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid ot_flex 
      Height          =   4770
      Left            =   720
      TabIndex        =   11
      Top             =   2040
      Width           =   9675
      _ExtentX        =   17066
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
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   10455
   End
   Begin VB.Label lbl_emp 
      Caption         =   "RETAINER / OTHERS  PAYMENT ENTRY"
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
      TabIndex        =   12
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm_retainer_vou_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Byte
Dim emp_cat As String
Function fillgrid()
    With ot_flex
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Amount"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
    End With
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "Select * from emp_voupay_mast where emp_company = '" & company_code & "' order by emp_doj"
    emp_cat = "W"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With ot_flex
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
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
    load_data
  End If
End Sub

Private Sub cmb_year_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     load_data
  End If
End Sub
Function load_data()
   If cmb_month.Text = "" And cmb_year.Text = "" Then
       MsgBox ("Select Month / Year ...")
       Exit Function
    End If
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from emp_otherpayment_entry  where ot_company = " & company_code & " and ot_finyear = " & finyear & " and ot_emp_cat = 'W' and ot_year = " & Val(cmb_year.Text) & " and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & ""
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To ot_flex.Rows - 1
                 If Trim(ot_flex.TextMatrix(i, 1)) <> "" Then
                      If ot_flex.TextMatrix(i, 1) = payrs.Fields("ot_emp_code") Then
                         ot_flex.TextMatrix(i, 3) = payrs.Fields("ot_amount")
                      End If
                End If
             Next
             payrs.MoveNext
        Wend
     End If
End Function

Private Sub NEW_Click()
  new_entry_chk = 0
  fillgrid
 ''ttn_dt.SetFocus
End Sub

Private Sub edit_Click()
    endrow = 0
    fillgrid
    load_data
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
   lbl_emp.Caption = "VOUCHER PAYMENT ENTRY"
   fillgrid
End Sub

Private Sub ot_flex_KeyPress(KeyAscii As Integer)
 If cmb_month.Text = "" Or cmb_year.Text = "" Then
    MsgBox ("Select Month / Year....")
    Exit Sub
 End If
 On Error GoTo err_handler
 Dim layoffdays As Double
 Dim fin_selrow%, fin_selcol%
 fin_selrow = ot_flex.Row
 fin_selcol = ot_flex.Col
 With ot_flex
 Select Case fin_selcol
        Case 3
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, ot_flex.TextMatrix(fin_selrow, fin_selcol), 8, 5, 2)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                ot_flex.TextMatrix(fin_selrow, fin_selcol) = ot_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
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
On Error GoTo err_handler
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
  sql = "delete from emp_otherpayment_entry  where ot_company = " & company_code & " and ot_finyear = " & finyear & " and ot_emp_cat = 'W' and ot_year = " & Val(cmb_year.Text) & " and ot_month = " & cmb_month.ItemData(cmb_month.ListIndex) & ""
  paydb.Execute sql
  For i = 1 To ot_flex.Rows - 1
      If Trim(ot_flex.TextMatrix(i, 3)) <> "" Then
         sql2 = "insert into emp_otherpayment_entry   values ( " & company_code & ", " & finyear & ", '" & ot_flex.TextMatrix(i, 1) & "' ,'W', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & Val(ot_flex.TextMatrix(i, 3)) & ")"
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
             For pin_cnt = 1 To ot_flex.Rows - 1
                If pin_cnt <> ot_flex.Row Then If LCase(ot_flex.TextMatrix(pin_cnt, 3)) = LCase(pst_rawname) Then pbl_status = False
             Next
                pst_row = ot_flex.Row
                If lst_code.ListIndex <> -1 Then
                    ot_flex.TextMatrix(pst_row, 3) = lst_code.Text
                 ot_flex.Col = 1
                 ot_flex.SetFocus
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
      
Private Sub ot_flex_EnterCell()
On Error GoTo err_handler
    Select Case ot_flex.Col
        Case 3
''            txt.Left = ot_flex.Left + ot_flex.CellLeft
''            txt.Top = ot_flex.Top + ot_flex.CellTop
''            txt.Width = ot_flex.CellWidth - 15
''            txt.Visible = True
''            lst_code.Left = ot_flex.Left + ot_flex.CellLeft
''            lst_code.Top = txt.Top + txt.Height
''            lst_code.Width = ot_flex.CellWidth
''            lst_code.ListIndex = -1
''            txt = ot_flex.Text
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




