VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_bank_loan_monthly_deduction 
   Caption         =   "BANK LOAN DEDUCTIONS"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   360
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
      Left            =   1920
      TabIndex        =   0
      Top             =   7080
      Width           =   4935
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   825
         Left            =   0
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bank_loan_monthly_deduction.frx":0000
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
         Picture         =   "frm_bank_loan_monthly_deduction.frx":066A
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
         Picture         =   "frm_bank_loan_monthly_deduction.frx":0CD4
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
         Picture         =   "frm_bank_loan_monthly_deduction.frx":133E
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
         Picture         =   "frm_bank_loan_monthly_deduction.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid bk_flex 
      Height          =   4770
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   8414
      _Version        =   393216
      Rows            =   3
      Cols            =   5
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
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   10455
   End
   Begin VB.Label lbl_emp 
      Caption         =   "EMPLOYEE BANK LOAN DEDUCTIONS"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frm_bank_loan_monthly_deduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Byte
Dim emp_cat As String
Function fillgrid()
    With bk_flex
        .Clear
        .Cols = 6
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "CAN Loan"
        .TextMatrix(0, 4) = "FEST Loan"
        .TextMatrix(0, 5) = "Oth Loan"
                
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        
    End With
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A'  and (emp_canbud_ac <> '' or emp_festloan_ac <> '' or emp_loan_ac <> '' )  order by emp_doj"
    emp_cat = "W"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   '' payrs.MoveFirst
    While Not payrs.EOF
        With bk_flex
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

Private Sub NEW_Click()
  new_entry_chk = 0
  fillgrid
 ''ttn_dt.SetFocus
End Sub

Private Sub edit_Click()
    endrow = 0
    fillgrid
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from  emp_monthly_bank_deduction  where bk_company = " & company_code & " and bk_finyear = " & finyear & " and bk_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and bk_year = " & Val(cmb_year.Text) & " and bk_emp_cat = 'W'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 2 To bk_flex.Rows - 1
                 If Trim(bk_flex.TextMatrix(i, 1)) <> "" Then
                      If bk_flex.TextMatrix(i, 1) = payrs.Fields("bk_emp_code") Then
                            bk_flex.TextMatrix(i, 3) = IIf(payrs.Fields("bk_loan1") > 0, payrs.Fields("bk_loan1"), "")
                            bk_flex.TextMatrix(i, 4) = IIf(payrs.Fields("bk_loan2") > 0, payrs.Fields("bk_loan2"), "")
                            bk_flex.TextMatrix(i, 5) = IIf(payrs.Fields("bk_loan3") > 0, payrs.Fields("bk_loan3"), "")
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
   fillgrid
End Sub

Private Sub bk_flex_KeyPress(KeyAscii As Integer)
 If cmb_month.Text = "" Or cmb_year.Text = "" Then
    MsgBox ("Select Month / Year....")
    Exit Sub
 End If
 On Error GoTo err_handler
 Dim layoffdays As Double
 Dim fin_selrow%, fin_selcol%
 fin_selrow = bk_flex.Row
 fin_selcol = bk_flex.Col
 With bk_flex
 Select Case fin_selcol
        Case 3, 4, 5
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, bk_flex.TextMatrix(fin_selrow, fin_selcol), 8, 5, 2)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                bk_flex.TextMatrix(fin_selrow, fin_selcol) = bk_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
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
  sql = "delete from emp_monthly_bank_deduction where bk_company = " & company_code & " and bk_finyear = " & finyear & " and bk_emp_cat = 'W' and bk_year = " & Val(cmb_year.Text) & " and bk_month = " & cmb_month.ItemData(cmb_month.ListIndex) & ""
  paydb.Execute sql
  For i = 1 To bk_flex.Rows - 1
      If Trim(bk_flex.TextMatrix(i, 5)) <> "" Then
         sql2 = "insert into emp_monthly_bank_deduction  values ( " & company_code & ", " & finyear & ", '" & bk_flex.TextMatrix(i, 1) & "' ,'W', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & Val(bk_flex.TextMatrix(i, 3)) & "," & Val(bk_flex.TextMatrix(i, 4)) & ", " & Val(bk_flex.TextMatrix(i, 5)) & ")"
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
             For pin_cnt = 1 To bk_flex.Rows - 1
                If pin_cnt <> bk_flex.Row Then If LCase(bk_flex.TextMatrix(pin_cnt, 3)) = LCase(pst_rawname) Then pbl_status = False
             Next
                pst_row = bk_flex.Row
                If lst_code.ListIndex <> -1 Then
                    bk_flex.TextMatrix(pst_row, 3) = lst_code.Text
                 bk_flex.Col = 1
                 bk_flex.SetFocus
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
      
Private Sub bk_flex_EnterCell()
On Error GoTo err_handler
    Select Case bk_flex.Col
        Case 3
''            txt.Left = bk_flex.Left + bk_flex.CellLeft
''            txt.Top = bk_flex.Top + bk_flex.CellTop
''            txt.Width = bk_flex.CellWidth - 15
''            txt.Visible = True
''            lst_code.Left = bk_flex.Left + bk_flex.CellLeft
''            lst_code.Top = txt.Top + txt.Height
''            lst_code.Width = bk_flex.CellWidth
''            lst_code.ListIndex = -1
''            txt = bk_flex.Text
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



