VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form watten_ent 
   Caption         =   "Daly Attendance Entry"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker attn_dt 
      Height          =   555
      Left            =   9240
      TabIndex        =   6
      Top             =   705
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   979
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444929
      CurrentDate     =   37519
   End
   Begin VB.ListBox lst_name 
      Height          =   1620
      Left            =   4425
      TabIndex        =   4
      Top             =   5190
      Width           =   2415
   End
   Begin VB.TextBox txt_itemname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   630
      TabIndex        =   3
      Top             =   5955
      Width           =   1620
   End
   Begin VB.ListBox lst_code 
      Height          =   1620
      Left            =   1800
      TabIndex        =   2
      Top             =   5190
      Width           =   2580
   End
   Begin VB.TextBox txt 
      Height          =   465
      Left            =   675
      TabIndex        =   1
      Top             =   5175
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   5355
      Left            =   225
      TabIndex        =   0
      Top             =   1740
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      BackColorFixed  =   16776960
      BackColorSel    =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label DATE 
      Caption         =   "DATE"
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
      Height          =   405
      Left            =   7800
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "watten_ent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fst_item$
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
        .ColWidth(2) = 4500
        .ColWidth(3) = 2200
        .ColWidth(4) = 1300
    End With
End Function

Private Sub Form_Load()
  fillgrid
  lst_code.Visible = False
  lst_name.Visible = False
  txt_itemname.Visible = False
  txt.Visible = False
  pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  sql = ("Select * from  emp_mas order by emp_dept")
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  payrs.MoveFirst
  While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs(0)
             .TextMatrix(.Rows - 1, 2) = payrs(1)
             .TextMatrix(.Rows - 1, 3) = "PRESENT"
        End With
        payrs.MoveNext
  Wend
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  sql = ("Select * from attn_status_mas")
  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  payrs.MoveFirst
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
Private Sub txt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
Dim ret%, pst_row%, pst_rawname$
 Static PrevIndex%
lst_code.Tag = "Keypress"
    Select Case KeyAscii
        Case 8
            If Trim(fst_item) <> "" Then fst_item = Mid(fst_item, 1, Len(fst_item) - 1)
        Case 13
             'Dim pin_cnt As Integer
             'Dim pbl_status As Boolean
             pbl_status = True
             If lst_code.ListIndex <> -1 Then pst_rawname = lst_code.Text
             For pin_cnt = 1 To att_flex.Rows - 1
                If pin_cnt <> att_flex.Row Then If LCase(att_flex.TextMatrix(pin_cnt, 3)) = LCase(pst_rawname) Then pbl_status = False
             Next
'             If pbl_status = False Then
'                 MsgBox MSG1, vbOKOnly + vbExclamation, "Warning"
'             Else
                pst_row = att_flex.Row
                If lst_code.ListIndex <> -1 Then
                    att_flex.TextMatrix(pst_row, 3) = lst_code.Text
'                     att_flex.TextMatrix(pst_row, 4) = lst_code.ItemData(lst_code.ListIndex)
'                End If
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

