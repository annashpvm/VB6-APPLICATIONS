VERSION 5.00
Begin VB.Form frm_pass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1560
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3900
   Begin VB.CommandButton cmd_enter 
      Caption         =   "Enter ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   2040
   End
   Begin VB.TextBox txt_pass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   270
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Left            =   165
      TabIndex        =   0
      Top             =   300
      Width           =   1635
   End
End
Attribute VB_Name = "frm_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_enter_Click()
    Dim pwd As Variant
    Dim pst_qry
    Dim rs_set As New ADODB.Recordset
    pwd = Trim(txt_pass.Text)
    
    If pwchk = 1 Then
        pst_qry = "select * from mas_users where usr_code = 45"
    ElseIf pwchk = 2 Then
        pst_qry = "select * from mas_users where usr_code = 45"
    ElseIf pwchk = 5 Then
        pst_qry = "select * from mas_users where usr_code = 45"
    
    ElseIf pwchk = 3 Or pwchk = 4 Then
        pst_qry = "select * from mas_users where usr_code = 46"
    End If
    rs_set.Open pst_qry, paydb, 1, 2
    If rs_set.EOF = False Then
        If Trim(rs_set!usr_pwd) <> Trim(pwd) Then
           rs_set.Close
           MsgBox "Invalid Password"
           txt_pass.Text = vbNullString
           txt_pass.SetFocus
           Exit Sub
        End If
    End If
    rs_set.Close
    cmd_enter.Caption = "Loading...."
    Me.MousePointer = 11
    If pwchk = 1 Then
       emp_mas_entry.Show
       emp_mas_entry.ZOrder
    ElseIf pwchk = 2 Then
       optchk = 2
       emp_mas_modify_new.Show
       emp_mas_modify_new.ZOrder
    ElseIf pwchk = 3 Then
       optchk = 3
       pay_slip_print.Show
       pay_slip_print.ZOrder
    ElseIf pwchk = 4 Then
       optchk = 3
       Salary_statement_prt.Show
       Salary_statement_prt.ZOrder
    ElseIf pwchk = 5 Then
       Load frm_mas_vou_payment
       frm_mas_vou_payment.ZOrder
       frm_mas_vou_payment.Show
    End If
    Me.MousePointer = 1
    Unload Me
End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_enter_Click
End Sub
