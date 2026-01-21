VERSION 5.00
Begin VB.Form frm_login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Login Form"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17025
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frm_login.frx":000C
   ScaleHeight     =   9060
   ScaleWidth      =   17025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmb_finyear 
      Height          =   315
      Left            =   8400
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton Continue 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7440
      TabIndex        =   3
      Top             =   8040
      Width           =   1395
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   2
      Top             =   8040
      Width           =   1395
   End
   Begin VB.TextBox txt_pw 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txt_user 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8760
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "SRI HARI VENAKTESWARA PAPER MILLS (P) Ltd.,"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   6960
      Picture         =   "frm_login.frx":264B6
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   4725
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset

Private Sub cmd_login_Click()

End Sub

Private Sub Continue_Click()

          frm_mainmenu.Show
          frm_login.Visible = False
          Unload frm_login
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()


    Call gen_dbconnection
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    


    pst_qry = "select fin_code,fin_year from mas_finyear where fin_code >=24 order by fin_code "

    cmb_finyear.Clear


    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_finyear.AddItem (adors("fin_year"))
                 cmb_finyear.ItemData(cmb_finyear.NewIndex) = adors("fin_code")
                 cmb_finyear.Text = adors("fin_year")
                 fincode = adors("fin_code")
                 gin_finid = adors("fin_code")
                 adors.MoveNext
        Next
    End If
    adors.Close


''txt.Text = ":7125307125340:7125307125340:"
''Dim cnt, cnt2 As Integer
''Dim findchar As String
''cnt = InStr(txt.Text, ":")
''MsgBox (Mid(txt.Text, cnt + 1))
''findchar = Mid(txt.Text, cnt + 1)
''cnt2 = InStr(findchar, ":")
''findchar = Mid(findchar, 1, cnt2 - 3)
''If Len(findchar) = 5 Then
''   MsgBox (Mid(findchar, 3))
''ElseIf Len(findchar) = 7 Then
''   MsgBox (Mid(findchar, 4))
''ElseIf Len(findchar) = 9 Then
''   MsgBox (Mid(findchar, 5))
''ElseIf Len(findchar) = 11 Then
''   MsgBox (Mid(findchar, 6))
''
''End If
''

''    MsgBox (Chr(185))

''      MsgBox (Chr(172))
''            MsgBox (Chr(181))
End Sub
