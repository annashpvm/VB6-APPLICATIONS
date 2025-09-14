VERSION 5.00
Begin VB.Form frm_login 
   BackColor       =   &H00FFC0FF&
   Caption         =   "ATTENDANCE SYSTEM "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_password2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   13440
      MaxLength       =   10
      PasswordChar    =   "#"
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   3720
      Picture         =   "frm_login.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox password 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   6
         PasswordChar    =   "#"
         TabIndex        =   5
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox cmb_year 
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton Exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Exit"
         Height          =   405
         Left            =   2640
         TabIndex        =   3
         Top             =   2880
         Width           =   1395
      End
      Begin VB.CommandButton Continue 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Login"
         Height          =   405
         Left            =   960
         TabIndex        =   2
         Top             =   2880
         Width           =   1395
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1425
      Top             =   5205
   End
   Begin VB.Label payroll 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ATTENDANCE SYSTEM"
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
      Height          =   345
      Left            =   4680
      TabIndex        =   0
      Top             =   960
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      Height          =   795
      Left            =   12120
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    Set payrs2 = New ADODB.Recordset
    paydb.Open pay
    sql = ("Select * from emp_mas")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        payrs(3) = UCase(payrs(3))
        payrs.Update
        payrs.MoveNext
    Wend
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub mill_cmb_click()
    Password.Text = ""
End Sub

Private Sub Continue_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    
    Dim superuser, adminuser   As String
    superuser = ""
    adminuser = ""
    sql = "select * from mas_users  where usr_code =47"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF() Then
           superuser = payrs.Fields("usr_pwd")
    End If
    payrs.Close
    sql = "select * from mas_users  where usr_code =65"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF() Then
           adminuser = payrs.Fields("usr_pwd")
    End If
    payrs.Close
    
    
    casual = "N"
    fyear = cmb_year.Text
    finyear = cmb_year.ItemData(cmb_year.ListIndex)
    
    
    If Password.Text = superuser Then
       
       If UCase(txt_password2.Text) = adminuser Then
          adminpw = 1
       End If
       
       If UCase(txt_password2.Text) = "shvpm" Then
          adminpw = 2
       End If
       
       MAINMENU.Show
       frm_login.Visible = False
       Unload frm_login
    Else
       If Password.Text = "CASUAL" Then
          casual = "Y"
          MAINMENU.Show
          frm_login.Visible = False
          Unload frm_login
       Else
         MsgBox ("User Password is wrong")
       End If
   End If
    
End Sub
Private Sub Form_Load()
    adminpw = 0
    poweruser = 0
''  pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=10.0.0.252"
    If paydb.state = adStateOpen Then
       paydb.Close
    End If
    
   If paydb.state = adStateOpen Then
       paydb.Close
    End If
    paydb.Open pay
    sql = "select * from mas_finyear where fin_code >= 21 order by fin_year"
    payrs.Open sql, paydb, adOpenDynamic
    While Not payrs.EOF
          cmb_year.AddItem payrs!fin_year
          cmb_year.ItemData(cmb_year.NewIndex) = payrs!fin_code
          cmb_year.Text = payrs!fin_year
          payrs.MoveNext
    Wend
    ''cmb_year.ListIndex = 15
    
    
    Password.Text = "attn"
End Sub

Private Sub password_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Password.Text <> "" Then Continue_Click
End Sub



