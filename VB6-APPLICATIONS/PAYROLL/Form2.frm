VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00FFC0FF&
   Caption         =   "PAYROLL SYSTEM "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Exit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6435
      TabIndex        =   7
      Top             =   5190
      Width           =   1395
   End
   Begin VB.ComboBox user_cmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3855
      TabIndex        =   6
      Top             =   3735
      Width           =   3975
   End
   Begin VB.CommandButton Continue 
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4620
      TabIndex        =   4
      Top             =   5175
      Width           =   1395
   End
   Begin VB.TextBox password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3870
      MaxLength       =   6
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   4380
      Width           =   1815
   End
   Begin VB.ComboBox Mill_cmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3900
      TabIndex        =   0
      Top             =   2970
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "User"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2415
      TabIndex        =   5
      Top             =   3825
      Width           =   915
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Password"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2445
      TabIndex        =   3
      Top             =   4545
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Mill Name"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2385
      TabIndex        =   2
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      Height          =   3780
      Left            =   2235
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   7470
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pay As String

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub mill_cmb_click()
    password.Text = ""
End Sub

Private Sub Continue_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("Select * from comp_mas where comp_name = '" & Mill_cmb.Text & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF() Then
       millname = Mill_cmb.Text
       If user_cmb.Text = "PAYROLL-USER" Then
          If password = payrs!comp_upw Then
             MsgBox ("User Passward is ok")
             Mainmenu.Show
          Else
             MsgBox ("User Password is wrong")
          End If
       Else
          If user_cmb.Text = "PAYROLL-HOD" Then
            If password = payrs!comp_hpw Then
               MsgBox ("HOD Passward is ok")
               Mainmenu.Show
            Else
               MsgBox ("HOD Password is wrong")
            End If
          End If
       End If
    End If
    
End Sub
Private Sub Form_Load()
    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
    paydb.Open pay
    sql = "Select * from comp_mas order by comp_code"
    payrs.Open sql, paydb, adOpenDynamic
    Mill_cmb.Text = payrs!comp_name
    While Not payrs.EOF
        Mill_cmb.AddItem (payrs!comp_name)
        payrs.MoveNext
    Wend
    user_cmb.AddItem ("PAYROLL-USER")
    user_cmb.AddItem ("PAYROLL-HOD")
    user_cmb.Text = "PAYROLL-USER"
End Sub




