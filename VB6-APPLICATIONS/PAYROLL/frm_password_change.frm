VERSION 5.00
Begin VB.Form frm_password_change 
   Caption         =   "PASSWORD CHANGE SCREEN"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   12615
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   855
      Left            =   4560
      TabIndex        =   11
      Top             =   6480
      Width           =   2055
      Begin VB.CommandButton emp_save 
         Caption         =   "&Update"
         Height          =   735
         Left            =   120
         Picture         =   "frm_password_change.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   1080
         Picture         =   "frm_password_change.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   8775
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   1200
         TabIndex        =   4
         Top             =   2400
         Width           =   6135
         Begin VB.TextBox txt_new_pw2 
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
            IMEMode         =   3  'DISABLE
            Left            =   2400
            MaxLength       =   8
            PasswordChar    =   "#"
            TabIndex        =   8
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txt_new_pw1 
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
            IMEMode         =   3  'DISABLE
            Left            =   2400
            MaxLength       =   8
            PasswordChar    =   "#"
            TabIndex        =   7
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txt_old_pw 
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
            IMEMode         =   3  'DISABLE
            Left            =   2400
            MaxLength       =   8
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "NEW PASSWORD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "NEW PASSWORD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "OLD PASSWORD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   6135
         Begin VB.OptionButton opt3 
            Caption         =   "Locking "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   480
            TabIndex        =   15
            Top             =   960
            Width           =   4335
         End
         Begin VB.OptionButton opt2 
            Caption         =   "Salary Statement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   600
            Width           =   4335
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Employee Master"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   480
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   4335
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "PASSWORD CHANGE OPTOIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frm_password_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub emp_save_Click()
    If txt_old_pw.Text = "" Then
       MsgBox ("Old Password is empty...")
       Exit Sub
    End If
    If txt_new_pw1.Text = "" Then
       MsgBox ("New Password is empty...")
       Exit Sub
    End If
    If txt_new_pw2.Text = "" Then
       MsgBox ("New Password is empty...")
       Exit Sub
    End If
    If Trim(txt_new_pw1.Text) <> Trim(txt_new_pw2.Text) Then
           MsgBox "New Pasword Not matched... "
           txt_new_pw1.SetFocus
           Exit Sub
    End If
    Dim pst_qry
    Dim rs_set As New ADODB.Recordset
    ''pwd = Trim(txt_pass.Text)
    If opt1.Value = True Then
        pst_qry = "select * from mas_users where usr_code = 45"
    ElseIf opt2.Value = True Then
        pst_qry = "select * from mas_users where usr_code = 46"
    ElseIf opt3.Value = True Then
        pst_qry = "select * from mas_users where usr_code = 47"
    End If
    rs_set.Open pst_qry, paydb, 1, 2
    If rs_set.EOF = False Then
        If Trim(rs_set!usr_pwd) <> Trim(txt_old_pw.Text) Then
           rs_set.Close
           MsgBox "Old Password is invalid"
           txt_old_pw.Text = vbNullString
           txt_old_pw.SetFocus
           Exit Sub
        End If
    End If
    rs_set.Close
    If opt1.Value = True Then
        pst_qry = "update mas_users set usr_pwd = '" & Trim(txt_new_pw1.Text) & "' where usr_code = 45"
    ElseIf opt2.Value = True Then
        pst_qry = "update mas_users set usr_pwd = '" & Trim(txt_new_pw1.Text) & "' where usr_code = 46"
    ElseIf opt3.Value = True Then
        pst_qry = "update mas_users set usr_pwd = '" & Trim(txt_new_pw1.Text) & "' where usr_code = 47"
    End If
    
    paydb.Execute pst_qry
    Me.MousePointer = 1
    MsgBox ("Password updated...")
    txt_new_pw1.Text = ""
    txt_new_pw2.Text = ""
    txt_old_pw.Text = ""
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub txt_new_pw2_KeyPress(KeyAscii As Integer)
    If txt_new_pw2.Text <> "" And KeyAscii = 13 Then emp_save_Click
End Sub
