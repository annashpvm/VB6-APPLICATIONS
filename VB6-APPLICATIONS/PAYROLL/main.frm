VERSION 5.00
Begin VB.Form frm_login 
   BackColor       =   &H00C0E0FF&
   Caption         =   "PAYROLL SYSTEM "
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Exit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5280
      TabIndex        =   15
      Top             =   5400
      Width           =   1395
   End
   Begin VB.CommandButton Continue 
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3720
      TabIndex        =   14
      Top             =   5400
      Width           =   1395
   End
   Begin VB.TextBox password 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   8
      PasswordChar    =   "#"
      TabIndex        =   13
      Text            =   "SHVPM"
      Top             =   4680
      Width           =   3015
   End
   Begin VB.ComboBox cmb_year 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3960
      Width           =   2775
   End
   Begin VB.ComboBox user_cmb 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3240
      Width           =   3975
   End
   Begin VB.ComboBox Mill_cmb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "HEAD Password "
      Height          =   855
      Left            =   7560
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txt_password 
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
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "#"
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
      Begin VB.OptionButton opt_new 
         BackColor       =   &H00FFC0FF&
         Caption         =   "NEW"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opt_old 
         BackColor       =   &H00FFC0FF&
         Caption         =   "OLD"
         Height          =   375
         Left            =   525
         TabIndex        =   7
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Height          =   135
      Left            =   10560
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   495
      Begin VB.OptionButton opt_ho 
         BackColor       =   &H00FFC0FF&
         Caption         =   "HO"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt_mill 
         BackColor       =   &H00FFC0FF&
         Caption         =   "MILL"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opt_all 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ALL"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1425
      Top             =   5205
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Password"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fin Year"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "User"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mill Name"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label payroll 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PAYROLL SYSTEM"
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
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   1200
      Width           =   2445
   End
   Begin VB.Shape Shape1 
      Height          =   4515
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   8655
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
 password.Text = "HR"
End Sub

Private Sub Continue_Click()
    poweruser = 0
    Dim pst_sdate, pst_edate As String
    pst_sdate = "01/04/" + Mid(cmb_year, 1, 4)
    pst_edate = "31/03/" + Mid(cmb_year, 6, 4)
    gdt_finsdate = CDate(pst_sdate)
    gdt_finedate = CDate(pst_edate)
    uname = user_cmb.Text
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    
    Dim superuser As String
    If user_cmb.Text = "Systems" Then
        superuser = ""
        sql = "select * from mas_users  where usr_code =64"
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF() Then
             superuser = payrs.Fields("usr_pwd")
        End If
        
        payrs.Close
        sql = ("Select * from comp_mas where comp_name = '" & Mill_cmb.Text & "'")
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF() Then
           company_code = payrs.Fields("comp_code")
           millname = Mill_cmb.Text
           compname = Mill_cmb.Text
           fyear = cmb_year.Text
           finyear = cmb_year.ItemData(cmb_year.ListIndex)
           If opt_All.Value = True Then
             data_source = "A"
           ElseIf opt_mill.Value = True Then
             data_source = "M"
           Else
             data_source = "H"
           End If
           If user_cmb.Text = "HR-USER" Then
              If password = payrs!comp_upw Then
                 hod = False
                 MAINMENU.Show
                 frm_login.Visible = False
                 Unload frm_login
              Else
                 MsgBox ("User Password is wrong")
              End If
           ElseIf user_cmb.Text = "HR-HOD" Then
                 If password = payrs!comp_hpw Then
                    hod = True
                    MAINMENU.Show
                    frm_login.Visible = False
                    Unload frm_login
                 Else
                    MsgBox ("HOD Password is wrong")
                 End If

                
           ElseIf user_cmb.Text = "Systems" Then
                  If password = payrs!comp_upw And txt_password.Text = superuser Then
                      poweruser = 1
                      hod = True
                      MAINMENU.Show
                      frm_login.Visible = False
                      Unload frm_login
                   Else
                      MsgBox ("Password is wrong")
                   End If
                End If
          End If
    Else
         sql = ("Select * from comp_mas where comp_name = '" & Mill_cmb.Text & "'")
         payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
         If Not payrs.EOF() Then
            company_code = payrs.Fields("comp_code")
            millname = Mill_cmb.Text
            
            If company_code = 1 Then
               millname = "I"
            ElseIf company_code = 2 Then
               millname = "II"
            ElseIf company_code = 5 Then
               millname = "II-C"
            End If
            
            millname = Mill_cmb.Text
            compname = Mill_cmb.Text
            
            fyear = cmb_year.Text
            finyear = cmb_year.ItemData(cmb_year.ListIndex)
            If opt_All.Value = True Then
               data_source = "A"
            ElseIf opt_mill.Value = True Then
               data_source = "M"
            Else
               data_source = "H"
            End If
          End If
          payrs.Close
          
          sql = "select * from mas_users  where usr_name = '" & user_cmb.Text & "'"
          payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
          If Not payrs.EOF() Then
             If password = payrs!usr_pwd Then
                userrights = payrs!usr_code
                MAINMENU.Show
                frm_login.Visible = False
                Unload frm_login
             Else
                MsgBox ("Password Error. Try again..")
             End If
          End If
    End If
End Sub
Private Sub Form_Load()
    poweruser = 0
''  pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=shvpm"
    If paydb.state = adStateOpen Then
       paydb.Close
    End If
    paydb.Open pay
    sql = "Select * from comp_mas order by comp_code"
    payrs.Open sql, paydb, adOpenDynamic
''  Mill_cmb.Text = payrs!comp_name
    While Not payrs.EOF
        Mill_cmb.AddItem (payrs!comp_name)
        Mill_cmb.ItemData(Mill_cmb.NewIndex) = payrs!comp_code
        payrs.MoveNext
    Wend
    Mill_cmb.ListIndex = 0
    payrs.Close
    sql = "select * from mas_finyear  order by fin_year"
    payrs.Open sql, paydb, adOpenDynamic
    While Not payrs.EOF
          cmb_year.AddItem payrs!fin_year
          cmb_year.ItemData(cmb_year.NewIndex) = payrs!fin_code
          cmb_year.Text = payrs!fin_year
          payrs.MoveNext
    Wend
''    cmb_year.ListIndex = 5
    payrs.Close
    sql = "select * from mas_users where (usr_dept =17 or usr_dept =7 or(usr_dept =3 and usr_code=7)) and usr_name not like 'z%' order by usr_dept desc"
    payrs.Open sql, paydb, adOpenDynamic
    While Not payrs.EOF
          user_cmb.AddItem payrs!usr_name
          user_cmb.Text = payrs!usr_name
          payrs.MoveNext
    Wend
    user_cmb.ListIndex = 0
    payrs.Close
''    user_cmb.AddItem ("PAYROLL-USER")
''    user_cmb.AddItem ("PAYROLL-HOD")
''    user_cmb.Text = "PAYROLL-USER"
''    password.Text = "Hari"
End Sub

Private Sub opt_new_Click()
    Mill_cmb.Clear
    If paydb.state = adStateOpen Then
       paydb.Close
    End If
    paydb.Open pay
    sql = "Select * from comp_mas where comp_code in (1,2,3,5,8,90) order by comp_code"
    payrs.Open sql, paydb, adOpenDynamic
''  Mill_cmb.Text = payrs!comp_name
    While Not payrs.EOF
        Mill_cmb.AddItem (payrs!comp_name)
        Mill_cmb.ItemData(Mill_cmb.NewIndex) = payrs!comp_code
        payrs.MoveNext
    Wend
    Mill_cmb.ListIndex = 0
End Sub

Private Sub opt_old_Click()
    Mill_cmb.Clear
    If paydb.state = adStateOpen Then
       paydb.Close
    End If
    paydb.Open pay
    sql = "Select * from comp_mas order by comp_code"
    payrs.Open sql, paydb, adOpenDynamic
''  Mill_cmb.Text = payrs!comp_name
    While Not payrs.EOF
        Mill_cmb.AddItem (payrs!comp_name)
        Mill_cmb.ItemData(Mill_cmb.NewIndex) = payrs!comp_code
        payrs.MoveNext
    Wend
    Mill_cmb.ListIndex = 0

End Sub

Private Sub password_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And password.Text <> "" Then Continue_Click
End Sub

Private Sub Timer1_Timer()
    Static i As Integer
    If i = 0 Then
       payroll.ZOrder
       i = 1
    End If
End Sub

Private Sub user_cmb_Click()
       Frame3.Visible = False
       If user_cmb.Text = "Systems" Then
          Frame3.Visible = True
       End If

End Sub
