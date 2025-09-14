VERSION 5.00
Begin VB.Form master_company 
   Caption         =   "MILL WISE DETAILS"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   20340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   6495
      Left            =   12360
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox mfda 
         Height          =   540
         Left            =   2520
         TabIndex        =   36
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox mpaise 
         Height          =   525
         Left            =   2520
         TabIndex        =   35
         Top             =   1920
         Width           =   2250
      End
      Begin VB.TextBox fdaamt 
         Height          =   555
         Left            =   6960
         TabIndex        =   34
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label Label2 
         Caption         =   "FDA POINTS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "PAISE / RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label label4 
         Caption         =   "FDA AMOUNT"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5040
         TabIndex        =   37
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1440
      TabIndex        =   10
      Top             =   8880
      Width           =   3975
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&xit"
         Height          =   870
         Left            =   2880
         Picture         =   "master_company.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   870
         Left            =   1920
         Picture         =   "master_company.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   870
         Left            =   960
         Picture         =   "master_company.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   870
         Left            =   0
         Picture         =   "master_company.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame DA_POINT 
      Caption         =   "MILL DETAILS ENTRY"
      ForeColor       =   &H00FF0000&
      Height          =   8445
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   11130
      Begin VB.Frame Frame3 
         Caption         =   "PF DETAILS"
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
         Height          =   1695
         Left            =   480
         TabIndex        =   24
         Top             =   3120
         Width           =   9495
         Begin VB.TextBox mpfno 
            Height          =   525
            Left            =   2520
            TabIndex        =   31
            Top             =   240
            Width           =   2265
         End
         Begin VB.TextBox txt_pf 
            Height          =   525
            Left            =   7200
            TabIndex        =   29
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox txt_pf_emplyr_contri 
            Height          =   525
            Left            =   7200
            TabIndex        =   26
            Top             =   960
            Width           =   945
         End
         Begin VB.TextBox txt_pf_emp_contri 
            Height          =   525
            Left            =   2520
            TabIndex        =   25
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label13 
            Caption         =   "MILL PF Number"
            ForeColor       =   &H00C00000&
            Height          =   405
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   2130
         End
         Begin VB.Label Label16 
            Caption         =   "PF Ceiling Lessthan"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   5280
            TabIndex        =   30
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label15 
            Caption         =   "Employer Contribution"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   5280
            TabIndex        =   28
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label Label14 
            Caption         =   "Employee Contribution"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1740
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ESI DETAILS"
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
         Height          =   2175
         Left            =   480
         TabIndex        =   15
         Top             =   5280
         Width           =   9495
         Begin VB.TextBox txt_eligibleamt 
            Height          =   525
            Left            =   7080
            TabIndex        =   23
            Top             =   360
            Width           =   1425
         End
         Begin VB.TextBox txt_emp1_contri 
            Height          =   525
            Left            =   2640
            TabIndex        =   20
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox txt_emp2_contri 
            Height          =   525
            Left            =   7080
            TabIndex        =   18
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox txt_esino 
            Height          =   525
            Left            =   2640
            TabIndex        =   17
            Top             =   360
            Width           =   2265
         End
         Begin VB.Label Label12 
            Caption         =   "ESI Eligible Lessthan"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   5160
            TabIndex        =   22
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label10 
            Caption         =   "Employer Contribution"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   5160
            TabIndex        =   21
            Top             =   1200
            Width           =   1740
         End
         Begin VB.Label Label11 
            Caption         =   "Employee Contribution"
            ForeColor       =   &H00C00000&
            Height          =   525
            Left            =   360
            TabIndex        =   19
            Top             =   1200
            Width           =   1860
         End
         Begin VB.Label Label9 
            Caption         =   "Mill ESI Number "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   405
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2220
         End
      End
      Begin VB.TextBox C_CODE 
         Height          =   525
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1080
         Width           =   2205
      End
      Begin VB.ComboBox comp_cmb 
         Height          =   360
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   7635
      End
      Begin VB.TextBox h_pw 
         Height          =   555
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1800
         Width           =   2235
      End
      Begin VB.TextBox u_pw 
         Height          =   540
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1800
         Width           =   2220
      End
      Begin VB.TextBox c_name 
         Height          =   405
         Left            =   2820
         TabIndex        =   1
         Top             =   570
         Width           =   7665
      End
      Begin VB.Label Label8 
         Caption         =   "COMPANY CODE"
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "HOD PASSWORD"
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   5280
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "USER PASSWORD"
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label Label5 
         Caption         =   "COMPANY NAME"
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   480
         TabIndex        =   6
         Top             =   705
         Width           =   2145
      End
   End
End
Attribute VB_Name = "master_company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newchk As Integer
Dim ccode As String

Private Sub C_CODE_KeyPress(KeyAscii As Integer)
  On Error GoTo err_handler
    chk_keyascii C_CODE, "N", 1, 0, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
Private Sub comp_cmb_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from comp_mas where rtrim(comp_name) = '" & comp_cmb & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       C_CODE = payrs(0)
       u_pw = payrs(2)
       h_pw = payrs(3)
       mpfno = payrs(4)
       mfda = payrs(5)
       fdaamt = payrs(6)
       mpaise = payrs(7)
       txt_esino.Text = payrs("comp_esi_no")
       txt_eligibleamt.Text = payrs("comp_esi_eligible")
       txt_emp1_contri.Text = payrs("comp_esi_emp1_contri")
       txt_emp2_contri.Text = payrs("comp_esi_emp2_contri")
       txt_pf_emp_contri.Text = payrs("comp_pf_emp1_contri")
       txt_pf_emplyr_contri.Text = payrs("comp_pf_emp2_contri")
       txt_pf.Text = payrs("comp_pf_eligible")
    End If
End Sub
Private Sub edit_Click()
     save.Visible = True
     c_name.Visible = False
     c_name.Text = ""
     comp_cmb.Visible = True
     u_pw = ""
     h_pw = ""
     mfda = ""
     mpfno = ""
     fdaamt = ""
     mpaise = ""
     Set paydb = New ADODB.Connection
     Set payrs = New ADODB.Recordset
     sql = "select * from comp_mas order by Comp_code"
     paydb.Open pay
     payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
     payrs.MoveFirst
     comp_cmb.Clear
     While Not payrs.EOF
        comp_cmb.AddItem payrs(1)
        payrs.MoveNext
     Wend
End Sub

Private Sub fdaamt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii fdaamt, "N", 10, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
Private Sub NEW_Click()
   c_name.Visible = True
   comp_cmb.Visible = False
   c_name.SetFocus
   c_name = ""
   C_CODE = ""
   u_pw = ""
   h_pw = ""
   mfda = ""
   mpfno = ""
   fdaamt = ""
   mpaise = ""
   newchk = 1
End Sub

Private Sub SAVE_Click()
     If c_name = " " Then
        MsgBox ("Enter Company name ..")
        c_name.SetFocus
        Exit Sub
     End If
     If C_CODE = " " Then
        MsgBox ("Enter Company Code ..")
        C_CODE.SetFocus
        Exit Sub
     End If
    If u_pw = " " Then
        MsgBox ("Enter User password ..")
        u_pw.SetFocus
        Exit Sub
     End If
     If h_pw = " " Then
        MsgBox ("Enter HOD password ..")
        h_pw.SetFocus
        Exit Sub
     End If
     If mpfno = " " Then
        MsgBox ("Enter Mill PF no ..")
        mpfno.SetFocus
        Exit Sub
     End If
     If mfda = " " Then
        MsgBox ("Enter FDA POINTS ...")
        mfda.SetFocus
        Exit Sub
     End If
     If fdaamt = " " Then
        MsgBox ("Enter FDA AMOUNT ..")
        fdaamt.SetFocus
        Exit Sub
     End If
     If mpaise = " " Then
        MsgBox ("Enter PAISE / POINTS ..")
        mpaise.SetFocus
        Exit Sub
     End If
     Set paydb = New ADODB.Connection
     Set payrs = New ADODB.Recordset
     sql = "select * from comp_mas where rtrim(comp_name) = '" & RTrim(comp_cmb.Text) & "'"
     paydb.Open pay
     payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
     If Not payrs.EOF Then
           payrs.Delete
           payrs.Update
     End If
     If c_name = "" Then c_name = comp_cmb.Text
     payrs.AddNew
     payrs.Fields("COMP_CODE") = C_CODE
     payrs.Fields("COMP_NAME") = c_name
     payrs.Fields("COMP_UPW") = u_pw
     payrs.Fields("COMP_HPW") = h_pw
     payrs.Fields("COMP_FDAPOINTS") = Val(mfda)
     payrs.Fields("COMP_FDAAMT") = Val(fdaamt)
     payrs.Fields("COMP_RATE") = Val(mpaise)
     payrs.Fields("COMP_PFNO") = mpfno
     payrs("comp_esi_no") = txt_esino.Text
     payrs("comp_esi_eligible") = txt_eligibleamt.Text
     payrs("comp_esi_emp1_contri") = txt_emp1_contri.Text
     payrs("comp_esi_emp2_contri") = txt_emp2_contri.Text
     payrs("comp_pf_eligible") = Val(txt_pf.Text)
     payrs("comp_pf_emp1_contri") = txt_pf_emp_contri.Text
     payrs("comp_pf_emp2_contri") = txt_pf_emplyr_contri.Text

       
       
     payrs.Update
     c_name = ""
     mpfno = 0
     mpaise = 0
     mfda = 0
     fdaamt = 0
     u_pw = ""
     h_pw = ""
     save.Visible = False
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ''pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
     Set paydb = New ADODB.Connection
     Set payrs = New ADODB.Recordset
     sql = "select * from comp_mas where comp_code = '" & company_code & "'"
     paydb.Open pay
     payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
     If Not payrs.EOF Then
        c_name = payrs.Fields("comp_name")
        C_CODE = payrs.Fields("comp_code")
        u_pw = payrs.Fields("comp_upw")
        h_pw = payrs.Fields("comp_hpw")
        mpfno = payrs.Fields("comp_pfno")
        mfda = payrs.Fields("comp_fdapoints")
        mpaise = payrs.Fields("comp_rate")
        fdaamt = payrs.Fields("comp_fdaamt")
       txt_esino.Text = payrs("comp_esi_no")
       txt_eligibleamt.Text = payrs("comp_esi_eligible")
       txt_emp1_contri.Text = payrs("comp_esi_emp1_contri")
       txt_emp2_contri.Text = payrs("comp_esi_emp2_contri")
       txt_pf_emp_contri.Text = payrs("comp_pf_emp1_contri")
       txt_pf_emplyr_contri.Text = payrs("comp_pf_emp2_contri")
     End If
     
     newchk = 0
End Sub

Private Sub mfda_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii mfda, "N", 10, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub


Private Sub mpaise_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii mpaise, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

