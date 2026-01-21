VERSION 5.00
Begin VB.Form master_comany 
   Caption         =   "MILL WISE DETAILS"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton exit 
      Caption         =   "&Exit"
      Height          =   870
      Left            =   4650
      TabIndex        =   8
      Top             =   7215
      Width           =   975
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   870
      Left            =   3675
      TabIndex        =   7
      Top             =   7215
      Width           =   975
   End
   Begin VB.Frame DA_POINT 
      Caption         =   "DA POINT ENTRY"
      ForeColor       =   &H00FF0000&
      Height          =   4785
      Left            =   3075
      TabIndex        =   0
      Top             =   1725
      Width           =   7260
      Begin VB.TextBox mpaise 
         Height          =   525
         Left            =   2835
         TabIndex        =   6
         Top             =   3135
         Width           =   1050
      End
      Begin VB.TextBox mfda 
         Height          =   615
         Left            =   2865
         TabIndex        =   5
         Top             =   1875
         Width           =   3390
      End
      Begin VB.TextBox mpfno 
         Height          =   525
         Left            =   2850
         TabIndex        =   4
         Top             =   585
         Width           =   3345
      End
      Begin VB.Label Label3 
         Caption         =   "PAISE / RATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   345
         TabIndex        =   3
         Top             =   3300
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "FDA POINTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   345
         TabIndex        =   2
         Top             =   2100
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Mill PF Number :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1980
      End
   End
End
Attribute VB_Name = "master_comany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SAVE_Click()
     Set paydb = New ADODB.Connection
     Set payrs = New ADODB.Recordset
     sql = "select * from pay_master where m_company = '" & company_code & "'"
     paydb.Open pay
     payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
     payrs.Update
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
     Set paydb = New ADODB.Connection
     Set payrs = New ADODB.Recordset
     sql = "select * from pay_master where m_company = '" & company_code & "'"
     paydb.Open pay
     payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
     If Not payrs.EOF Then
        mpfno = payrs.Fields("m_comp_no")
        mfda = payrs.Fields("m_fdapoints")
        mpaise = payrs.Fields("m_rate")
     Else
        MsgBox ("Please Enter DA POINTS & RATE in master")
        Exit Sub
     End If
End Sub

Private Sub mfda_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii splall, "N", 10, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub


Private Sub mpaise_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii splall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
