VERSION 5.00
Begin VB.Form master 
   Caption         =   "MASTER ENTRY "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   4095
      Begin VB.CommandButton Save 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&SAVE"
         Height          =   855
         Left            =   120
         Picture         =   "master.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton ADD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&ADD"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00C0E0FF&
         Picture         =   "master.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Exit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   855
         Left            =   3000
         Picture         =   "master.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Delete 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&DELETE"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2040
         Picture         =   "master.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Edit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&EDIT"
         Height          =   855
         Left            =   1080
         Picture         =   "master.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame masframe 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5010
      Left            =   375
      TabIndex        =   0
      Top             =   2160
      Width           =   11070
      Begin VB.TextBox txt_printingorder 
         Height          =   480
         Left            =   3840
         TabIndex        =   16
         Top             =   4200
         Width           =   1320
      End
      Begin VB.TextBox salary_day 
         Height          =   480
         Left            =   3855
         TabIndex        =   7
         Top             =   2925
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.ComboBox deduction_type 
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
         Left            =   3870
         TabIndex        =   5
         Top             =   3390
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.ComboBox mas_cmb 
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
         Left            =   3885
         TabIndex        =   3
         Top             =   690
         Visible         =   0   'False
         Width           =   6765
      End
      Begin VB.TextBox data 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3870
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1695
         Width           =   6720
      End
      Begin VB.Label Label1 
         Caption         =   "PRINTING ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   4320
         Width           =   3030
      End
      Begin VB.Label salary_type 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SALARY ADDITION / DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   390
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label ded_label 
         Caption         =   "DEDUCTION TYPE "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   375
         TabIndex        =   6
         Top             =   3420
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.Label cmblabel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   330
         TabIndex        =   4
         Top             =   660
         Width           =   3285
      End
      Begin VB.Label headname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   285
         TabIndex        =   2
         Top             =   1710
         Width           =   3285
      End
   End
End
Attribute VB_Name = "master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim code As Integer
''Dim pay As String
Dim opt As Byte
Dim savechk As Byte
Dim oldtext As String
Private Sub ADD_Click()
    cmblabel.Caption = "  "
    masframe.Caption = name1
    headname.Caption = name2
    mas_cmb.Visible = False
    data.Enabled = True
    Save.Visible = True
    savechk = 0
End Sub
Private Sub Delete_Click()
    If menuchk = 1 Then
       sql = ("Select * from  pdept_mas order by dept_code")
       sql2 = ("Select * from  pdept_mas where dept_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdept_mas where dept_name = '" & oldtext & "'")
       sql4 = "Select * from  pdept_mas"
       sql5 = ("Select * from  pdept_mas order by dept_name")
    End If
    If menuchk = 2 Then
       sql = ("Select * from   pemptype_mas order by dtype_code")
       sql2 = ("Select * from  pemptype_mas where dtype_name = '" & data.Text & "'")
       sql3 = ("Select * from  pemptype_mas where dtype_name = '" & oldtext & "'")
       sql4 = "Select * from   pemptype_mas "
       sql5 = ("Select * from  pemptype_mas order by dtype_name")
    End If
    If menuchk = 3 Then
       sql = ("Select * from   pdesi_mas order by pdesi_code")
       sql2 = ("Select * from  pdesi_mas where pdesi_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdesi_mas where pdesi_name = '" & oldtext & "'")
       sql4 = "Select * from   pdesi_mas "
       sql5 = ("Select * from  pdesi_mas order by pdesi_name")
    End If
    If menuchk = 4 Then
       sql = ("Select * from   pqly_mas order by pqly_code")
       sql2 = ("Select * from  pqly_mas where pqly_name = '" & data.Text & "'")
       sql3 = ("Select * from  pqly_mas where pqly_name = '" & oldtext & "'")
       sql4 = "Select * from   pqly_mas "
       sql5 = ("Select * from  pqly_mas order by pqly_name")
    End If
    If menuchk = 5 Then
       sql = ("Select * from   preli_mas order by preli_code")
       sql2 = ("Select * from  preli_mas where preli_name = '" & data.Text & "'")
       sql3 = ("Select * from  preli_mas where preli_name = '" & oldtext & "'")
       sql4 = "Select * from   preli_mas"
       sql5 = ("Select * from  preli_mas order by preli_name")
    End If
    If menuchk = 6 Then
       sql = ("Select * from   pcomm_mas order by pcomm_code")
       sql2 = ("Select * from  pcomm_mas where pcomm_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcomm_mas where pcomm_name = '" & oldtext & "'")
       sql4 = "Select * from   pcomm_mas"
       sql5 = ("Select * from  pcomm_mas order by pcomm_name")
    End If
    If menuchk = 7 Then
       sql = ("Select * from   pcast_mas order by pcast_code")
       sql2 = ("Select * from  pcast_mas where pcast_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcast_mas where pcast_name = '" & oldtext & "'")
       sql4 = "Select * from   pcast_mas"
       sql5 = ("Select * from  pcast_mas order by pcast_name")
    End If
    
     If menuchk = 8 Then
       sql = ("Select * from   pdedu_mas order by pdedu_code")
       sql2 = ("Select * from  pdedu_mas where pdedu_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdedu_mas where pdedu_name = '" & oldtext & "'")
       sql4 = "Select * from   pdedu_mas"
       sql5 = ("Select * from  pdedu_mas order by pdedu_name")
    End If
    If menuchk = 9 Then
       sql = ("Select * from   attn_status_mas order by attn_type_code")
       sql2 = ("Select * from  attn_status_mas where attn_type_name = '" & data.Text & "'")
       sql3 = ("Select * from  attn_status_mas where attn_type_name = '" & oldtext & "'")
       sql4 = "Select * from   attn_status_mas"
       sql5 = ("Select * from  attn_status_mas order by attn_type_name")
    End If
   
    savechk = 2
    mas_cmb.Clear
    ADD.Visible = True
    ADD.Enabled = False
    Edit.Enabled = False
    data.Enabled = True
    cmblabel.Caption = name3
    masframe.Caption = name5
    mas_cmb.Visible = True
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql5, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        mas_cmb.AddItem payrs(1)
        payrs.MoveNext
    Wend
End Sub
Private Sub edit_Click()
   If menuchk = 1 Then
       sql = ("Select * from  pdept_mas order by dept_code")
       sql2 = ("Select * from  pdept_mas where dept_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdept_mas where dept_name = '" & oldtext & "'")
       sql4 = "Select * from  pdept_mas"
       sql5 = ("Select * from  pdept_mas order by dept_name")
    End If
    If menuchk = 2 Then
       sql = ("Select * from   pemptype_mas order by dtype_code")
       sql2 = ("Select * from  pemptype_mas where dtype_name = '" & data.Text & "'")
       sql3 = ("Select * from  pemptype_mas where dtype_name = '" & oldtext & "'")
       sql4 = "Select * from   pemptype_mas "
       sql5 = ("Select * from  pemptype_mas order by dtype_name")
    End If
    If menuchk = 3 Then
       sql = ("Select * from   pdesi_mas order by pdesi_code")
       sql2 = ("Select * from  pdesi_mas where pdesi_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdesi_mas where pdesi_name = '" & oldtext & "'")
       sql4 = "Select * from   pdesi_mas "
       sql5 = ("Select * from  pdesi_mas order by pdesi_name")
    End If
    If menuchk = 4 Then
       sql = ("Select * from   pqly_mas order by pqly_code")
       sql2 = ("Select * from  pqly_mas where pqly_name = '" & data.Text & "'")
       sql3 = ("Select * from  pqly_mas where pqly_name = '" & oldtext & "'")
       sql4 = "Select * from   pqly_mas "
       sql5 = ("Select * from  pqly_mas order by pqly_name")
    End If
   If menuchk = 5 Then
       sql = ("Select * from   preli_mas order by preli_code")
       sql2 = ("Select * from  preli_mas where preli_name = '" & data.Text & "'")
       sql3 = ("Select * from  preli_mas where preli_name = '" & oldtext & "'")
       sql4 = "Select * from   preli_mas"
       sql5 = ("Select * from  preli_mas order by preli_name")
    End If
    If menuchk = 6 Then
       sql = ("Select * from   pcomm_mas order by pcomm_code")
       sql2 = ("Select * from  pcomm_mas where pcomm_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcomm_mas where pcomm_name = '" & oldtext & "'")
       sql4 = "Select * from   pcomm_mas"
       sql5 = ("Select * from  pcomm_mas order by pcomm_name")
    End If
    If menuchk = 7 Then
       sql = ("Select * from   pcast_mas order by pcast_code")
       sql2 = ("Select * from  pcast_mas where pcast_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcast_mas where pcast_name = '" & oldtext & "'")
       sql4 = "Select * from   pcast_mas"
       sql5 = ("Select * from  pcast_mas order by pcast_name")
    End If
    If menuchk = 8 Then
       sql = ("Select * from   pdedu_mas order by pdedu_code")
       sql2 = ("Select * from  pdedu_mas where pdedu_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdedu_mas where pdedu_name = '" & oldtext & "'")
       sql4 = "Select * from   pdedu_mas"
       sql5 = ("Select * from  pdedu_mas order by pdedu_name")
    End If
    If menuchk = 9 Then
       sql = ("Select * from   attn_status_mas order by attn_type_code")
       sql2 = ("Select * from  attn_status_mas where attn_type_name = '" & data.Text & "'")
       sql3 = ("Select * from  attn_status_mas where attn_type_name = '" & oldtext & "'")
       sql4 = "Select * from   attn_status_mas"
       sql5 = ("Select * from  attn_status_mas order by attn_type_name")
       salary_type.Visible = True
       salary_type.Visible = True
    End If
    mas_cmb.Clear
    savechk = 1
    ADD.Visible = False
    Save.Visible = True
    Edit.Enabled = False
    data.Enabled = True
    cmblabel.Caption = name3
    masframe.Caption = name4
    mas_cmb.Visible = True
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    payrs.Open sql5, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        mas_cmb.AddItem payrs(1)
        payrs.MoveNext
    Wend
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    mas_cmb.Visible = False
    If menuchk = 8 Then
       ded_label.Visible = True
       deduction_type.Visible = True
       deduction_type.AddItem ("COMMON DEDUCTIONS")
       deduction_type.AddItem ("STAFFS DEDUCTIONS")
       deduction_type.AddItem ("WORKERS DEDUCTIONS")
       deduction_type.AddItem ("OCCASINAL DEDUCTIONS")
    End If
    If menuchk = 9 Then
       salary_type.Visible = True
       salary_day.Visible = True
    End If
    masframe.Caption = name1
    headname.Caption = name2
    data.Enabled = False
    ADD.Enabled = True
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
End Sub
Private Sub mas_cmb_Click()
    data.Text = mas_cmb.Text
    oldtext = mas_cmb.Text
     If menuchk = 1 Then
       sql = ("Select * from  pdept_mas order by dept_code")
       sql2 = ("Select * from pdept_mas where dept_name = '" & data.Text & "'")
       sql3 = ("Select * from pdept_mas where dept_name = '" & oldtext & "'")
       sql4 = "Select * from  pdept_mas"
       sql5 = ("Select * from pdept_mas order by dept_name")
    End If
    If menuchk = 2 Then
       sql = ("Select * from   pemptype_mas order by dtype_code")
       sql2 = ("Select * from  pemptype_mas where dtype_name = '" & data.Text & "'")
       sql3 = ("Select * from  pemptype_mas where dtype_name = '" & oldtext & "'")
       sql4 = "Select * from   pemptype_mas "
       sql5 = ("Select * from  pemptype_mas order by dtype_name")
    End If
    If menuchk = 3 Then
       sql = ("Select * from   pdesi_mas order by pdesi_code")
       sql2 = ("Select * from  pdesi_mas where pdesi_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdesi_mas where pdesi_name = '" & oldtext & "'")
       sql4 = "Select * from   pdesi_mas "
       sql5 = ("Select * from  pdesi_mas order by pdesi_name")
    End If
    If menuchk = 4 Then
       sql = ("Select * from   pqly_mas order by pqly_code")
       sql2 = ("Select * from  pqly_mas where pqly_name = '" & data.Text & "'")
       sql3 = ("Select * from  pqly_mas where pqly_name = '" & oldtext & "'")
       sql4 = "Select * from   pqly_mas "
       sql5 = ("Select * from  pqly_mas order by pqly_name")
    End If
    If menuchk = 5 Then
       sql = ("Select * from   preli_mas order by preli_code")
       sql2 = ("Select * from  preli_mas where preli_name = '" & data.Text & "'")
       sql3 = ("Select * from  preli_mas where preli_name = '" & oldtext & "'")
       sql4 = "Select * from   preli_mas"
       sql5 = ("Select * from  preli_mas order by preli_name")
    End If
    If menuchk = 6 Then
       sql = ("Select * from   pcomm_mas order by pcomm_code")
       sql2 = ("Select * from  pcomm_mas where pcomm_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcomm_mas where pcomm_name = '" & oldtext & "'")
       sql4 = "Select * from   pcomm_mas"
       sql5 = ("Select * from  pcomm_mas order by pcomm_name")
    End If
    If menuchk = 7 Then
       sql = ("Select * from   pcast_mas order by pcast_code")
       sql2 = ("Select * from  pcast_mas where pcast_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcast_mas where pcast_name = '" & oldtext & "'")
       sql4 = "Select * from   pcast_mas"
       sql5 = ("Select * from  pcast_mas order by pcast_name")
    End If
    If menuchk = 8 Then
       sql = ("Select * from   pdedu_mas order by pdedu_code")
       sql2 = ("Select * from  pdedu_mas where pdedu_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdedu_mas where pdedu_name = '" & oldtext & "'")
       sql4 = "Select * from   pdedu_mas"
       sql5 = ("Select * from  pdedu_mas order by pdedu_name")
    End If
    If menuchk = 9 Then
       sql = ("Select * from   attn_status_mas order by attn_type_code")
       sql2 = ("Select * from  attn_status_mas where attn_type_name = '" & data.Text & "'")
       sql3 = ("Select * from  attn_status_mas where attn_type_name = '" & oldtext & "'")
       sql4 = "Select * from   attn_status_mas"
       sql5 = ("Select * from  attn_status_mas order by attn_type_name")
    End If
   
    If savechk = 2 Then
       opt = MsgBox("Are you sure want to delete", vbYesNo, "Press")
       If opt = 6 Then
            Set paydb = New ADODB.Connection
            Set payrs = New ADODB.Recordset
            paydb.Open pay
            payrs.Open sql2, paydb, adOpenDynamic, adLockOptimistic
            If Not payrs.EOF Then
               payrs.Delete
               payrs.Update
            End If
       End If
       ADD.Visible = True
       ADD.Enabled = True
       Edit.Visible = True
       Edit.Enabled = ture
       mas_cmb.Clear
       mas_cmb.Visible = False
       masframe.Caption = name1
       headname.Caption = name2
       data.Enabled = False
       cmblabel.Caption = "  "
       data.Text = "  "
    Else
       oldtext = mas_cmb.Text
       data.Text = mas_cmb.Text
       If menuchk = 8 Then
          Set paydb = New ADODB.Connection
          Set payrs = New ADODB.Recordset
          paydb.Open pay
          payrs.Open sql2, paydb, adOpenDynamic, adLockOptimistic
          If Not payrs.EOF Then
             If payrs(2) = 1 Then deduction_type.Text = "COMMON DEDUCTIONS"
             If payrs(2) = 2 Then deduction_type.Text = "STAFFS DEDUCTIONS"
             If payrs(2) = 3 Then deduction_type.Text = "WORKERS DEDUCTIONS"
             If payrs(2) = 4 Then deduction_type.Text = "OCCATIONAL DEDUCTIONS"
          End If
       End If
       If menuchk = 9 Then
          Set paydb = New ADODB.Connection
          Set payrs = New ADODB.Recordset
          paydb.Open pay
          payrs.Open sql2, paydb, adOpenDynamic, adLockOptimistic
          If Not payrs.EOF Then
             salary_day = payrs(2)
          End If
       End If
       If menuchk = 1 Or menuchk = 3 Then
          Set paydb = New ADODB.Connection
          Set payrs = New ADODB.Recordset
          paydb.Open pay
          payrs.Open sql2, paydb, adOpenDynamic, adLockOptimistic
          If Not payrs.EOF Then
             txt_printingorder.Text = payrs(2)
          End If
       End If
       
    End If
End Sub
Private Sub SAVE_Click()
    data.Text = UCase(Trim(data.Text))
    If menuchk = 1 Then
       sql = ("Select * from  pdept_mas order by dept_code")
       sql2 = ("Select * from  pdept_mas where dept_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdept_mas where dept_name = '" & oldtext & "'")
       sql4 = "Select * from  pdept_mas"
       sql5 = ("Select * from  pdept_mas order by dept_name")
    End If
    If menuchk = 2 Then
       sql = ("Select * from   pemptype_mas order by dtype_code")
       sql2 = ("Select * from  pemptype_mas where dtype_name = '" & data.Text & "'")
       sql3 = ("Select * from  pemptype_mas where dtype_name = '" & oldtext & "'")
       sql4 = "Select * from   pemptype_mas "
       sql5 = ("Select * from  pemptype_mas order by dtype_name")
    End If
    If menuchk = 3 Then
       sql = ("Select * from   pdesi_mas order by pdesi_code")
       sql2 = ("Select * from  pdesi_mas where pdesi_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdesi_mas where pdesi_name = '" & oldtext & "'")
       sql4 = "Select * from   pdesi_mas "
       sql5 = ("Select * from  pdesi_mas order by pdesi_name")
    End If
    If menuchk = 4 Then
       sql = ("Select * from   pqly_mas order by pqly_code")
       sql2 = ("Select * from  pqly_mas where pqly_name = '" & data.Text & "'")
       sql3 = ("Select * from  pqly_mas where pqly_name = '" & oldtext & "'")
       sql4 = "Select * from   pqly_mas "
       sql5 = ("Select * from  pqly_mas order by pqly_name")
    End If
    If menuchk = 5 Then
       sql = ("Select * from   preli_mas order by preli_code")
       sql2 = ("Select * from  preli_mas where preli_name = '" & data.Text & "'")
       sql3 = ("Select * from  preli_mas where preli_name = '" & oldtext & "'")
       sql4 = "Select * from   preli_mas"
       sql5 = ("Select * from  preli_mas order by preli_name")
    End If
    If menuchk = 6 Then
       sql = ("Select * from   pcomm_mas order by pcomm_code")
       sql2 = ("Select * from  pcomm_mas where pcomm_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcomm_mas where pcomm_name = '" & oldtext & "'")
       sql4 = "Select * from   pcomm_mas"
       sql5 = ("Select * from  pcomm_mas order by pcomm_name")
    End If
    If menuchk = 7 Then
       sql = ("Select * from   pcast_mas order by pcast_code")
       sql2 = ("Select * from  pcast_mas where pcast_name = '" & data.Text & "'")
       sql3 = ("Select * from  pcast_mas where pcast_name = '" & oldtext & "'")
       sql4 = "Select * from   pcast_mas"
       sql5 = ("Select * from  pcast_mas order by pcast_name")
    End If
    If menuchk = 8 Then
       sql = ("Select * from   pdedu_mas order by pdedu_code")
       sql2 = ("Select * from  pdedu_mas where pdedu_name = '" & data.Text & "'")
       sql3 = ("Select * from  pdedu_mas where pdedu_name = '" & oldtext & "'")
       sql4 = "Select * from   pdedu_mas"
       sql5 = ("Select * from  pdedu_mas order by pdedu_name")
    End If
    If menuchk = 9 Then
       sql = ("Select * from   attn_status_mas order by attn_type_code")
       sql2 = ("Select * from  attn_status_mas where attn_type_name = '" & data.Text & "'")
       sql3 = ("Select * from  attn_status_mas where attn_type_name = '" & oldtext & "'")
       sql4 = "Select * from   attn_status_mas"
       sql5 = ("Select * from  attn_status_mas order by attn_type_name")
    End If
   
    If savechk = 0 And Trim(data.Text) <> "" Then
       code = 1
       Set paydb = New ADODB.Connection
       Set payrs = New ADODB.Recordset
       paydb.Open pay
       payrs.Open sql2, paydb, adOpenDynamic, adLockOptimistic
       If payrs.EOF Then
           Set paydb = New ADODB.Connection
           Set payrs = New ADODB.Recordset
           paydb.Open pay
           payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
           While Not payrs.EOF()
                code = payrs(0) + 1
                payrs.MoveNext
           Wend
           Set paydb = New ADODB.Connection
           Set payrs = New ADODB.Recordset
           paydb.Open pay
           payrs.Open sql4, paydb, adOpenDynamic, adLockOptimistic
           payrs.AddNew
           payrs(0) = code
           payrs(1) = RTrim(data.Text)
           
           If menuchk = 1 Or menuchk = 3 Then
              payrs(2) = Val(txt_printingorder.Text)
           End If
           
           If menuchk = 8 Then
           
              If deduction_type.Text = "COMMON DEDUCTIONS" Then
                  payrs(2) = 1
              Else
                  If deduction_type.Text = "STAFFS DEDUCTIONS" Then
                     payrs(2) = 2
                  Else
                     If deduction_type.Text = "WORKERS DEDUCTIONS)" Then
                        payrs(2) = 3
                     Else
                        payrs(2) = 4
                     End If
                  End If
              End If
           End If
           If menuchk = 9 Then
              payrs(2) = salary_day
           End If
           payrs.Update
           salary_day = 0
           data.Text = " "
        Else
           MsgBox ("Data Already Available")
        End If
    Else
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        paydb.Open pay
        payrs.Open sql3, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF Then
           payrs(1) = Trim(data.Text)
           If menuchk = 8 Then
              If deduction_type.Text = "COMMON DEDUCTIONS" Then
                  payrs(2) = 1
              Else
                  If deduction_type.Text = "STAFFS DEDUCTIONS" Then
                     payrs(2) = 2
                  Else
                     If deduction_type.Text = "WORKERS DEDUCTIONS)" Then
                        payrs(2) = 3
                     Else
                        payrs(2) = 4
                     End If
                  End If
              End If
           End If
           If menuchk = 9 Then
              payrs(2) = salary_day
           End If
           If menuchk = 1 Or menuchk = 3 Then
              payrs(2) = Val(txt_printingorder.Text)
           End If
           salary_day = 0
           payrs.Update
        End If
    End If
    mas_cmb.Visible = False
    data.Enabled = False
    ADD.Enabled = True
    Edit.Enabled = True
    Save.Visible = False
    ADD.Visible = True
    data.Text = " "
    savechk = 0
    cmblabel.Caption = "  "
    mas_cmb.Visible = False
    masframe.Caption = name1
    headname.Caption = name2
End Sub
