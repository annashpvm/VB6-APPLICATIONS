VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form VDA_ENTRY 
   BackColor       =   &H00FFFFC0&
   Caption         =   "VDA ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   1080
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   39359
      End
      Begin VB.Label Label10 
         Caption         =   "Report From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame vda_frame 
      BackColor       =   &H00C0E0FF&
      Caption         =   "VARIABLE DEARNESS ALLOWANCE ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5670
      Left            =   1365
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   9330
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Height          =   1095
         Left            =   960
         TabIndex        =   12
         Top             =   3840
         Width           =   3615
         Begin VB.CommandButton exit2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Exit"
            Height          =   855
            Left            =   2400
            Picture         =   "VDA_ENTRY.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   135
            Width           =   1110
         End
         Begin VB.CommandButton refresh 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Refresh"
            Height          =   855
            Left            =   1260
            Picture         =   "VDA_ENTRY.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1110
         End
         Begin VB.CommandButton Save 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Save"
            Height          =   855
            Left            =   120
            Picture         =   "VDA_ENTRY.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1110
         End
      End
      Begin VB.TextBox txt_dapoints 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2475
         TabIndex        =   3
         Top             =   1410
         Width           =   2235
      End
      Begin VB.TextBox fda_amount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3255
         TabIndex        =   4
         Top             =   2625
         Width           =   1650
      End
      Begin VB.ComboBox cmb_year 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7080
         TabIndex        =   2
         Top             =   675
         Width           =   1335
      End
      Begin VB.ComboBox cmb_month 
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
         Left            =   2475
         TabIndex        =   1
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox vda_amount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7215
         TabIndex        =   5
         Top             =   2670
         Width           =   1785
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "DA POINTS"
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
         Height          =   555
         Left            =   720
         TabIndex        =   11
         Top             =   1485
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "FDA AMOUNT"
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
         Left            =   615
         TabIndex        =   9
         Top             =   2670
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "MONTH"
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
         Height          =   330
         Left            =   720
         TabIndex        =   8
         Top             =   735
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "YEAR"
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
         Height          =   285
         Left            =   5955
         TabIndex        =   7
         Top             =   750
         Width           =   885
      End
      Begin VB.Line Line1 
         X1              =   -150
         X2              =   9165
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "VDA AMOUNT"
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
         Left            =   5235
         TabIndex        =   6
         Top             =   2730
         Width           =   1560
      End
      Begin VB.Line Line2 
         X1              =   15
         X2              =   9330
         Y1              =   3450
         Y2              =   3450
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FDA AMOUNT"
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
      Left            =   3390
      TabIndex        =   10
      Top             =   2580
      Width           =   1560
   End
End
Attribute VB_Name = "VDA_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fdapoint As Currency
Dim darate As Currency
Dim daamt As Currency
Dim fdaamt As Currency

Private Sub txt_dapoints_Change()
    daamt = Round((Val(txt_dapoints.Text) - fdapoint) * darate, 0)
    If daamt < 0 Then daamt = 0
    vda_amount = daamt
End Sub

Private Sub txt_dapoints_KeyDown(KeyCode As Integer, Shift As Integer)
    daamt = (Val(txt_dapoints) - fdapoint) * darate
    If daamt < 0 Then daamt = 0
    vda_amount = daamt
End Sub

Private Sub txt_dapoints_KeyPress(KeyAscii As Integer)
    daamt = (Val(txt_dapoints) - fdapoint) * darate
    If daamt < 0 Then daamt = 0
    vda_amount = daamt
    On Error GoTo err_handler
       chk_keyascii txt_dapoints, "N", 8, 2, KeyAscii
       Exit Sub
err_handler:
       chk = gen_Validation(Err.Number, Err.Description)
       If chk = 1 Then Resume
End Sub

Private Sub exit2_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     vda_frame.Visible = True
     With cmb_month
        .AddItem "January"
        .ItemData(.NewIndex) = 1
        .AddItem "February"
        .ItemData(.NewIndex) = 2
        .AddItem "March"
        .ItemData(.NewIndex) = 3
        .AddItem "April"
        .ItemData(.NewIndex) = 4
        .AddItem "May"
        .ItemData(.NewIndex) = 5
        .AddItem "June"
        .ItemData(.NewIndex) = 6
        .AddItem "July"
        .ItemData(.NewIndex) = 7
        .AddItem "August"
        .ItemData(.NewIndex) = 8
        .AddItem "September"
        .ItemData(.NewIndex) = 9
        .AddItem "October"
        .ItemData(.NewIndex) = 10
        .AddItem "November"
        .ItemData(.NewIndex) = 11
        .AddItem "December"
        .ItemData(.NewIndex) = 12
    End With
''    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''    End With
''    cmb_year.Text = "2015"
    
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    
    ''pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from comp_mas where comp_code = '" & company_code & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
        mpfno = payrs.Fields("comp_pfno")
        fdapoint = payrs.Fields("comp_fdapoints")
        darate = payrs.Fields("comp_rate")
        fda_amount = payrs.Fields("comp_fdaamt")
    Else
        MsgBox ("Please Enter DA POINTS & RATE in DA Master")
        Exit Sub
    End If
End Sub

Private Sub refresh_Click()
     txt_dapoints = " "
      cmb_month = ""
End Sub

Private Sub SAVE_Click()
    If v_vdaamount < 0 Then Exit Sub
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_vda where v_year = " & Trim(cmb_year.Text) & " and v_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and v_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       payrs.Fields("v_vdaamount") = vda_amount
       payrs.Fields("v_dapoints") = txt_dapoints.Text
       payrs.Update
    Else
       payrs.AddNew
       payrs.Fields("v_company") = company_code
       payrs.Fields("v_month") = cmb_month.ItemData(cmb_month.ListIndex)
       payrs.Fields("v_year") = Trim(cmb_year.Text)
       payrs.Fields("v_vdaamount") = vda_amount
       payrs.Fields("v_fdaamount") = fda_amount
       payrs.Fields("v_dapoints") = Val(txt_dapoints.Text)
       payrs.Update
    End If
    MsgBox ("VDA amount updated...")
    vda_amount = 0
    fda_amount = 0
   txt_dapoints = 0
End Sub

Private Sub cmb_month_Click()
   find_dates
   txt_dapoints = ""
    If Trim(cmb_month.Text) = "" Then
       MsgBox ("Select month")
       Exit Sub
    End If
    If Trim(cmb_year.Text) = "" Then
       MsgBox ("Select year")
       Exit Sub
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_vda where v_year = " & Trim(cmb_year.Text) & " and v_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and v_company = " & company_code & "")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       fda_amount.Text = payrs.Fields("v_fdaamount")
       vda_amount = payrs.Fields("v_vdaamount")
       txt_dapoints = payrs.Fields("v_dapoints")
    Else
       vda_amount = 0
    End If
End Sub



Private Sub cmb_year_Click()
    find_dates
    If Trim(cmb_year.Text) = "" Then
       MsgBox ("Select year")
       Exit Sub
    End If
    If Trim(cmb_month.Text) = "" Then
       MsgBox ("Select Month")
       Exit Sub
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_vda where v_year = " & Trim(cmb_year.Text) & " and v_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and v_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       fda_amount.Text = payrs.Fields("v_fdaamount")
       vda_amount = payrs.Fields("v_vdaamount")
       txt_dapoints = payrs.Fields("v_dapoints")
    Else
       vda_amount = 0
      txt_dapoints = 0
    End If
End Sub

Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
    Dim mdays, diff As Integer
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If cmb_year.Text = "" Then Exit Sub
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
End Sub


