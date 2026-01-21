VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dec_holiday 
   BackColor       =   &H00FFC0FF&
   Caption         =   "DECLARE HOLIDAY MASTER ENTRY"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   9690
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Height          =   3255
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   2415
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4260
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&xit"
      Height          =   975
      Left            =   3810
      Picture         =   "dec_holiday.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1185
   End
   Begin VB.CommandButton Edit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Edit"
      Height          =   975
      Left            =   2610
      Picture         =   "dec_holiday.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1185
   End
   Begin VB.CommandButton save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      Height          =   975
      Left            =   1410
      Picture         =   "dec_holiday.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1185
   End
   Begin VB.Frame dec_holiday_frame 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Declare Holiday details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2280
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   9240
      Begin VB.TextBox dec_holi_name 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1200
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dec_holiday 
         Height          =   615
         Left            =   3465
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59834369
         CurrentDate     =   37544
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Declare holiday name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   375
         TabIndex        =   4
         Top             =   1440
         Width           =   2910
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Select Declare holiday"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   405
         TabIndex        =   3
         Top             =   480
         Width           =   2160
      End
   End
End
Attribute VB_Name = "dec_holiday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub edit_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(dec_holiday, "mm/dd/yyyy") & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       dec_holi_name = payrs(1)
    End If
End Sub

Private Sub Form_Load()
    fillgrid
    Dim fdate, edata As Date
    If year(DATE) = year(gdt_finsdate) Then
       fdate = DateValue("01/01/" + Str(year(gdt_finsdate)))
       EDATE = DateValue("12/31/" + Str(year(gdt_finsdate)))
    Else
       fdate = DateValue("01/01/" + Str(year(gdt_finedate)))
       EDATE = DateValue("12/31/" + Str(year(gdt_finedate)))
    End If
    dec_holiday = DATE
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "Select * from emp_dec_holiday where emp_dec_holiday between '" & Format(fdate, "MM/dd/yyyy") & "' and '" & Format(EDATE, "MM/dd/yyyy") & "' order by emp_dec_holiday "
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = Format(payrs!emp_dec_holiday, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 2) = payrs!emp_dec_holiname
        payrs.MoveNext
        flx_data.Rows = flx_data.Rows + 1
        i = i + 1
    Wend
    payrs.Close
    
    
End Sub

Private Sub SAVE_Click()
    If Trim(dec_holi_name) = "" Then
       MsgBox ("Holiday name cannot be empty")
       Exit Sub
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(dec_holiday, "mm/dd/yyyy") & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
         payrs.Delete
         payrs.Update
    Wend
    payrs.AddNew
    payrs(0) = dec_holiday
    payrs(1) = dec_holi_name
    payrs.Update
    MsgBox ("Updated..")
    dec_holi_name.Text = " "
End Sub


Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 3
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Description"
     .ColWidth(0) = 1000
     .ColWidth(1) = 1800
     .ColWidth(2) = 4500
     
     .Redraw = True
   End With
End Sub


