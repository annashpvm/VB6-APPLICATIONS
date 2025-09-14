VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rep_attn_daily 
   Caption         =   "ATTENDANCE REPORTS"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton opt_worker 
      Caption         =   "WORKER"
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
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton opt_staff 
      Caption         =   "STAFF"
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
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "SELECT EMPLOYEE TYPE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   4935
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   8775
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   39359
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   4320
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmb_mill 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   7455
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   360
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT MILL"
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
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "rep_attn_daily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_print_Click()
      If opt_staff.Value = True Then disname = "STAFF"
      If opt_worker.Value = True Then disname = "WORKERS"
      MousePointer = vbDefault
      millname = cmb_mill.Text
      gst_repconnect = "dsn=servall;uid=sa;pwd=serdat"
      cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
      cry_rep1.Formulas(1) = ("employee = '" & disname & "'")
      cry_rep1.Formulas(2) = ("startdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'")
      cry_rep1.Formulas(3) = ("enddate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'")
      cry_rep1.ReportFileName = "\\Appln\VBCRYREP\EDP REPORTS\day_attendance.rpt"
      If opt_staff.Value = True Then
         If cmb_mill.Text = "DANALAKSHMI PAPER MILLS PRIVATE LIMITED" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '1' and {attendance.emptype} = 'S'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PRIVATE LIMITED" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'S'"
         ElseIf cmb_mill.Text = "VIJAYALAKSHMI PAPER MILLS" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'S'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PVT LTD(POWER)" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'S'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI OIL EXTRACTION PLANT" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'S'"
         End If
      Else
         If cmb_mill.Text = "DANALAKSHMI PAPER MILLS PRIVATE LIMITED" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '1' and {attendance.emptype} = 'W'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PRIVATE LIMITED" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'W'"
         ElseIf cmb_mill.Text = "VIJAYALAKSHMI PAPER MILLS" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'W'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PVT LTD(POWER)" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'W'"
         ElseIf cmb_mill.Text = "SERVALAKSHMI OIL EXTRACTION PLANT" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'W'"
         End If
      
      End If
      cry_rep1.WindowState = crptMaximized
      cry_rep1.Connect = gst_repconnect
      cry_rep1.Action = 1
End Sub

Private Sub EXIT_Click()
     Unload Me
End Sub

Private Sub Form_Load()
 '' attndb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=servall;Data Source=servalldata"
  Fill_Combo_mill "Select company_name,company_code from mas_company order by company_code", cmb_mill
  opt_staff.Value = True
  st_date = Date
  end_date = Date

End Sub


