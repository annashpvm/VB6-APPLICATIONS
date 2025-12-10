VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_attn_latecomers 
   Caption         =   "LATE COMERS REPORT"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   " EMPLOYEE IN "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   975
      Left            =   1920
      TabIndex        =   13
      Top             =   3240
      Width           =   6375
      Begin VB.TextBox etime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox stime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "TIME TO"
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
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "TIME FROM"
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
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   7455
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.OptionButton opt_staff 
      Caption         =   "STAFF"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton opt_worker 
      Caption         =   "WORKER"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   720
      TabIndex        =   5
      Top             =   4320
      Width           =   8775
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39359
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
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
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date"
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT MILL"
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
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "rep_attn_latecomers"
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
      cry_rep1.ReportFileName = "\\Appln\VBCRYREP\EDP REPORTS\latecomers.rpt"
      If opt_staff.Value = True Then
         If cmb_mill.Text = "DANALAKSHMI PAPER MILLS PRIVATE LIMITED" Then
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '1' and {attendance.emptype} = 'S' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PRIVATE LIMITED" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'S'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'S' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "VIJAYALAKSHMI PAPER MILLS" Then
'            cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'S'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'S' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PVT LTD(POWER)" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'S'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'S' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI OIL EXTRACTION PLANT" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'S'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'S' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         End If
      Else
         If cmb_mill.Text = "DANALAKSHMI PAPER MILLS PRIVATE LIMITED" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '1' and {attendance.emptype} = 'W'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '1' and {attendance.emptype} = 'W' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PRIVATE LIMITED" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'W'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '2' and {attendance.emptype} = 'W' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "VIJAYALAKSHMI PAPER MILLS" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'W'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '3' and {attendance.emptype} = 'W' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI PAPER AND BOARDS PVT LTD(POWER)" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'W'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '4' and {attendance.emptype} = 'W' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
         ElseIf cmb_mill.Text = "SERVALAKSHMI OIL EXTRACTION PLANT" Then
'             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'W'"
             cry_rep1.ReplaceSelectionFormula "{attendance.attndate} >= date(" & Format(st_date.Value, "yyyy,mm,dd") & ") and {attendance.attndate}  <= date(" & Format(end_date.Value, "yyyy,mm,dd") & ") and {attendance.empmill} = '5' and {attendance.emptype} = 'W' and {attendance.intime} >= '" & stime & "' and {attendance.intime} <= '" & etime & "' "
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



