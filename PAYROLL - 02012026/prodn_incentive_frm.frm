VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form prodn_incentive_frm 
   Caption         =   "PRODUCTION INCENTIVE REPORT"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   1290
      Top             =   7530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Report period"
      Height          =   5955
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   9360
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   4320
         TabIndex        =   9
         Top             =   4320
         Width           =   2175
         Begin VB.CommandButton exit 
            Caption         =   "&Exit"
            Height          =   825
            Left            =   1080
            Picture         =   "prodn_incentive_frm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   960
         End
         Begin VB.CommandButton print 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "prodn_incentive_frm.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   960
         End
      End
      Begin MSComCtl2.DTPicker SDATE 
         Height          =   510
         Left            =   2910
         TabIndex        =   5
         Top             =   2970
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   900
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117768193
         CurrentDate     =   37562
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select "
         Height          =   1080
         Left            =   1140
         TabIndex        =   1
         Top             =   675
         Width           =   7320
         Begin VB.OptionButton opt_trainee 
            Caption         =   "TRAINEE"
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
            Height          =   405
            Left            =   5370
            TabIndex        =   4
            Top             =   435
            Width           =   1605
         End
         Begin VB.OptionButton opt_permenent 
            Caption         =   "PERMANENT"
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
            Height          =   405
            Left            =   2400
            TabIndex        =   3
            Top             =   390
            Width           =   1830
         End
         Begin VB.OptionButton opt_all 
            Caption         =   "ALL"
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
            Height          =   405
            Left            =   405
            TabIndex        =   2
            Top             =   375
            Width           =   1035
         End
      End
      Begin MSComCtl2.DTPicker EDATE 
         Height          =   510
         Left            =   6795
         TabIndex        =   6
         Top             =   2955
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117768193
         CurrentDate     =   37562
      End
      Begin VB.Label Label2 
         Caption         =   "END DATE"
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
         Height          =   270
         Left            =   5220
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "START   DATE"
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
         Height          =   510
         Left            =   1140
         TabIndex        =   7
         Top             =   3150
         Width           =   1815
      End
   End
End
Attribute VB_Name = "prodn_incentive_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    SDATE = DATE
    EDATE = DATE
    ''pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
End Sub


Private Sub print_Click()
   If opt_all.Value = True Then disname = "ALL WORKERS INCENTIVE WAGES FROM "
   If opt_permenent.Value = True Then disname = "PERMENENT WORKER INCENTIVE WAGES FROM "
   If opt_trainee.Value = True Then disname = "TEMPERORY WORKER INCENTIVE WAGES FROM "
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.Formulas(0) = ("millname= '" & millname & "'")
   cry_rep1.Formulas(1) = ("sthead = '" & disname & "'")
   cry_rep1.Formulas(2) = ("sdate = '" & Format(SDATE.Value, "dd/mm/yyyy") & "'")
   cry_rep1.Formulas(3) = ("edate = '" & Format(EDATE.Value, "dd/mm/yyyy") & "'")
   cry_rep1.ReportFileName = "\\annadurai\d\payroll\prod_incentive_rep.rpt"
   cry_rep1.ReplaceSelectionFormula "{attn_entry.attn_date}  >= date(" & Format(SDATE.Value, "yyyy,mm,dd") & ") and {attn_entry.attn_date}  <= date(" & Format(EDATE.Value, "yyyy,mm,dd") & ") and {attn_entry.attn_emptype} =  2"
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
End Sub


