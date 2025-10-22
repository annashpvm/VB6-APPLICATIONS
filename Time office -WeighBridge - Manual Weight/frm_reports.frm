VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_reports 
   Caption         =   "WEIGHMENT REPORT"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dt_startdate 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   122290177
         CurrentDate     =   45262
      End
      Begin MSComCtl2.DTPicker dt_enddate 
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   122290177
         CurrentDate     =   45262
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   480
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_print_Click()
            
''t_wb_date = '" & Format(dt_ticket, "yyyy/MM/dd") & "'
            
            qry = "{trn_weighbridge_entry.t_wb_date}  >=  date(" & Format$(dt_startdate, "yyyy,mm,dd") & ") and {trn_weighbridge_entry.t_wb_date}  <= date(" & Format$(dt_enddate, "yyyy,mm,dd") & ")"
            qry = "{trn_weighbridge_entry.t_wb_2nd_time}  >=  date(" & Format$(dt_startdate, "yyyy,mm,dd") & ") and {trn_weighbridge_entry.t_wb_2nd_time}  <= date(" & Format$(dt_enddate, "yyyy,mm,dd") & ")"
''            qry = "{trn_weighbridge_entry.t_wb_date}  >=  date(" & Format$(dt_startdate, "yyyy,mm,dd") & ") and {trn_weighbridge_entry.t_wb_date}  <= date(" & Format$(dt_enddate, "yyyy,mm,dd") & ")"
            
            MousePointer = vbDefault
            
            Cry_rep1.Formulas(0) = "sdate = '" & Format(dt_startdate.Value, "dd/mm/yyyy") & "'"
            Cry_rep1.Formulas(1) = "edate = '" & Format(dt_enddate.Value, "dd/mm/yyyy") & "'"
            
            gst_repconnect = "dsn=mysql;uid=root;pwd=P@ssw0rD;database=shvpm"
            
''            Cry_rep1.PrinterSelect
''            Cry_rep1.ReportFileName = "d:\wbreports\weighment_details.rpt"
            Cry_rep1.ReportFileName = "\\10.0.0.243\wbreports\weighment_details.rpt"
            Cry_rep1.ReplaceSelectionFormula (qry)
            Cry_rep1.WindowState = crptMaximized
            Cry_rep1.Connect = gst_repconnect
            MousePointer = vbNormal
            Cry_rep1.Action = 1
            
End Sub

Private Sub Form_Load()
    dt_startdate.Value = Now
    dt_enddate.Value = Now
    Call gen_dbconnection
    adocmd_mysql.ActiveConnection = gen_connection_mysql


End Sub
