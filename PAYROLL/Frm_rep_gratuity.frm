VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_rep_gratuity 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   600
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "GRATUITY STATEMENT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   7275
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9240
      Begin MSComCtl2.DTPicker dt 
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   41872
      End
      Begin VB.Frame Frame7 
         Caption         =   "STAFF / WORKER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   600
         TabIndex        =   21
         Top             =   2400
         Width           =   7815
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton opt_sw 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            Height          =   375
            Left            =   4560
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   600
         TabIndex        =   17
         Top             =   3480
         Width           =   7800
         Begin VB.OptionButton opt_cbe 
            Caption         =   "COIMBATORE"
            Height          =   405
            Left            =   4560
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton opt_vpt 
            Caption         =   "MILLS"
            Height          =   405
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_all_location 
            Caption         =   "ALL"
            Height          =   405
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SELECT PERMENANT / TRANIEE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   825
         Left            =   600
         TabIndex        =   13
         Top             =   4440
         Width           =   7800
         Begin VB.OptionButton opt_permenent 
            Caption         =   "PERMENANT"
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_all_per_trainee 
            Caption         =   "ALL"
            Height          =   285
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_trainee 
            Caption         =   "TRANIEES"
            Height          =   285
            Left            =   4680
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3840
         TabIndex        =   10
         Top             =   6120
         Width           =   1695
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "Frm_rep_gratuity.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "Frm_rep_gratuity.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "MILLWISE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1185
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   7785
         Begin VB.OptionButton opt_millall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt_millselective 
            Caption         =   "SELECTIVE"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame frame_mill 
            Caption         =   "SELECT MILL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   1920
            TabIndex        =   2
            Top             =   240
            Width           =   5535
            Begin VB.OptionButton opt_dpm 
               Caption         =   "DPM"
               Height          =   375
               Left            =   240
               TabIndex        =   7
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opt_slpb 
               Caption         =   "SLPB"
               Height          =   375
               Left            =   1200
               TabIndex        =   6
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt_vjpm 
               Caption         =   "VJPM"
               Height          =   375
               Left            =   2040
               TabIndex        =   5
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_cogen 
               Caption         =   "COGEN"
               Height          =   375
               Left            =   3000
               TabIndex        =   4
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_solvent 
               Caption         =   "SOLVENT"
               Height          =   375
               Left            =   4080
               TabIndex        =   3
               Top             =   360
               Width           =   1215
            End
         End
      End
      Begin VB.Label Label1 
         Caption         =   "As On "
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
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_rep_gratuity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 dt.Value = Now
End Sub

Private Sub opt_millall_Click()
    frame_mill.Enabled = False
End Sub

Private Sub opt_millselective_Click()
    frame_mill.Enabled = True
End Sub

Private Sub PROCESS_Click()
   Dim wp, qry1, qry2, qry3 As String
   Dim date1 As String
   MousePointer = vbDefault
   qry1 = ""
   qry2 = ""
   qry3 = ""
   date1 = Format(DateAdd("m", -54, dt.Value), "yyyy,mm,dd")
   date1 = Format(dt.Value, "yyyy,mm,dd")
         wp = ""
      qry1 = ""
''   If opt_emptypeall.Value = True Then
''   ElseIf opt_emptype_active.Value = True Then
''      If qry1 <> "" Then
''         qry1 = qry1 + " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
''      Else
''         qry1 = " ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
''      End If
''   ElseIf opt_emptype_resigned.Value = True Then
''      If qry1 <> "" Then
''         qry1 = qry1 + " and {emp_mas.emp_status} = 'R'"
''      Else
''         qry1 = " {emp_mas.emp_status} = 'R'"
''      End If
''   End If
   
   If opt_sw.Value = True Then
   ElseIf opt_staff.Value = True Then
      If qry1 <> "" Then
          qry1 = qry1 + " and {emp_mas.emp_cat} = 'S'"
      Else
          qry1 = "{emp_mas.emp_cat} = 'S'"
      End If
   ElseIf opt_worker.Value = True Then
      If qry1 <> "" Then
         qry1 = qry1 + "and {emp_mas.emp_cat} = 'W'"
      Else
         qry1 = "{emp_mas.emp_cat} = 'W'"
      End If
   End If
        
   If opt_all_per_trainee.Value = True Then
   ElseIf opt_permenent.Value = True Then
       If qry1 <> "" Then
          qry1 = qry1 + "and ({emp_mas.emp_type} = 0  or {emp_mas.emp_type} = 2) "
       Else
          qry1 = "({emp_mas.emp_type} = 0  or {emp_mas.emp_type} = 2) "
       End If
   ElseIf opt_trainee.Value = True Then
       If qry1 <> "" Then
          qry1 = qry1 + "and ({emp_mas.emp_type} = 1  or {emp_mas.emp_type} = 3) "
       Else
          qry1 = "({emp_mas.emp_type} = 1  or {emp_mas.emp_type} = 3) "
       End If
   End If
   
   If opt_millselective.Value = True Then
       If opt_dpm.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {emp_mas.emp_company} = 1"
            Else
               qry1 = "{emp_mas.emp_company} = 1"
            End If
       End If
       If opt_slpb.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {emp_mas.emp_company} = 2"
            Else
               qry1 = "{emp_mas.emp_company} = 2"
            End If
       End If
       If opt_vjpm.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {emp_mas.emp_company} = 3"
            Else
               qry1 = "{emp_mas.emp_company} = 3"
            End If
       End If
       If opt_cogen.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {emp_mas.emp_company} = 5"
            Else
               qry1 = "{emp_mas.emp_company} = 5"
            End If
       End If
       If opt_solvent.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and {emp_mas.emp_company} = 8"
            Else
               qry1 = "{emp_mas.emp_company} = 8"
            End If
       End If
       
   End If
''   If opt_millall.Value = True And opt_all_location.Value = True And opt_all_per_trainee.Value = True And opt_sw.Value = True Then
''       'qry1 = "{emp_salary.s_month} = " & dt.Month & "  and  {emp_salary.s_finyear}= " & finyear & " and  {emp_salary.s_year} = " & dt.year & "   and {emp_mas.emp_doj}< Date(" & date1 & ") "
   ''Else
''        qry1 = qry1 + " and {emp_salary.s_month} = " & dt.Month & "  and  {emp_salary.s_finyear}= " & finyear & " and  {emp_salary.s_year} = " & dt.year & "  and {emp_mas.emp_doj}< Date(" & date1 & ")"
   ''End If
   
   If opt_millall.Value = True And opt_all_location.Value = True And opt_all_per_trainee.Value = True And opt_sw.Value = True Then
       qry1 = "{emp_salary.s_month} = " & dt.Month & "  and  {emp_salary.s_finyear}= " & finyear & " and  {emp_salary.s_year} = " & dt.year & "  "
   Else
        qry1 = qry1 + " and {emp_salary.s_month} = " & dt.Month & "  and  {emp_salary.s_finyear}= " & finyear & " and  {emp_salary.s_year} = " & dt.year & "  "
   End If
   
   
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Gratuity_deptwise.rpt"
   
   
   ''cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Gratuity_statement.rpt"
   
   
   
   
   
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   
   
   cry_rep1.Formulas(5) = ("Ason='" & Format(dt.Value, "yyyy/MM/dd") & "'")
   If opt_dpm.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD'")
   ElseIf opt_slpb.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD''")
   ElseIf opt_vjpm.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD''")
   ElseIf opt_cogen.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'T'")
   ElseIf opt_solvent.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'OIL PLANT'")
   End If
   
   If opt_staff.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF'")
   ElseIf opt_worker.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'WORKER'")
   ElseIf opt_sw.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF / WORKER'")
   End If
   
''   If opt_emptypeall.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT + RESIGNED EMPLOYEES'")
''   ElseIf opt_emptype_active.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT EMPLOYEES'")
''   ElseIf opt_emptype_resigned.Value = True Then
''      cry_rep1.Formulas(4) = ("empstatus= 'RESIGNED EMPLOYEES'")
''
''   End If
   
   cry_rep1.DiscardSavedData = True
   cry_rep1.ReplaceSelectionFormula (qry1)
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
   Exit Sub
 End Sub



