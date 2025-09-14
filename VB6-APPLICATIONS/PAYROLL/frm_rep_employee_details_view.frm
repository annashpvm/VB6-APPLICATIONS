VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_employee_details_view 
   Caption         =   "EMPLOYEE DETAILS"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   16965
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   3015
      Left            =   10440
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
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
         Left            =   0
         TabIndex        =   32
         Top             =   3960
         Width           =   2505
         Begin VB.OptionButton opt_millselective 
            Caption         =   "SELECTIVE"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt_millall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton opt_retainer 
         Caption         =   "RETAINER"
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
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
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   7800
         Begin VB.OptionButton opt_all_location 
            Caption         =   "ALL"
            Height          =   405
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_vpt 
            Caption         =   "MILLS"
            Height          =   405
            Left            =   2400
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_cbe 
            Caption         =   "COIMBATORE"
            Height          =   405
            Left            =   4560
            TabIndex        =   28
            Top             =   240
            Width           =   1695
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1800
         Begin VB.OptionButton opt_permenent 
            Caption         =   "PERMENANT"
            Height          =   285
            Left            =   2400
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_all_per_trainee 
            Caption         =   "ALL"
            Height          =   285
            Left            =   360
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_trainee 
            Caption         =   "TRANIEES"
            Height          =   285
            Left            =   4680
            TabIndex        =   24
            Top             =   360
            Width           =   1335
         End
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
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1815
         Begin VB.OptionButton opt_dpm 
            Caption         =   "UNIT"
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt_slpb 
            Caption         =   "M-II"
            Height          =   375
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opt_vjpm 
            Caption         =   "VJPM"
            Height          =   375
            Left            =   3960
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opt_cogen 
            Caption         =   "II"
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opt_solvent 
            Caption         =   "I"
            Height          =   375
            Left            =   4080
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4680
      TabIndex        =   8
      Top             =   7200
      Width           =   1695
      Begin VB.CommandButton EXIT 
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   840
         Picture         =   "frm_rep_employee_details_view.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton PROCESS 
         Caption         =   "&PRINT"
         Height          =   825
         Left            =   120
         Picture         =   "frm_rep_employee_details_view.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE STATEMENT "
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
      Height          =   6795
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      Begin VB.Frame Frame6 
         Caption         =   "REPORT FOR"
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
         Height          =   1095
         Left            =   480
         TabIndex        =   11
         Top             =   3480
         Width           =   7695
         Begin MSComCtl2.DTPicker dt_to 
            Height          =   255
            Left            =   4560
            TabIndex        =   13
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   130220033
            CurrentDate     =   42187
         End
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   255
            Left            =   2040
            TabIndex        =   12
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   130220033
            CurrentDate     =   42187
         End
         Begin VB.Label Label2 
            Caption         =   "To"
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
            Left            =   3720
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "From"
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
            Left            =   960
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
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
         Left            =   480
         TabIndex        =   4
         Top             =   2160
         Width           =   7815
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton opt_sw 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            Height          =   375
            Left            =   4200
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "JOINED / RESIGNED"
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
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   7815
         Begin VB.OptionButton opt_emptype_joined 
            Caption         =   "NEWLY JOINED"
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptype_resigned 
            Caption         =   "RESIGNED"
            Height          =   375
            Left            =   4560
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   240
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_employee_details_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   dt_to.Value = Now
   dt_from.Value = Now - Day(dt_to.Value) + 1
  
End Sub

Private Sub opt_millall_Click()
    frame_mill.Enabled = False
End Sub

Private Sub opt_millselective_Click()
    frame_mill.Enabled = True
End Sub

Private Sub PROCESS_Click()
   Dim wp, qry1, qry2, qry3 As String
   MousePointer = vbDefault
   qry1 = ""
   qry2 = ""
   qry3 = ""
        
''   If opt_retainer.Value = True Then
''
''          wp = "SHVPM"
''          qry1 = ""
''
''       If opt_emptype_joined.Value = True Then
''          If qry1 <> "" Then
''             qry1 = qry1 + " and ({emp_voupay_mast.emp_status} = 'A' OR {emp_voupay_mast.emp_status} = 'B')"
''          Else
''             qry1 = " ({emp_voupay_mast.emp_status} = 'A' OR {emp_voupay_mast.emp_status} = 'B')"
''          End If
''       ElseIf opt_emptype_resigned.Value = True Then
''          If qry1 <> "" Then
''             qry1 = qry1 + " and {emp_voupay_mast.emp_status} = 'R'"
''          Else
''             qry1 = " {emp_voupay_mast.emp_status} = 'R'"
''          End If
''       End If
''
''       If qry1 <> "" Then
''              qry1 = qry1 + " and {emp_voupay_mast.emp_cat} = 'R'"
''       Else
''              qry1 = "{emp_mas.emp_cat} = 'R'"
''       End If
''
''
''       If opt_millselective.Value = True Then
''           If opt_dpm.Value = True Then
''                If qry1 <> "" Then
''                   qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 1"
''                Else
''                   qry1 = "{emp_voupay_mast.emp_company} = 1"
''                End If
''           End If
''           If opt_slpb.Value = True Then
''                If qry1 <> "" Then
''                   qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 2"
''                Else
''                   qry1 = "{emp_voupay_mast.emp_company} = 2"
''                End If
''           End If
''           If opt_vjpm.Value = True Then
''                If qry1 <> "" Then
''                   qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 3"
''                Else
''                   qry1 = "{emp_voupay_mast.emp_company} = 3"
''                End If
''           End If
''           If opt_cogen.Value = True Then
''                If qry1 <> "" Then
''                   qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 5"
''                Else
''                   qry1 = "{emp_voupay_mast.emp_company} = 5"
''                End If
''           End If
''           If opt_solvent.Value = True Then
''                If qry1 <> "" Then
''                   qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 8"
''                Else
''                   qry1 = "{emp_voupay_mast.emp_company} = 8"
''                End If
''           End If
''
''
''       End If
''
''       cry_rep1.Formulas(5) = "sdate = '" & Format(dt_from.Value, "dd/mm/yyyy") & "'"
''       cry_rep1.Formulas(6) = "edate = '" & Format(dt_to.Value, "dd/mm/yyyy") & "'"
''
''       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_vou_empdetails_joining_deptwise.rpt"
''       gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
''       cry_rep1.PrinterSelect
''       If opt_dpm.Value = True Then
''          cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PVT LTD'")
''       ElseIf opt_slpb.Value = True Then
''          cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARAPAPER AND BOARDS PVT LTD '")
''          cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PVT LTD (UNIT-II) '")
''       ElseIf opt_vjpm.Value = True Then
''          cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARAPAPER MILLS'")
''       ElseIf opt_cogen.Value = True Then
''          cry_rep1.Formulas(2) = ("millname= 'T'")
''       ElseIf opt_solvent.Value = True Then
''          cry_rep1.Formulas(2) = ("millname= 'OIL PLANT'")
''       End If
''
''       cry_rep1.Formulas(3) = ("sthead= 'RETAINER'")
''       If opt_emptype_joined = True Then
''          cry_rep1.Formulas(4) = ("empstatus= 'NEWLY JOINED EMPLOYEES'")
''          pst_qry = "{emp_voupay_mast.emp_doj} >= " & " date(" & Format$(dt_from, "yyyy,mm,dd") & ") AND " _
''                  & " {emp_voupay_mast.emp_doj} <= " & " date(" & Format$(dt_to, "yyyy,mm,dd") & ") "
''          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_vou_empdetails_joining_deptwise.rpt"
''
''       ElseIf opt_emptype_resigned.Value = True Then
''          cry_rep1.Formulas(4) = ("empstatus= 'RESIGNED EMPLOYEES'")
''          pst_qry = "{emp_voupay_mast.emp_resigneddate} >= " & " date(" & Format$(dt_from, "yyyy,mm,dd") & ") AND " _
''                  & " {emp_voupay_mast.emp_resigneddate} <= " & " date(" & Format$(dt_to, "yyyy,mm,dd") & ") "
''
''          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_vou_empdetails_resigned_deptwise.rpt"
''
''       End If
''       If qry1 = "" Then
''             cry_rep1.ReplaceSelectionFormula (pst_qry)
''       Else
''             cry_rep1.ReplaceSelectionFormula (pst_qry + " and " + qry1)
''       End If
''
''   Else
       
       wp = ""
       If opt_emptype_joined.Value = True Then
          If qry1 <> "" Then
             qry1 = qry1 + " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
          Else
             qry1 = " ({emp_mas.emp_status} = 'A')"
          End If
       ElseIf opt_emptype_resigned.Value = True Then
          If qry1 <> "" Then
             qry1 = qry1 + " and {emp_mas.emp_status} = 'R'"
          Else
             qry1 = " {emp_mas.emp_status} = 'R'"
          End If
       End If
       
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
            
            
                If qry1 <> "" Then
                   qry1 = qry1 + "and {emp_mas.emp_company} = 1"
                Else
                   qry1 = "{emp_mas.emp_company} = 1"
                End If
            
           
              
       cry_rep1.Formulas(5) = "sdate = '" & Format(dt_from.Value, "dd/mm/yyyy") & "'"
       cry_rep1.Formulas(6) = "edate = '" & Format(dt_to.Value, "dd/mm/yyyy") & "'"
            
       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_empdetails_deptwise.rpt"
       gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
       cry_rep1.PrinterSelect
    ''   cry_rep1.Formulas(5) = ("Ason='" & Format(dt_ason.Value, "yyyy/MM/dd") & "'")
       
     cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD'")

       
       If opt_staff.Value = True Then
          cry_rep1.Formulas(3) = ("sthead= 'STAFF'")
       ElseIf opt_worker.Value = True Then
          cry_rep1.Formulas(3) = ("sthead= 'WORKER'")
       ElseIf opt_sw.Value = True Then
          cry_rep1.Formulas(3) = ("sthead= 'STAFF / WORKER'")
       End If
       If opt_emptype_joined = True Then
          cry_rep1.Formulas(4) = ("empstatus= 'NEWLY JOINED EMPLOYEES'")
          pst_qry = "{emp_mas.emp_doj} >= " & " date(" & Format$(dt_from, "yyyy,mm,dd") & ") AND " _
                  & " {emp_mas.emp_doj} <= " & " date(" & Format$(dt_to, "yyyy,mm,dd") & ") "
          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_empdetails_joining_deptwise.rpt"
    
       ElseIf opt_emptype_resigned.Value = True Then
          cry_rep1.Formulas(4) = ("empstatus= 'RESIGNED EMPLOYEES'")
          pst_qry = "{emp_mas.emp_resigneddate} >= " & " date(" & Format$(dt_from, "yyyy,mm,dd") & ") AND " _
                  & " {emp_mas.emp_resigneddate} <= " & " date(" & Format$(dt_to, "yyyy,mm,dd") & ") "
          
          cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\rpt_empdetails_resigning_deptwise.rpt"
    
       End If
       If qry1 = "" Then
             cry_rep1.ReplaceSelectionFormula (pst_qry)
       Else
             cry_rep1.ReplaceSelectionFormula (pst_qry + " and " + qry1)
       End If
''   End If
   cry_rep1.DiscardSavedData = True
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
   
   
   Exit Sub
 End Sub



