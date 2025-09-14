VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_prodn_incentive 
   Caption         =   "OVER TIME REPORTS"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   17940
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Caption         =   "DEPARTMENT"
      Height          =   5175
      Left            =   8400
      TabIndex        =   20
      Top             =   1200
      Width           =   7335
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1935
         Begin VB.OptionButton opt_selective_dept 
            Caption         =   "Selective"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   25
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton opt_alldept 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   24
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   5055
         Begin VB.ListBox lst_dept 
            Enabled         =   0   'False
            Height          =   4110
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2160
      TabIndex        =   15
      Top             =   6960
      Width           =   1935
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   960
         Picture         =   "frm_rep_prodn_incentive.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   0
         Picture         =   "frm_rep_prodn_incentive.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Select Mill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   9
      Top             =   960
      Width           =   15
      Begin VB.OptionButton opt_cogen 
         Caption         =   "COGEN"
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
         Height          =   195
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_vjpm 
         Caption         =   "VJPM"
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
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "All"
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
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton opt_dpm3 
         Caption         =   "DPM-2"
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
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_dpm1 
         Caption         =   "DPM-1"
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
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   8055
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130351105
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130351105
         CurrentDate     =   39359
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
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   7815
      Begin VB.OptionButton opt_emp_ot_wagesNew 
         Caption         =   "Employeewise Over Time - Wages - New"
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
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   5295
      End
      Begin VB.Frame frame_ot 
         Height          =   1455
         Left            =   5520
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
         Begin VB.OptionButton opt_ot_nonpf 
            Caption         =   "NON PF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton opt_ot_pf 
            Caption         =   "PF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton opt_ot_all 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.OptionButton opt_emp_daywise_abstract 
         Caption         =   "Employee-Daywise  Over Time Abstract "
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
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   6615
      End
      Begin VB.OptionButton opt_emp_ot_wages 
         Caption         =   "Employeewise Over Time - Wages"
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
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   5055
      End
      Begin VB.OptionButton opt_emp_abstract 
         Caption         =   "Employeewise Over Time - Abstract"
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   5055
      End
      Begin VB.OptionButton opt_emp_datewise 
         Caption         =   "Employee-Datewise  Over Time Reports"
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
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   6615
      End
      Begin VB.OptionButton opt_datewise 
         Caption         =   "Datewise Over Time Report"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   4695
      End
   End
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_prodn_incentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_view_Click()
   Dim date1 As Date
   date1 = 1 & "/" & 1 & " /" & 1900
   Dim sw, ds, emp, mill As String
   

   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   mname = "SRI HARI VENKATESWAR PAPER MILLS PVT LTD"
   Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   Cry_rep1.Formulas(3) = ""
   Cry_rep1.PrinterSelect
   
''   emp = ""
''   If opt_selective_emp.Value = True Then
''        Dim pin_row, i As Integer
''        i = 0
''        If lst_emp.ListCount > 0 Then
''           For pin_row = 0 To lst_emp.ListCount - 1
''               If lst_emp.Selected(pin_row) = True Then
''                  If i = 0 Then
''                     emp = " and ( {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
''                     i = i + 1
''                  Else
''                     emp = emp + " or {bio_empmas.bioemp_name} = '" & lst_emp.List(pin_row) & "'"
''                  End If
''               End If
''           Next pin_row
''        End If
''   End If
''   If emp <> "" Then emp = emp + ")"
''

   
   dept = ""
   If opt_selective_dept.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                     dept = " and ( {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                     i = i + 1
                  Else
                     dept = dept + " or {bio_empmas.bioemp_dept}= '" & lst_dept.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If dept <> "" Then dept = dept + ")"
  
  
   ''ds = sw + emp + mill
   ds = mill + dept
   
   
   If opt_datewise.Value = True Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_datewise_OT.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf opt_emp_datewise.Value = True Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_emp_datewise_OT.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf opt_emp_abstract.Value = True Then
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_OT_abstract.rpt"
         Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf opt_emp_ot_wages.Value = True Then
         If opt_ot_pf.Value = True Then
             If ds <> "" Then
                ds = ds + " and ({emp_mas.EMP_PFELIGIBLE} ='Y')"
             Else
                ds = " and ({emp_mas.EMP_PFELIGIBLE} ='Y')"
             End If
              Cry_rep1.Formulas(3) = ("pf_or_nonpf= ' (PF MEMBERS)'")

         ElseIf opt_ot_nonpf.Value = True Then
             If ds <> "" Then
                 ds = ds + " and ({emp_mas.EMP_PFELIGIBLE} ='N')"
             Else
                 ds = " and ({emp_mas.EMP_PFELIGIBLE} ='N')"
             End If
              Cry_rep1.Formulas(3) = ("pf_or_nonpf= ' (NON PF MEMBERS)'")
         Else
              Cry_rep1.Formulas(3) = ""
         End If
         
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_OT_wages.rpt"
         Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
   ElseIf opt_emp_ot_wagesNew.Value = True Then
         If opt_ot_pf.Value = True Then
             If ds <> "" Then
                ds = ds + " and ({emp_mas.EMP_PFELIGIBLE} ='Y')"
             Else
                ds = " and ({emp_mas.EMP_PFELIGIBLE} ='Y')"
             End If
              Cry_rep1.Formulas(3) = ("pf_or_nonpf= ' (PF MEMBERS)'")

         ElseIf opt_ot_nonpf.Value = True Then
             If ds <> "" Then
                 ds = ds + " and ({emp_mas.EMP_PFELIGIBLE} ='N')"
             Else
                 ds = " and ({emp_mas.EMP_PFELIGIBLE} ='N')"
             End If
              Cry_rep1.Formulas(3) = ("pf_or_nonpf= ' (NON PF MEMBERS)'")
         Else
              Cry_rep1.Formulas(3) = ""
         End If
         
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_OT_wagesNew.rpt"
         Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
                  
   ElseIf opt_emp_daywise_abstract.Value = True Then
         Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_emp_datewise_OT_abstract.rpt"
         Cry_rep1.ReplaceSelectionFormula ("{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
         pst_qry = "{bio_worker_daily_pihrs.w_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_worker_daily_pihrs.w_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " "
   End If
   
   

   
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1

End Sub

Private Sub exit_Click()
    Unload Me
End Sub
    
Private Sub Form_Load()
    mill = ""
    mcode = ""
    mname = "SRI HARI VENKATESWARA PAPER MILLS PVT LTD'"
    st_date.Value = Now
    end_date.Value = Now

    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    sql = "select bioemp_dept from bio_empmas group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
''        lst_dept.ItemData(lst_dept.NewIndex) = payrs("dept_code")
        payrs.MoveNext
    Wend
    payrs.Close
    
End Sub

Private Sub opt_alldept_Click()
     lst_dept.Enabled = False
End Sub

Private Sub opt_emp_ot_wages_Click()
    frame_ot.Visible = True
End Sub

Private Sub opt_emp_ot_wagesNew_Click()
    frame_ot.Visible = True
End Sub

Private Sub opt_selective_dept_Click()
     lst_dept.Enabled = True
End Sub
