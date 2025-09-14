VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_employee_details 
   Caption         =   "EMPLOYEE DETAILS"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   13245
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   5175
      Left            =   11160
      TabIndex        =   36
      Top             =   720
      Width           =   7335
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   120
         TabIndex        =   39
         Top             =   120
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   2160
         TabIndex        =   37
         Top             =   120
         Width           =   5055
         Begin VB.ListBox lst_dept 
            Enabled         =   0   'False
            Height          =   4110
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   38
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   1575
      Left            =   15840
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Frame Frame6 
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
         Left            =   720
         TabIndex        =   29
         Top             =   5040
         Visible         =   0   'False
         Width           =   7800
         Begin VB.OptionButton opt_ctc 
            Caption         =   "CTC"
            Height          =   285
            Left            =   360
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_master 
            Caption         =   "MASTER"
            Height          =   285
            Left            =   2400
            TabIndex        =   30
            Top             =   360
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
         TabIndex        =   25
         Top             =   3960
         Width           =   6240
         Begin VB.OptionButton opt_permenent 
            Caption         =   "PERMENANT"
            Height          =   285
            Left            =   2400
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton opt_all_per_trainee 
            Caption         =   "ALL"
            Height          =   285
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_trainee 
            Caption         =   "TRANIEES"
            Height          =   285
            Left            =   4680
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton opt_retainer 
         Caption         =   "RETAINER"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton opt_retainer_address 
         Caption         =   "RETAINER WITH ADDRESS"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   1920
         Width           =   2535
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
         Height          =   225
         Left            =   3240
         TabIndex        =   14
         Top             =   1200
         Width           =   105
         Begin VB.OptionButton opt_millall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt_millselective 
            Caption         =   "SELECTIVE"
            Height          =   375
            Left            =   240
            TabIndex        =   21
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
            TabIndex        =   15
            Top             =   240
            Width           =   5535
            Begin VB.OptionButton opt_shvpm 
               Caption         =   "I"
               Height          =   375
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opt_slpb 
               Caption         =   "-II"
               Height          =   375
               Left            =   1200
               TabIndex        =   19
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt_vjpm 
               Caption         =   "III"
               Height          =   375
               Left            =   2040
               TabIndex        =   18
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_cogen 
               Caption         =   "C"
               Height          =   375
               Left            =   3000
               TabIndex        =   17
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_solvent 
               Caption         =   "S"
               Height          =   375
               Left            =   4080
               TabIndex        =   16
               Top             =   360
               Width           =   1215
            End
         End
      End
   End
   Begin MSComCtl2.DTPicker dt_ason 
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   130809857
      CurrentDate     =   42187
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
      Height          =   7755
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   9240
      Begin VB.Frame Frame4 
         Caption         =   "DETAILS"
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
         Height          =   2505
         Left            =   720
         TabIndex        =   32
         Top             =   3000
         Width           =   6600
         Begin VB.OptionButton opt_salary_details_ctc 
            Caption         =   "SALARY DETAILS  - CTC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   42
            Top             =   1800
            Width           =   3975
         End
         Begin VB.OptionButton opt_salary_details 
            Caption         =   "SALARY DETAILS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   35
            Top             =   1200
            Width           =   3975
         End
         Begin VB.OptionButton opt_all_details 
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opt_locationwise 
            Caption         =   "LOCATION WISE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   600
            TabIndex        =   33
            Top             =   720
            Width           =   3975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ACTIVE / RESIGNED"
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
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   6855
         Begin VB.OptionButton opt_emptype_resigned 
            Caption         =   "RESIGNED"
            Height          =   375
            Left            =   4560
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptype_active 
            Caption         =   "ACTIVE"
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt_emptypeall 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3000
         TabIndex        =   4
         Top             =   6240
         Width           =   1695
         Begin VB.CommandButton EXIT 
            Caption         =   "E&XIT"
            Height          =   825
            Left            =   840
            Picture         =   "frm_rep_employee_details.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   720
         End
         Begin VB.CommandButton PROCESS 
            Caption         =   "&PRINT"
            Height          =   825
            Left            =   120
            Picture         =   "frm_rep_employee_details.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   720
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
         Left            =   720
         TabIndex        =   1
         Top             =   1920
         Width           =   6855
         Begin VB.OptionButton opt_worker 
            Caption         =   "WORKER"
            Height          =   375
            Left            =   4440
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opt_sw 
            Caption         =   "ALL"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "STAFF"
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   480
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_employee_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   dt_ason.Value = Now
   opt_ctc.Value = True
   
       sql = "select * from pdept_mas   order by dept_name"
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    lst_dept.Clear
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("dept_name")
        payrs.MoveNext
    Wend
End Sub

Private Sub opt_millall_Click()
    frame_mill.Enabled = False
End Sub

Private Sub opt_millselective_Click()
    frame_mill.Enabled = True
End Sub

Private Sub opt_selective_dept_Click()
     lst_dept.Enabled = True
End Sub

Private Sub PROCESS_Click()
   Dim wp, qry1, qry2, qry3, dept As String
   MousePointer = vbDefault
   qry1 = ""
   qry2 = ""
   qry3 = ""
           wp = ""
           qry1 = ""
   
   
   
   dept = ""
   If opt_selective_dept.Value = True Then
        Dim pin_row, i As Integer
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                     dept = " and ( {pdept_mas.dept_name} = '" & lst_dept.List(pin_row) & "'"
                     i = i + 1
                  Else
                     dept = dept + " or {pdept_mas.dept_name}= '" & lst_dept.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   If dept <> "" Then dept = dept + ")"
   
   
      
      
   If opt_retainer.Value = True Or opt_retainer_address.Value = True Then
     
        
        If opt_emptypeall.Value = True Then
        ElseIf opt_emptype_active.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + " and ({emp_voupay_mast.emp_status} = 'A' OR {emp_voupay_mast.emp_status} = 'B')"
           Else
              qry1 = " ({emp_voupay_mast.emp_status} = 'A' OR {emp_voupay_mast.emp_status} = 'B')"
           End If
        ElseIf opt_emptype_resigned.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + " and {emp_voupay_mast.emp_status} = 'R'"
           Else
              qry1 = " {emp_voupay_mast.emp_status} = 'R'"
           End If
        End If
        
        If opt_sw.Value = True Then
        ElseIf opt_staff.Value = True Then
           If qry1 <> "" Then
               qry1 = qry1 + " and {emp_voupay_mast.emp_cat} = 'S'"
           Else
               qry1 = "{emp_voupay_mast.emp_cat} = 'S'"
           End If
        ElseIf opt_worker.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + "and {emp_voupay_mast.emp_cat} = 'W'"
           Else
              qry1 = "{emp_voupay_mast.emp_cat} = 'W'"
           End If
        ElseIf opt_retainer.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + "and {emp_voupay_mast.emp_cat} = 'R'"
           Else
              qry1 = "{emp_voupay_mast.emp_cat} = 'R'"
           End If
        
        End If
             
        If opt_all_per_trainee.Value = True Then
        ElseIf opt_permenent.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and ({emp_voupay_mast.emp_type} = 0  or {emp_voupay_mast.emp_type} = 2) "
            Else
               qry1 = "({emp_voupay_mast.emp_type} = 0  or {emp_voupay_mast.emp_type} = 2) "
            End If
        ElseIf opt_trainee.Value = True Then
            If qry1 <> "" Then
               qry1 = qry1 + "and ({emp_voupay_mast.emp_type} = 1  or {emp_voupay_mast.emp_type} = 3) "
            Else
               qry1 = "({emp_voupay_mast.emp_type} = 1  or {emp_voupay_mast.emp_type} = 3) "
            End If
        End If
   Else
              wp = ""
           qry1 = ""

        
        If opt_emptypeall.Value = True Then
        ElseIf opt_emptype_active.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + " and ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
           Else
              qry1 = " ({emp_mas.emp_status} = 'A' OR {emp_mas.emp_status} = 'B')"
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
        ElseIf opt_retainer.Value = True Then
           If qry1 <> "" Then
              qry1 = qry1 + "and {emp_voupay_mast.emp_cat} = 'R'"
           Else
              qry1 = "{emp_voupay_mast.emp_cat} = 'R'"
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
        
   
   End If
   
   
''---------------------
   If opt_retainer.Value = True Or opt_retainer_address.Value = True Then
   
        If opt_millselective.Value = True Then
            If opt_shvpm.Value = True Then
                  
                 If qry1 <> "" Then
                    qry1 = qry1 + "and {emp_voupay_mast.emp_company} = 1"
                 Else
                    qry1 = "{emp_voupay_mast.emp_company} = 1"
                 End If
            End If
            
            
        Else
                 If qry1 <> "" Then
                    qry1 = qry1 + "and ( {emp_voupay_mast.emp_company} = 1 or {emp_voupay_mast.emp_company} = 2 or {emp_voupay_mast.emp_company} = 3  or {emp_voupay_mast.emp_company} = 5)  "
                 Else
                    qry1 = "( {emp_voupay_mast.emp_company} = 1 or {emp_voupay_mast.emp_company} = 2 or {emp_mas.emp_company} = 3  or {emp_voupay_mast.emp_company} = 5) "
                 End If
        End If
        If opt_retainer.Value = True Then
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\massalary_retainer_deptwise.rpt"
        Else
           cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\massalary_retainer_address.rpt"
        End If
        
   Else
        If opt_millselective.Value = True Then
            If opt_shvpm.Value = True Then
                  
                 If qry1 <> "" Then
                    qry1 = qry1 + "and {emp_mas.emp_company} = 1"
                 Else
                    qry1 = "{emp_mas.emp_company} = 1"
                 End If
            End If
            
        Else
                 If qry1 <> "" Then
                    qry1 = qry1 + "and ( {emp_mas.emp_company} = 1 )  "
                 Else
                    qry1 = "( {emp_mas.emp_company} = 1 ) "
                 End If
        End If
'''        If opt_locationwise.Value = True Then
'''            If opt_ctc.Value = True Then
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\massalary_location_deptwise.rpt"
'''            Else
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\master_data_salary.rpt"
'''            End If
'''        Else
'''            If opt_ctc.Value = True Then
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\massalary_deptwise.rpt"
'''            Else
'''               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\master_data_salary.rpt"
'''            End If

        If opt_locationwise.Value = True Then
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\master_data_salary.rpt"
        ElseIf opt_locationwise.Value = True Then
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\master_data_salary.rpt"
        ElseIf opt_salary_details.Value = True Then
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\empdetails_deptwise.rpt"
        Else
               cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\empdetails_deptwise_CTC.rpt"
        End If
   
   End If
   
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.PrinterSelect
   cry_rep1.Formulas(5) = ("Ason='" & Format(dt_ason.Value, "yyyy/MM/dd") & "'")
   If opt_shvpm.Value = True Then
      cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LTD'")

   End If
   
   If opt_staff.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF'")
   ElseIf opt_worker.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'WORKER'")
   ElseIf opt_retainer.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'RETAINER'")
      
   ElseIf opt_sw.Value = True Then
      cry_rep1.Formulas(3) = ("sthead= 'STAFF / WORKER'")
   End If
   
   If opt_emptypeall.Value = True Then
      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT + RESIGNED EMPLOYEES'")
   ElseIf opt_emptype_active.Value = True Then
      cry_rep1.Formulas(4) = ("empstatus= 'CURRENT EMPLOYEES'")
   ElseIf opt_emptype_resigned.Value = True Then
      cry_rep1.Formulas(4) = ("empstatus= 'RESIGNED EMPLOYEES'")

   End If
   
   cry_rep1.DiscardSavedData = True
   cry_rep1.ReplaceSelectionFormula (qry1 & dept)
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
   Exit Sub
 End Sub


