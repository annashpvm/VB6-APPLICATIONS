VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rep_cost_report 
   Caption         =   "COST REPORT"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   18825
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3840
      TabIndex        =   7
      Top             =   7800
      Width           =   1695
      Begin VB.CommandButton cmd_print 
         Caption         =   "&PRINT"
         Height          =   825
         Left            =   120
         Picture         =   "frm_rep_cost_report.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton EXIT 
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   840
         Picture         =   "frm_rep_cost_report.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Frame Frame5 
         Height          =   4455
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   7335
         Begin VB.Frame Frame11 
            Height          =   615
            Left            =   3600
            TabIndex        =   16
            Top             =   240
            Width           =   3495
            Begin VB.OptionButton opt_de_select_all 
               Caption         =   "De-Select All"
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
               Left            =   1440
               TabIndex        =   18
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton opt_select_all 
               Caption         =   "Select All"
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
               TabIndex        =   17
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame7 
            Height          =   3375
            Left            =   360
            TabIndex        =   14
            Top             =   960
            Width           =   6615
            Begin VB.ListBox lst_dept 
               Enabled         =   0   'False
               Height          =   2985
               Left            =   480
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   15
               Top             =   240
               Width           =   5895
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   3135
            Begin VB.OptionButton opt_selective_dept 
               Caption         =   "Selective"
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
               Left            =   1560
               TabIndex        =   13
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton opt_alldept 
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
               TabIndex        =   12
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1335
         Left            =   840
         TabIndex        =   4
         Top             =   6480
         Width           =   5775
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   495
            Left            =   2640
            TabIndex        =   5
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61407233
            CurrentDate     =   42187
         End
         Begin MSComCtl2.DTPicker dt_to 
            Height          =   495
            Left            =   2640
            TabIndex        =   19
            Top             =   720
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61407233
            CurrentDate     =   42187
         End
         Begin VB.Label Label2 
            Caption         =   "TO"
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
            Left            =   1200
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "AS ON"
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
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   1665
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   5640
         Begin VB.OptionButton opt_deptwise_abstract 
            Caption         =   "DEPARTMENTWISE ABSTRACT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   4575
         End
         Begin VB.OptionButton opt_empwise_details 
            Caption         =   "EMPLOYEEWISE DETAILS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Top             =   840
            Width           =   4335
         End
      End
   End
   Begin Crystal.CrystalReport cry_rep1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_rep_cost_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_print_Click()
   Dim dept As String
   dept = ""
   If opt_selective_dept.Value = True Then
        i = 0
        If lst_dept.ListCount > 0 Then
           For pin_row = 0 To lst_dept.ListCount - 1
               If lst_dept.Selected(pin_row) = True Then
                  If i = 0 Then
                    dept = " and ( {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                     i = i + 1
                  Else
                    dept = dept + " or {bio_empmas.bioemp_dept} = '" & lst_dept.List(pin_row) & "'"
                  End If
               End If
           Next pin_row
        End If
   End If
   
   
   
   If dept <> "" Then dept = dept + ")"
   
   
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.Formulas(0) = "dtfrom = '" & Format(dt_from.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(1) = "dtto = '" & Format(dt_to.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(2) = ("millname= 'SRI HARI VENKATESWARA PAPER MILLS PRIVATE LIMITD'")
   cry_rep1.Formulas(3) = ""
   
   
   cry_rep1.PrinterSelect
   
   If opt_deptwise_abstract.Value = True Then
       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_costing_deptwise_abstract.rpt"
   Else
       cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_daily_costing_deptwise.rpt"
   End If
   cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.ds_date} =  date(" & Format$(dt_from, "yyyy,mm,dd") & ")  and {bio_device_shiftlogs.ds_shift_in} <> Date (1900,01,01)" & dept)

   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1


End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dt_from.Value = Now
    dt_to.Value = Now
    
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    sql = "select bioemp_dept  from bio_empmas group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    lst_dept.Visible = True
End Sub

Private Sub PROCESS_Click()

End Sub

Private Sub opt_alldept_Click()
     For i = 0 To lst_dept.ListCount() - 1
        lst_dept.Selected(i) = True
    Next
End Sub

Private Sub opt_de_select_all_Click()
     For i = 1 To lst_dept.ListCount() - 1
        lst_dept.Selected(i) = False
    Next
End Sub

Private Sub opt_select_all_Click()
     For i = 0 To lst_dept.ListCount() - 1
        lst_dept.Selected(i) = True
    Next
    
End Sub

Private Sub opt_selective_dept_Click()
    lst_dept.Enabled = True
    
End Sub
