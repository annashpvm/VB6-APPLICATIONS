VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_applicatoin_reports 
   Caption         =   "Reports"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4920
      TabIndex        =   11
      Top             =   7200
      Width           =   1695
      Begin VB.CommandButton cmd_print 
         Caption         =   "&PRINT"
         Height          =   825
         Left            =   120
         Picture         =   "frm_applicatoin_reports.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   840
         Picture         =   "frm_applicatoin_reports.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   3720
      TabIndex        =   6
      Top             =   6000
      Width           =   4935
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   123011073
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   123011073
         CurrentDate     =   39359
      End
      Begin VB.Label Label9 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame11 
      Height          =   5175
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
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
            TabIndex        =   5
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
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4695
         Left            =   2160
         TabIndex        =   1
         Top             =   120
         Width           =   5055
         Begin VB.ListBox lst_dept 
            Enabled         =   0   'False
            Height          =   4110
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   240
            Width           =   4815
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
Attribute VB_Name = "frm_applicatoin_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PROCESS_Click()

End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_print_Click()
   Dim ds, dept As String
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
   ds = dept
   
   cry_rep1.Formulas(0) = "stdate='" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(1) = "enddate='" & Format(end_date.Value, "dd/mm/yyyy") & "'"

   

   
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\PAYROLL\Applicatoin_Report1.rpt"
   cry_rep1.ReplaceSelectionFormula (" {mas_applications.a_date}>=date(" & Format(st_date.Value, "yyyy,mm,dd") & ")" _
                                                  & " and {mas_applications.a_date}<=date(" & Format(end_date.Value, "yyyy,mm,dd") & ")" & ds)

   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1
       

End Sub

Private Sub Form_Load()
    st_date.Value = Now
    end_date.Value = Now
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

Private Sub opt_alldept_Click()
     lst_dept.Enabled = False
End Sub

Private Sub opt_selective_dept_Click()
     lst_dept.Enabled = True
End Sub
