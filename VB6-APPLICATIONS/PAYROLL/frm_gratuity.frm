VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_gratuity_report 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Gratuity Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00FFC0FF&
      Height          =   4695
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      Begin Crystal.CrystalReport Cry_rep1 
         Left            =   1320
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4800
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00FFC0FF&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dt_left 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60227585
         CurrentDate     =   41537
      End
      Begin VB.ComboBox cmb_employeename 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Date of Left from service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1080
         TabIndex        =   3
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_gratuity_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_Refresh_Click()

End Sub

Private Sub cmd_view_Click()
If Trim(cmb_employeename.Text) = "" Then
      MsgBox ("Select the Employee Name")
      Exit Sub
End If

      gst_repconnect = "dsn=pay_dsn;uid=sa;pwd=serdat"
      Cry_rep1.PrinterSelect
      Cry_rep1.ReportFileName = "\\192.168.11.6\vbcryrep\payroll\Gratuity_settlement_statement.rpt"
''      Cry_rep1.Formulas(1) = ("emp_name = '" & cmb_employeename & "'")
''      Cry_rep1.Formulas(2) = ("date = '" & dt_left & "'")
''      cry_rep1.ReplaceSelectionFormula ("emp_mas.emp_name} = " & cmb_employeename & " and {emp_salary.s_year} = " & Year(dt_left) & " and {emp_salary.s_month}= " & Month(dt_left) & _
                                              "and {emp_salary.s_company} = '" & company_code & "'")
''      cry_rep1.ReplaceSelectionFormula ("emp_mas.emp_name} = " & cmb_employeename & " and {emp_salary.s_company} = " & company_code & "")

'       Cry_rep1.ReplaceSelectionFormula ("{emp_mas.emp_name} = " & cmb_employeename.Text & " and {emp_salary.s_month} = " & Month(dt_left) & _
                                               "and {emp_salary.s_company} = '" & company_code & "' and {emp_salary.s_emptype} = " & 0)
        
                                              
      Cry_rep1.WindowState = crptMaximized
      Cry_rep1.Connect = gst_repconnect
      Cry_rep1.Action = 1
    
End Sub

Private Sub Form_Load()

Set paydb = New ADODB.Connection
Set payrs = New ADODB.Recordset

sql = "select * from emp_mas  where emp_company = '" & company_code & "' order by emp_name"
paydb.Open pay
payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
If Not payrs.EOF Then
    payrs.MoveFirst
    cmb_employeename.Clear
    While Not payrs.EOF
        cmb_employeename.AddItem payrs("emp_name")
        payrs.MoveNext
    Wend
End If
End Sub
