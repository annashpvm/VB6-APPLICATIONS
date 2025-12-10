VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_rep_canteen_details 
   Caption         =   "CANTEEN DETAILS"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   17895
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   7335
      Begin VB.OptionButton opt_empwise_recovery 
         Caption         =   "Canteen Recovery - Employeewise"
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
         TabIndex        =   11
         Top             =   1560
         Width           =   6615
      End
      Begin VB.OptionButton opt_canteen_expenses 
         Caption         =   "Canteen Expenses"
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
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   4695
      End
      Begin VB.OptionButton opt_date_empwise_recovery 
         Caption         =   "Canteen Recovery - DATE- Employeewise"
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
         TabIndex        =   9
         Top             =   840
         Width           =   6615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   5040
      Width           =   7335
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121176065
         CurrentDate     =   44594
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121176065
         CurrentDate     =   44594
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
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   0
         Picture         =   "frm_rep_canteen_details.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   960
         Picture         =   "frm_rep_canteen_details.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   960
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
Attribute VB_Name = "frm_rep_canteen_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_view_Click()
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   mname = "SRI HARI VENKATESWAR PAPER MILLS PVT LTD"
    
   Cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   Cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   Cry_rep1.Formulas(3) = ""
   Dim canteen As String
   If Year(st_date.Value) <= 2022 And Month(st_date.Value) <= 5 And Month(st_date.Value) <= 15 Then
       canteen = "GANESHAMOORTHY CANTEEN EXPENSES BILL FOR THE THE PERIOD FROM"
   ElseIf Year(st_date.Value) <= 2022 And Month(st_date.Value) <= 6 And Month(st_date.Value) <= 30 Then
      canteen = "ALAGARSAMY CANTEEN EXPENSES BILL FOR THE THE PERIOD FROM"
   ElseIf Year(st_date.Value) <= 2023 And Month(st_date.Value) <= 1 And Month(st_date.Value) <= 31 Then
      canteen = "RAJESH KUMAR CANTEEN EXPENSES BILL FOR THE THE PERIOD FROM"
   Else
      canteen = "VIJAYALAKSHMI CATERING EXPENSES BILL FOR THE THE PERIOD FROM"
   End If
   Cry_rep1.PrinterSelect
   If opt_canteen_expenses.Value = True Then
      Cry_rep1.Formulas(3) = ("canteen= '" & canteen & "'")
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_canteen_expenses.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{canteen_expenses.ce_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {canteen_expenses.ce_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")")
   ElseIf opt_date_empwise_recovery.Value = True Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_date_empwise_canteen_recovery.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{canteen_recovery.cr_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {canteen_recovery.cr_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")")
   
   ElseIf opt_empwise_recovery.Value = True Then
      Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\rpt_empwise_canteen_recovery.rpt"
      Cry_rep1.ReplaceSelectionFormula ("{canteen_recovery.cr_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {canteen_recovery.cr_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")")
   
   
   End If
   
 
   

   
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    st_date.Value = Now
    end_date.Value = Now
End Sub
