VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_duplicate_ticket 
   Caption         =   "Duplicate Ticket Printing"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   600
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   6375
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
         TabIndex        =   6
         Top             =   2520
         Width           =   915
      End
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
         TabIndex        =   5
         Top             =   2520
         Width           =   915
      End
      Begin VB.ComboBox cmb_TicketNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dt_ticket 
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   240
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
         Format          =   129564673
         CurrentDate     =   45262
      End
      Begin VB.Label Label2 
         Caption         =   "Ticket No"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ticket Date"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_duplicate_ticket"
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
            qry = "{trn_weighbridge_entry.t_wb_ticketno}  =  " & cmb_TicketNo.Text
            Cry_rep1.Formulas(0) = "opt = 1"
            MousePointer = vbDefault
            ''gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
           gst_repconnect = "dsn=mysql;uid=root;pwd=P@ssw0rD;database=shvpm"
            
''            Cry_rep1.PrinterSelect
''            Cry_rep1.ReportFileName = "d:\wbreports\weigh_slip.rpt"
            Cry_rep1.ReportFileName = "\\10.0.0.243\wbreports\weigh_slipTemp.rpt"
            Cry_rep1.ReplaceSelectionFormula (qry)
            Cry_rep1.WindowState = crptMaximized
            Cry_rep1.Connect = gst_repconnect
            
            Cry_rep1.Action = 1
End Sub

Function get_Tickets()
    cmb_TicketNo.Clear
    
    yy = 24
    compcode = 1
    
    pst_qry = "select * from trn_weighbridge_entry where t_wb_compcode = " & compcode & "  and  t_wb_year =  " & yy & "  and t_wb_date = '" & Format(dt_ticket, "yyyy/MM/dd") & "' "
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_TicketNo.AddItem (adors("t_wb_ticketno"))
                 adors.MoveNext
        Next
        adors.Close
    End If

End Function

Private Sub dt_ticket_Change()
  Call get_Tickets
End Sub

Private Sub Form_Load()
    dt_ticket.Value = Now
    
    Call gen_dbconnection
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    Call get_Tickets

End Sub
