VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_manual_inout_punch_head 
   Caption         =   "Manual IN/OUT PUNCH"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   18015
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3600
      TabIndex        =   17
      Top             =   5760
      Width           =   2175
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_manual_inout_punch_head.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_manual_inout_punch_head.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      Begin VB.TextBox txt_fpcode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign Manual In / Out inpunch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         TabIndex        =   16
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   405
         Left            =   5760
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   8175
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
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
            CustomFormat    =   "dd/MM/yyyy HH:MM"
            Format          =   55967745
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dt_to 
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
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
            CustomFormat    =   "dd/MM/yyyy HH:MM"
            Format          =   55967745
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dtout 
            Height          =   375
            Left            =   5640
            TabIndex        =   6
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   55967746
            CurrentDate     =   41387.3333333333
         End
         Begin MSComCtl2.DTPicker dtIn 
            Height          =   375
            Left            =   5640
            TabIndex        =   7
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   55967746
            CurrentDate     =   41387.7083333333
         End
         Begin VB.Label Label5 
            Caption         =   "To time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4320
            TabIndex        =   11
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "From time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4320
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lbl2 
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Fing. Pass Code"
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
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Emp.Name"
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
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp. Code"
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
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Employes"
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
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_manual_inout_punch_head"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Assign_Click()
paydb.BeginTrans
On Error GoTo err_handler

    Dim idate As Date
    Dim sql As String

    pst_qry = "select max(bio_logid)+1 as endno from bio_manual_logs"
    payrs.Open pst_qry, paydb, 1, 2
    no = 1
    If Not IsNull(payrs!endno) Then
        If Not payrs.EOF Then
             no = payrs!endno
        End If
    End If
    payrs.Close

    For idate = dt_from To dt_to
    
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_empcode.Text & ",'" & no & "', '" & Format(idate, "MM/dd/yyyy") & "', '" & Format(idate, "MM/dd/yyyy 08:00:00") & "','M','N','in')"
        paydb.Execute sql
        no = no + 1
        
        pst_qry = "update bio_manual_logs set bio_logid = " & no
        paydb.Execute pst_qry
        
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_empcode.Text & ",'" & no & "', '" & Format(idate, "MM/dd/yyyy") & "', '" & Format(idate, "MM/dd/yyyy 17:00:00") & "','M','N','out')"
        paydb.Execute sql
        no = no + 1
        
        pst_qry = "update bio_manual_logs set bio_logid = " & no
        paydb.Execute pst_qry
        
    Next


    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
  
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dt_from.Value = Now
    dt_to.Value = Now

     dtIn.Value = "08.30 AM"
     dtout.Value = "05.30 PM"
    
    
    Dim payrs As New ADODB.Recordset
    lst_employee.Clear

    sql = "Select * from  bio_empmas where  bioemp_fpcode in (10)"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        payrs.MoveNext
    Wend
    payrs.Close
    If adminpw = 0 Then
       cmd_Assign.Enabled = False
    End If

End Sub

Private Sub lst_employee_Click()
    Dim payrs As New ADODB.Recordset

    sql = "Select * from  bio_empmas where  bioemp_name = '" & Trim(lst_employee.Text) & "'"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        txt_empcode.Text = payrs("bioemp_id")
        txt_fpcode.Text = payrs("bioemp_fpcode")
        txt_empname.Text = payrs("bioemp_name")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub
