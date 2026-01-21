VERSION 5.00
Begin VB.Form frm_bank_master 
   Caption         =   "BANK MASTER"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   5055
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Refresh"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Edit"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&New"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
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
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   9615
      Begin VB.TextBox txt_bank 
         Height          =   495
         Left            =   3240
         TabIndex        =   2
         Top             =   2520
         Width           =   5895
      End
      Begin VB.ComboBox cmb_bank 
         Height          =   1935
         Left            =   3120
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name to be Changed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_bank_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fst_save As String
Dim pst_qry As String

Private Sub cmd_Edit_Click()
fst_save = "EDIT"
Label2.Visible = True
txt_bank.Visible = True
End Sub

Private Sub cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_New_Click()
Label2.Visible = False
txt_bank.Visible = False
End Sub

Private Sub cmd_Refresh_Click()
Label2.Visible = False
txt_bank.Visible = False
    cmb_bank.Clear
    Dim rs_pay  As New ADODB.Recordset
    pst_qry = "select * from payroll_bank order by bank_name"
    rs_pay.Open pst_qry, paydb, 1, 2
    While Not rs_pay.EOF
        cmb_bank.AddItem rs_pay("bank_name")
        rs_pay.MoveNext
    Wend
    rs_pay.Close
End Sub

Private Sub cmd_Save_Click()
paydb.BeginTrans
On Error GoTo err_handler
Dim pst_respo As String
Dim pst_qry As String
Dim no As Integer
If fst_save = "NEW" Then
    If cmb_bank.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Select Bank Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_bank.SetFocus
        Exit Sub
    End If
    pst_respo = MsgBox("Do You want to save the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        paydb.RollbackTrans
        MousePointer = vbDefault
        Exit Sub
    End If
    Dim rs_pay As New ADODB.Recordset
    pst_qry = "select * from payroll_bank where bank_name='" & cmb_bank.Text & "'"
    rs_pay.Open pst_qry, paydb, 1, 2
        If Not (rs_pay.EOF) Then
            rs_pay.Close
            paydb.RollbackTrans
            MsgBox " Details Already Exists ..", vbOKOnly + vbExclamation, "Message"
            MousePointer = vbDefault
            Exit Sub
        End If
    rs_pay.Close
    pst_qry = "select max(bank_code)+1 as endno from payroll_bank"
    rs_pay.Open pst_qry, paydb, 1, 2
    no = 1
    If Not IsNull(rs_pay!endno) Then
        If Not rs_pay.EOF Then
             no = rs_pay!endno
        End If
    End If
    rs_pay.Close
    pst_qry = "insert into payroll_bank values(" & no & ",'" & cmb_bank.Text & "')"
    paydb.Execute pst_qry
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    cmb_bank.Text = ""
    Exit Sub
ElseIf fst_save = "EDIT" Then
    If cmb_bank.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Select Bank Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_bank.SetFocus
        Exit Sub
    End If
    If txt_bank.Text = "" Then
        paydb.RollbackTrans
        MsgBox " Enter Bank Name to be Change as ", vbOKOnly + vbExclamation, "vbInformation "
        txt_bank.SetFocus
        Exit Sub
      End If
    pst_respo = MsgBox("Do You want to Modify the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        paydb.RollbackTrans
        MousePointer = vbDefault
        Exit Sub
    End If

    pst_qry = "update payroll_bank set bank_name='" & txt_bank.Text & "' where bank_name='" & cmb_bank.Text & "'"
    paydb.Execute pst_qry
    MsgBox "Records Updated", vbOKOnly + vbInformation, "vbInformation"
    cmb_bank.Text = ""
    txt_bank.Text = ""
    paydb.CommitTrans
End If
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
fst_save = "NEW"
Label2.Visible = False
txt_bank.Visible = False
    Dim rs_pay  As New ADODB.Recordset
    pst_qry = "select * from payroll_bank order by bank_name"
    rs_pay.Open pst_qry, paydb, 1, 2
    While Not rs_pay.EOF
        cmb_bank.AddItem rs_pay("bank_name")
        rs_pay.MoveNext
    Wend
    rs_pay.Close
End Sub


