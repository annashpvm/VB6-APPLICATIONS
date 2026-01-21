VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form pf_nominee 
   BackColor       =   &H00C0E0FF&
   Caption         =   "EMPLOYEE PF NOMINIEE MASTER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Exit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Exit"
      Height          =   855
      Left            =   1860
      Picture         =   "pfnominee.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7035
      Width           =   885
   End
   Begin VB.CommandButton save 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Save"
      Height          =   855
      Left            =   945
      Picture         =   "pfnominee.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7020
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid pf_flex 
      Height          =   2475
      Left            =   705
      TabIndex        =   7
      Top             =   4260
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   4366
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox emp_no 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4785
      TabIndex        =   5
      Top             =   2460
      Width           =   4260
   End
   Begin VB.ComboBox emp_cmb 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4785
      TabIndex        =   4
      Top             =   1740
      Width           =   6210
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employee pf details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3435
      Left            =   705
      TabIndex        =   0
      Top             =   405
      Width           =   10680
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   810
         Left            =   1770
         TabIndex        =   1
         Top             =   315
         Width           =   7260
         Begin VB.OptionButton opt_worker 
            BackColor       =   &H00C0E0FF&
            Caption         =   "WORKER"
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
            Height          =   345
            Left            =   3690
            TabIndex        =   3
            Top             =   300
            Width           =   2145
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00C0E0FF&
            Caption         =   "STAFF"
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
            Height          =   300
            Left            =   1335
            TabIndex        =   2
            Top             =   300
            Width           =   1440
         End
      End
      Begin VB.TextBox pf_no 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4095
         TabIndex        =   6
         Top             =   2745
         Width           =   4290
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "EMPLOYEE PF - NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   420
         TabIndex        =   10
         Top             =   2880
         Width           =   3465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "EMPLOYEE NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   405
         TabIndex        =   9
         Top             =   2145
         Width           =   3705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "EMPLOYEE NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   390
         TabIndex        =   8
         Top             =   1365
         Width           =   3390
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ENTER NOMINEE DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   705
      TabIndex        =   11
      Top             =   3915
      Width           =   3390
   End
End
Attribute VB_Name = "pf_nominee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paydb As New ADODB.Connection
Dim payrs As New ADODB.Recordset
Dim share As Integer
Dim endrow As Integer
Dim endrowchk As Integer
Private Sub emp_cmb_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  emp_mas where emp_name = '" & emp_cmb.Text & "' and emp_company = '" & company_code & "'")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       pf_no = payrs.Fields("emp_pfno")
       emp_no = payrs.Fields("emp_code")
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  emp_nominee_mas where e_empcode = " & Val(emp_no) & " and e_company = '" & company_code & "'")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    pf_flex.Clear
    fillgrid
    endrow = 0
    While Not payrs.EOF
         If payrs.Fields("e_empcode") = Val(emp_no) And payrs.Fields("e_company") = company_code Then
            With pf_flex
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = payrs.Fields("e_nominee")
            .TextMatrix(i, 1) = payrs.Fields("e_relation")
            .TextMatrix(i, 2) = payrs.Fields("e_ndob")
            .TextMatrix(i, 3) = payrs.Fields("e_share")
            i = i + 1
            endrow = endrow + 1
            End With
         End If
         payrs.MoveNext
    Wend
    endrowchk = endrow
End Sub

Private Sub exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    opt_staff.Value = True
    opt_staff_Click
End Sub

Private Sub opt_staff_Click()
    emp_no = ""
    pf_no = ""
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_type = 0 or emp_type = 1 order by emp_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    emp_cmb.Clear
    While Not payrs.EOF
        emp_cmb.AddItem payrs("emp_name")
      ''  emp_cmb.ItemData(emp_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    fillgrid
End Sub
Private Sub opt_worker_Click()
    emp_no = ""
    pf_no = ""
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_type = 2 or emp_type = 3 order by emp_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    emp_cmb.Clear
    While Not payrs.EOF
        emp_cmb.AddItem payrs("emp_name")
        emp_cmb.ItemData(emp_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    fillgrid
End Sub

Function fillgrid()
    With pf_flex
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "NOMINEE NAME"
        .TextMatrix(0, 1) = "NOMINEE RELATION SHIP"
        .TextMatrix(0, 2) = "DATE OF BIRTH"
        .TextMatrix(0, 3) = "SHARE %"
        .ColWidth(0) = 4500
        .ColWidth(1) = 3300
        .ColWidth(2) = 1600
        .ColWidth(3) = 1200
        .Rows = .Rows + 1
    End With
End Function

Private Sub pf_flex_DBLClick()
   fin_selrow = pf_flex.Row
   pst_ans = MsgBox("Are u sure want to delele YES Delete ", vbYesNo)
   If pst_ans = vbYes Then
      If pf_flex.Rows < 3 Then
         MsgBox "No rows to remove"
      Else
         pf_flex.RemoveItem fin_selrow
      End If
   End If
End Sub

Private Sub pf_flex_KeyPress(KeyAscii As Integer)
    On Error GoTo err_handler
    Dim fin_selrow%, fin_selcol%
    fin_selrow = pf_flex.Row
    fin_selcol = pf_flex.Col
    With pf_flex
        Select Case fin_selcol
            Case 0, 1
                If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                    endrow = fin_selrow
                    pf_flex.TextMatrix(fin_selrow, fin_selcol) = pf_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
                ElseIf KeyAscii = 8 Then
                  If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                  KeyAscii = 0
                End If
            Case 2
                If KeyAscii = 13 Then
                   Exit Sub
                End If
                If KeyAscii <> 13 Then
                   endrow = fin_selrow
                   KeyAscii = date_Chk(KeyAscii, pf_flex.TextMatrix(fin_selrow, fin_selcol))
                End If
                If KeyAscii <> 0 And KeyAscii <> 8 Then
                   endrow = fin_selrow
                   pf_flex.TextMatrix(fin_selrow, fin_selcol) = pf_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
                ElseIf KeyAscii = 8 Then
                  If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                  KeyAscii = 0
                End If
            Case 3
                If KeyAscii <> 13 Then
                   endrow = fin_selrow
                   KeyAscii = Numeric_Chk(KeyAscii, pf_flex.TextMatrix(fin_selrow, fin_selcol), 6, 3, 2)
                End If
                If KeyAscii <> 0 And KeyAscii <> 8 Then
                   endrow = fin_selrow
                   pf_flex.TextMatrix(fin_selrow, fin_selcol) = pf_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
                ElseIf KeyAscii = 8 Then
                  If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                  KeyAscii = 0
                End If
                If KeyAscii = 13 Then
                   If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                   KeyAscii = 0
                   .Rows = .Rows + 1
                   endrow = fin_selrow
                   Exit Sub
                End If
        End Select
    End With
    Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
       If chk = 1 Then
           Resume
       End If
End Sub

Private Sub SAVE_Click()
    If Trim(emp_cmb.Text) = "" Then
       MsgBox ("Select Employee Name")
       Exit Sub
    End If
    share = 0
    If endrowchk > endrow Then endrow = endrowchk
    For i = 1 To endrow
       share = share + Val(pf_flex.TextMatrix(i, 3))
    Next
    If share <> 100# Then
       MsgBox ("SHARE % is not equal to 100. Please check it ")
       Exit Sub
    End If
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  emp_nominee_mas where e_empcode = " & Val(emp_no) & " and e_company = '" & company_code & "'")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
            payrs.Delete
            payrs.Update
            payrs.MoveNext
       Wend
    End If
    For i = 1 To endrow
        If Trim(pf_flex.TextMatrix(i, 0)) <> "" Then
            payrs.AddNew
            payrs.Fields("e_company") = company_code
            payrs.Fields("e_empcode") = Val(emp_no)
            payrs.Fields("e_nominee") = pf_flex.TextMatrix(i, 0)
            payrs.Fields("e_relation") = pf_flex.TextMatrix(i, 1)
            payrs.Fields("e_ndob") = Format(pf_flex.TextMatrix(i, 2), "dd/mm/yyyy")
            payrs.Fields("e_share") = pf_flex.TextMatrix(i, 3)
            payrs.Update
        End If
    Next
    emp_cmb.Text = " "
    emp_no.Text = " "
    pf_no.Text = " "
    fillgrid
End Sub
