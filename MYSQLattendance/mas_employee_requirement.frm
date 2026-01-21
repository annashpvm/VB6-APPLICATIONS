VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form mas_employee_requirement 
   Caption         =   "EMPLOYEE REQUIREMENT"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_all_dept 
      Caption         =   "ALL  DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13080
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmd_print_selective 
      Caption         =   "SELECTIVE DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   17
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   11775
      Begin VB.TextBox txt_tot_manpower 
         Height          =   375
         Left            =   8880
         TabIndex        =   15
         Top             =   5880
         Width           =   900
      End
      Begin VB.TextBox txt_tot_general 
         Height          =   375
         Left            =   7920
         TabIndex        =   14
         Top             =   5880
         Width           =   900
      End
      Begin VB.TextBox txt_tot_shiftC 
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   5880
         Width           =   900
      End
      Begin VB.TextBox txt_tot_shiftB 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   5880
         Width           =   900
      End
      Begin VB.TextBox txt_tot_shiftA 
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   5880
         Width           =   900
      End
      Begin VB.ComboBox cmb_dept 
         Height          =   315
         Left            =   5280
         TabIndex        =   8
         Top             =   360
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   4080
         TabIndex        =   1
         Top             =   7560
         Width           =   3855
         Begin VB.CommandButton exit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Exit"
            Height          =   705
            Left            =   3000
            MaskColor       =   &H000000FF&
            Picture         =   "mas_employee_requirement.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Refresh 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Refresh"
            Height          =   705
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "mas_employee_requirement.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Save"
            Height          =   705
            Left            =   1560
            MaskColor       =   &H000000FF&
            Picture         =   "mas_employee_requirement.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton edit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Edit"
            Height          =   705
            Left            =   840
            MaskColor       =   &H000000FF&
            Picture         =   "mas_employee_requirement.frx":1116
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton NEW 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&New"
            Height          =   705
            Left            =   120
            MaskColor       =   &H000000FF&
            Picture         =   "mas_employee_requirement.frx":1780
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flx_budget 
         Height          =   4335
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7646
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL"
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
         Left            =   3360
         TabIndex        =   10
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "DEPARTMENT"
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
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1935
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
   Begin VB.Label Label3 
      Caption         =   "MAN POWER REQUIREMENT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "mas_employee_requirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_dept_Click()
    If cmb_dept.ItemData(cmb_dept.ListIndex) > 0 Then
        flx_budget.Clear
        fillgrid
        Dim payrs As New ADODB.Recordset
        sql = "select pdesi_order,pdesi_name,pdesi_code from emp_mas , pdesi_mas where emp_status = 'A' and emp_design = pdesi_code and  emp_dept = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  group by pdesi_order,pdesi_name,pdesi_code order by pdesi_order"
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF()
            With flx_budget
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = payrs("pdesi_name")
                .TextMatrix(.Rows - 1, 1) = payrs("pdesi_code")
            End With
            payrs.MoveNext
        Wend
        payrs.Close
    
        sql = "select * from pmas_manpower where  m_deptcode = " & cmb_dept.ItemData(cmb_dept.ListIndex) & ""
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF()
              For i = 1 To flx_budget.Rows - 1
                 If Trim(flx_budget.TextMatrix(i, 1)) <> "" Then
                      If Val(flx_budget.TextMatrix(i, 1)) = payrs.Fields("m_desicode") Then
                         flx_budget.TextMatrix(i, 2) = payrs.Fields("m_shiftA")
                         flx_budget.TextMatrix(i, 3) = payrs.Fields("m_shiftB")
                         flx_budget.TextMatrix(i, 4) = payrs.Fields("m_shiftC")
                         flx_budget.TextMatrix(i, 5) = payrs.Fields("m_general")
                         flx_budget.TextMatrix(i, 6) = payrs.Fields("m_total")
                      End If
                 End If
             Next

            payrs.MoveNext
        Wend
        payrs.Close
        
       grid_tot
             
    End If
End Sub
Function grid_tot()
        txt_tot_shiftA.Text = 0
        txt_tot_shiftB.Text = 0
        txt_tot_shiftC.Text = 0
        txt_tot_general.Text = 0
        txt_tot_manpower.Text = 0
        Dim tot_ShiftA, tot_ShiftB, tot_ShiftC, tot_General, tot_manpower As Integer
        For i = 1 To flx_budget.Rows - 1
             tot_ShiftA = tot_ShiftA + Val(flx_budget.TextMatrix(i, 2))
             tot_ShiftB = tot_ShiftB + Val(flx_budget.TextMatrix(i, 3))
             tot_ShiftC = tot_ShiftC + Val(flx_budget.TextMatrix(i, 4))
             tot_General = tot_General + Val(flx_budget.TextMatrix(i, 5))
             tot_manpower = tot_manpower + Val(flx_budget.TextMatrix(i, 6))
        Next
        txt_tot_shiftA.Text = tot_ShiftA
        txt_tot_shiftB.Text = tot_ShiftB
        txt_tot_shiftC.Text = tot_ShiftC
        txt_tot_general.Text = tot_General
        txt_tot_manpower.Text = tot_manpower
End Function

Private Sub cmd_all_dept_Click()
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\manpower_departmentwise.rpt"
   Cry_rep1.ReplaceSelectionFormula ("")
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1
End Sub

Private Sub cmd_print_selective_Click()

   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   Cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\manpower_department.rpt"
   Cry_rep1.ReplaceSelectionFormula ("{pmas_manpower.m_deptcode} = " & cmb_dept.ItemData(cmb_dept.ListIndex))
   Cry_rep1.WindowState = crptMaximized
   Cry_rep1.Connect = gst_repconnect
   Cry_rep1.Action = 1

End Sub

Private Sub Command1_Click()

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_budget_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler

 Dim fin_selrow%, fin_selcol%
 fin_selrow = flx_budget.Row
 fin_selcol = flx_budget.Col
 With flx_budget

    If fin_selcol = 2 Or fin_selcol = 3 Or fin_selcol = 4 Or fin_selcol = 5 Then

           If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                'allow backspace and the enter keys
                    MsgBox "Enter OT hours in numbers only "
                    If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
                    .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                    End If
                    flx_budget.SetFocus
                Else
                
                    flx_budget.TextMatrix(fin_selrow, fin_selcol) = flx_budget.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)

                End If
           End If
           If KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
              .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              End If
              
           End If
        End If
        .TextMatrix(fin_selrow, 6) = Val(.TextMatrix(fin_selrow, 2)) + Val(.TextMatrix(fin_selrow, 3)) + Val(.TextMatrix(fin_selrow, 4)) + Val(.TextMatrix(fin_selrow, 5))
    
End With
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

End Sub

Private Sub Form_Load()

    
    Dim payrs As New ADODB.Recordset
    cmb_dept.Clear
    sql = "select * from pdept_mas where left(dept_name,2) != ' '  order by dept_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_dept.AddItem payrs("dept_name")
        cmb_dept.ItemData(cmb_dept.NewIndex) = payrs("dept_code")
        payrs.MoveNext
    Wend
    payrs.Close
    fillgrid

End Sub

Function fillgrid()
    With flx_budget
        .Clear
        .Cols = 7
        .Rows = 1
        .TextMatrix(0, 0) = "Designation"
        .TextMatrix(0, 1) = "Desi.code"
        .TextMatrix(0, 2) = "Shift - A"
        .TextMatrix(0, 3) = "Shift - B"
        .TextMatrix(0, 4) = "Shift - C"
        .TextMatrix(0, 5) = "General"
        .TextMatrix(0, 6) = "Total"
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
    End With
End Function



Private Sub save_Click()


Dim qry, cat As String
On Error GoTo err_handler
   If flx_budget.Rows < 2 Then
      MsgBox (" Details not available ")
      Exit Sub
   End If

  Me.MousePointer = 11
  Set payrs = New ADODB.Recordset
    
    sql = "delete from pmas_manpower where m_deptcode  = " & cmb_dept.ItemData(cmb_dept.ListIndex) & ""
    paydb.Execute sql
  
  sql = "select * from pmas_manpower where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 1 To flx_budget.Rows - 1
      If Val(flx_budget.TextMatrix(i, 6)) > 0 Then
            payrs.AddNew
            

        
            payrs.Fields("m_deptcode") = cmb_dept.ItemData(cmb_dept.ListIndex)
            payrs.Fields("m_desicode") = Val(flx_budget.TextMatrix(i, 1))
            payrs.Fields("m_shiftA") = Val(flx_budget.TextMatrix(i, 2))
            payrs.Fields("m_shiftB") = Val(flx_budget.TextMatrix(i, 3))
            payrs.Fields("m_shiftC") = Val(flx_budget.TextMatrix(i, 4))
            payrs.Fields("m_general") = Val(flx_budget.TextMatrix(i, 5))
            payrs.Fields("m_total") = Val(flx_budget.TextMatrix(i, 6))
            
            payrs.Update
      End If
  Next
  MsgBox ("Records are saved")
  payrs.Close
  fillgrid
  Me.MousePointer = 1
  Exit Sub
err_handler:
   
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
    
End Sub

Private Sub Text1_Change()

End Sub
