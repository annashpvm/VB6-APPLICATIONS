VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_eligible_leave_entry 
   Caption         =   "Eligible Leave entry"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_update 
      Caption         =   "GO"
      Height          =   315
      Left            =   8520
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txt_leave 
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox cmb_year 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   7200
      Width           =   3855
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_eligible_leave_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_eligible_leave_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "frm_eligible_leave_entry.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "frm_eligible_leave_entry.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "frm_eligible_leave_entry.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   153223169
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   153223169
         CurrentDate     =   39359
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
         TabIndex        =   4
         Top             =   240
         Width           =   1935
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
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   5610
      Left            =   1200
      TabIndex        =   12
      Top             =   1320
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9895
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   4
      BackColorFixed  =   16776960
      BackColorSel    =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Apply Leave for All Employees"
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
      Height          =   405
      Left            =   5280
      TabIndex        =   15
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "YEAR"
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
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   840
      Width           =   885
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   9735
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE ELIGIBLE LEAVE ENTRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frm_eligible_leave_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pst_qry As String
Dim new_entry_chk As Integer
Dim endrow As Byte
Dim emp_cat As String
Dim loadchk As Integer
Function fillgrid()
    With att_flex
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 4) = "EL.Leave"
        .TextMatrix(1, 4) = "for Year"
        .TextMatrix(0, 3) = "D.O.J"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
    End With
End Function

Function filldata()
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    
    If emptype_chk = 0 Then
           sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
           emp_cat = "S"
        ElseIf emptype_chk = 1 Then
           sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
           emp_cat = "W"
        ElseIf emptype_chk = 2 Then
           sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas where emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) " & loc & "   order by convert(int, EMP_CODE)")
           emp_cat = "M"
    End If
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = Format(payrs("emp_doj"), "dd/MM/yyyy")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    If emptype_chk = 2 Then Exit Function
    If emptype_chk = 0 Then
       sql = ("Select emp_code as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' " & loc & "  and EMP_CODE  like '%A'")
       emp_cat = "S"
    ElseIf emptype_chk = 1 Then
       sql = ("Select emp_code as ecode,* from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A'  and EMP_CODE  like '%A'")
       emp_cat = "W"
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = Format(payrs("emp_doj"), "dd/MM/yyyy")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
End Function

Private Sub cmd_update_Click()
  Dim i As Integer
    For i = 2 To att_flex.Rows - 1
       att_flex.TextMatrix(i, 4) = txt_leave.Text
    Next
End Sub

Private Sub NEW_Click()
  new_entry_chk = 0
  fillgrid
  If emptype_chk = 3 Then
     filldata_retainer
  Else
     filldata
  End If
 
End Sub

Private Sub edit_Click()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    If emptype_chk = 3 Then
       filldata_retainer
    Else
       filldata
    End If
    
    i = 2
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'"
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'"
''
''    ElseIf emptype_chk = 1 Then
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'"
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'"
''
''    ElseIf emptype_chk = 2 Then
''       sql = "select * from emp_eligible_leave a , emp_mas b  where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
''       sql = "select * from emp_eligible_leave a , emp_mas b  where s_company = " & company_code & " and s_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
''    End If
    
    If emptype_chk = 0 Then
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'"
       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'"
    
    ElseIf emptype_chk = 1 Then
''       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'"
       sql = "select * from emp_eligible_leave where s_company = " & company_code & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'"
    
    ElseIf emptype_chk = 2 Then
''       sql = "select * from emp_eligible_leave a , emp_mas b  where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
       sql = "select * from emp_eligible_leave a , emp_mas b  where s_company = " & company_code & " and s_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
    ElseIf emptype_chk = 3 Then
''       sql = "select * from emp_eligible_leave a , emp_mas b  where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
       sql = "select * from emp_eligible_leave  a , emp_voupay_mast b  where s_company = " & company_code & " and s_year = " & Val(cmb_year.Text) & " and emp_cat = 'R'   and s_company = emp_company and s_empcode = emp_code and s_empcat = emp_cat"
    End If
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 2 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                      If att_flex.TextMatrix(i, 1) = payrs.Fields("s_empcode") Then
                         att_flex.TextMatrix(i, 4) = IIf(payrs.Fields("s_el") > 0, payrs.Fields("s_el"), "")
                      End If
                End If
             Next
             payrs.MoveNext
       Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
        
    payrs.Close
        
End Sub

Public Sub load_data()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    
    i = 2
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  "
''    Else
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'"
''    End If
    
    If emptype_chk = 0 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  "
       sql = "select cast(emp_code as int) as ecode,* from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE not like '%A'   order by convert(int, EMP_CODE)"
    ElseIf emptype_chk = 1 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'"
       sql = "select  cast(emp_code as int) as ecode,* from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE not like '%A'   order by convert(int, EMP_CODE)"
    ElseIf emptype_chk = 2 Then
''       sql = "select * from attn_entry a , emp_mas b  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat"
       sql = "select * from attn_entry a , emp_mas b , pdept_mas c  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat and emp_dept = dept_code order by attn_empcode  "
    End If
    
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
            att_flex.Rows = att_flex.Rows + 1
            att_flex.TextMatrix(att_flex.Rows - 1, 0) = payrs.Fields("dept_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 1) = payrs.Fields("attn_empcode")
            att_flex.TextMatrix(att_flex.Rows - 1, 2) = payrs.Fields("emp_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 3) = payrs.Fields("attn_act_wdays")
            att_flex.TextMatrix(att_flex.Rows - 1, 4) = payrs.Fields("attn_work_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 5) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 6) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 7) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 8) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 9) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 10) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 11) = payrs.Fields("attn_salary_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 12) = payrs.Fields("attn_empcat")
        
        
''             For i = 2 To att_flex.Rows - 1
''                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
''                      If att_flex.TextMatrix(i, 1) = payrs.Fields("attn_empcode") Then
''                            att_flex.TextMatrix(i, 3) = payrs.Fields("attn_act_wdays")
''                            att_flex.TextMatrix(i, 4) = payrs.Fields("attn_work_days")
''                            att_flex.TextMatrix(i, 5) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
''                            att_flex.TextMatrix(i, 6) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
''                            att_flex.TextMatrix(i, 7) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
''                            att_flex.TextMatrix(i, 8) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
''                            att_flex.TextMatrix(i, 9) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
''                            att_flex.TextMatrix(i, 10) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
''                            att_flex.TextMatrix(i, 11) = payrs.Fields("attn_salary_days")
''                      End If
''                End If
''             Next
             payrs.MoveNext
        Wend
    End If
    payrs.Close
    If emptype_chk = 0 Then
       sql = "select * from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE like '%A' "
    ElseIf emptype_chk = 1 Then
       sql = "select * from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE like '%A' "
    End If
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
            att_flex.Rows = att_flex.Rows + 1
            att_flex.TextMatrix(att_flex.Rows - 1, 0) = payrs.Fields("dept_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 1) = payrs.Fields("attn_empcode")
            att_flex.TextMatrix(att_flex.Rows - 1, 2) = payrs.Fields("emp_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 3) = payrs.Fields("attn_act_wdays")
            att_flex.TextMatrix(att_flex.Rows - 1, 4) = payrs.Fields("attn_work_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 5) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 6) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 7) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 8) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 9) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 10) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 11) = payrs.Fields("attn_salary_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 12) = payrs.Fields("attn_empcat")
            payrs.MoveNext
        Wend
    End If
    payrs.Close
    
    
    
    sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       If payrs("pay_attn_lock") = "Y" Then
          save.Enabled = False
          lbl_disp.Caption = "Attendence Locked .. Can't Modify"
       End If
    Else
       lbl_disp.Caption = ""
       save.Enabled = True
    End If
    payrs.Close
       
    
End Sub

Private Sub exit_Click()
   Unload Me
End Sub
 
Private Sub Form_Load()
''    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''        .Text = "2015"
''    End With
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    
    new_entry_chk = 0
    
    If emptype_chk = 0 Then
       frm_eligible_leave_entry.Caption = "Yearly Eligible Leave entry for STAFF"
       lbl_emp.Caption = "STAFF LEAVE ELIGIBLE ENTRY"
    ElseIf emptype_chk = 1 Then
       frm_eligible_leave_entry.Caption = "Yearly Eligible Leave Entry for WORKER"
       lbl_emp.Caption = "WORKER LEAVE ELIGIBLE ENTRY"
    ElseIf emptype_chk = 2 Then
       frm_eligible_leave_entry.Caption = "Yearly Eligible Entry for Management"
       lbl_emp.Caption = "MANAGEMENT LEAVE ENTRY"
    ElseIf emptype_chk = 3 Then
       frm_eligible_leave_entry.Caption = "Yearly Eligible Entry for Retainer "
       lbl_emp.Caption = "RETAINER LEAVE ELIGIBLE ENTRY"
       
    End If
''  new_entry_chk = 0
''  attn_dt = Format(Now, "dd/mm/yyyy")
''  sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(attn_dt, "mm/dd/yyyy") & "'"
''  Set paydb = New ADODB.Connection
''  Set payrs = New ADODB.Recordset
''  paydb.Open pay
''  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''  If Not payrs.EOF Then
''     attstatus = payrs(1)
''  Else
''     attstatus = "PRESENT"
''  End If
''  endrow = 0
  loadchk = 0
  
  fillgrid
  If emptype_chk = 3 Then
     filldata_retainer
  Else
     filldata
  End If
  loadchk = 1
''  lst_code.Visible = False
''  lst_name.Visible = False
''  txt_itemname.Visible = False
''txt.Visible = False
End Sub

Private Sub att_flex_KeyPress(KeyAscii As Integer)
 If cmb_year.Text = "" Then
    MsgBox ("Select  Year....")
    Exit Sub
 End If
 On Error GoTo err_handler
 Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
 Select Case fin_selcol
        Case 4
        If KeyAscii <> 13 Then
              If fin_selcol = 3 Then
''                 KeyAscii = attndays_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
''              Else
                 KeyAscii = Numeric_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
            End If
            End If
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              KeyAscii = 0
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

Private Sub refresh_Click()
    fillgrid
    new_entry_chk = 0
    save.Enabled = True
End Sub
Private Sub SAVE_Click()

On Error GoTo err_handler
  If att_flex.Rows < 3 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
  Me.MousePointer = 11

''  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
''  paydb.Open pay
  paydb.BeginTrans
''  If emptype_chk = 0 Then
''        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'" & loc2
''        paydb.Execute sql
''  ElseIf emptype_chk = 1 Then
''        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'" & loc2
''        paydb.Execute sql
''  ElseIf emptype_chk = 2 Then
''        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and s_finyear = " & finyear & " and s_year = " & Val(cmb_year.Text) & " and s_empcat = 'M'" & loc2
''        paydb.Execute sql
''  End If

  If emptype_chk = 0 Then
        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'S'" & loc2
        paydb.Execute sql
  ElseIf emptype_chk = 1 Then
        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'W'" & loc2
        paydb.Execute sql
  ElseIf emptype_chk = 2 Then
        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'M'" & loc2
        paydb.Execute sql
  ElseIf emptype_chk = 3 Then
        sql = "delete from emp_eligible_leave where s_company = " & company_code & " and  s_year = " & Val(cmb_year.Text) & " and s_empcat = 'R'"
        paydb.Execute sql
  
  End If

  sql = "select * from emp_eligible_leave where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 2 To att_flex.Rows - 1

      If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
            payrs.AddNew
            payrs.Fields("s_company") = company_code
            payrs.Fields("s_finyear") = finyear
            payrs.Fields("s_year") = Val(cmb_year.Text)
            payrs.Fields("s_empcode") = att_flex.TextMatrix(i, 1)
            payrs.Fields("s_el") = Val(att_flex.TextMatrix(i, 4))
            If emptype_chk = 0 Then
               payrs.Fields("s_empcat") = "S"
            ElseIf emptype_chk = 1 Then
               payrs.Fields("s_empcat") = "W"
            ElseIf emptype_chk = 2 Then
               payrs.Fields("s_empcat") = "M"
            Else
               payrs.Fields("s_empcat") = "R"
            End If
            payrs.Update
      End If
  Next
  MsgBox ("Records are saved")
  paydb.CommitTrans
  payrs.Close
''  paydb.Close
  fillgrid
  Me.MousePointer = 1
  Exit Sub
err_handler:
    paydb.RollbackTrans
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
  
End Sub
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_handler
    If KeyCode = 40 Then
        lst_code.ListIndex = IIf(lst_code.ListIndex + 1 = lst_code.ListCount, lst_code.ListIndex, lst_code.ListIndex + 1)
        lst_code.SetFocus
    End If
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
      
Private Sub att_flex_EnterCell()
On Error GoTo err_handler
   Select Case att_flex.Col
        Case 4
''            txt.Left = att_flex.Left + att_flex.CellLeft
''            txt.Top = att_flex.Top + att_flex.CellTop
''            txt.Width = att_flex.CellWidth - 15
''            txt.Visible = True
''            lst_code.Left = att_flex.Left + att_flex.CellLeft
''            lst_code.Top = txt.Top + txt.Height
''            lst_code.Width = att_flex.CellWidth
''            lst_code.ListIndex = -1
''            txt = att_flex.Text
''            lst_code.Visible = True
''            txt_itemname.Visible = False
''            lst_name.Visible = False
''            txt.SetFocus
''      Case 4, 1, 2
'            txt.Visible = False
'            lst_code.Visible = False
    End Select
    Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

End Sub

Private Sub lst_code_DblClick()
On Error GoTo err_handler
     If lst_code.ListIndex <> -1 And lst_code.Tag = "" Then txt_KeyPress = 13
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
Private Sub lst_code_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
     If KeyAscii = 13 Then lst_code_DblClick
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub



Public Sub find_dates()

    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1
End Sub



Function filldata_retainer()
''    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
'''---
loc = ""
'''-
    
    sql = ("Select cast(emp_code as int) as ecode,* from  emp_voupay_mast where emp_company = '" & company_code & "' and emp_cat = 'R' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)")
    emp_cat = "R"
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = Format(payrs("emp_doj"), "dd/MM/yyyy")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
End Function

