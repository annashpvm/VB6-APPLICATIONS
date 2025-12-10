VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_deduction_individual 
   Caption         =   "DEDUCTION ENTRY"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_import 
      Caption         =   "IMPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txt_total 
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
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   -240
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   120913921
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   120913921
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   8655
      Begin VB.ComboBox cmb_deduction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   13
         Top             =   720
         Width           =   5085
      End
      Begin VB.ComboBox cmb_month 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   10
         Top             =   300
         Width           =   2655
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
         Left            =   6120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label ded_label 
         Caption         =   "DEDUCTION"
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
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "YEAR"
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
         Height          =   285
         Left            =   5445
         TabIndex        =   12
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
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
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   8520
      Width           =   4935
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   825
         Left            =   0
         MaskColor       =   &H000000FF&
         Picture         =   "frm_deduction_individual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   825
         Left            =   960
         MaskColor       =   &H000000FF&
         Picture         =   "frm_deduction_individual.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   825
         Left            =   1920
         MaskColor       =   &H000000FF&
         Picture         =   "frm_deduction_individual.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   825
         Left            =   2880
         MaskColor       =   &H000000FF&
         Picture         =   "frm_deduction_individual.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   825
         Left            =   3900
         MaskColor       =   &H000000FF&
         Picture         =   "frm_deduction_individual.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   5490
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   9684
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
   Begin VB.Label Label3 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   7800
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   6375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   8655
   End
   Begin VB.Label lbl_emp 
      Caption         =   "EMPLOYEE DEDUCTENCE ENTRY"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frm_deduction_individual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Byte
Dim emp_cat As String

Function sumvalue()
     txt_total.Text = ""
     Dim value1 As Double
     value1 = 0
     For i = 1 To att_flex.Rows - 1
          value1 = value1 + Val(att_flex.TextMatrix(i, 4))
     Next
     txt_total.Text = value1
End Function
Function fillgrid()

    With att_flex
        .Clear
        .Cols = 5
        .Rows = 1

        .TextMatrix(0, 0) = "Emp code"
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Department"
        .TextMatrix(0, 3) = "CAT"
        .TextMatrix(0, 4) = "Deduction"
                
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 500
        .ColWidth(4) = 1000
    End With
    
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1)order by emp_name")
''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' " & loc & "   order by emp_doj")
''       emp_cat = "S"
''    Else
''''       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3)order by emp_name")
''        sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' " & loc & "   order by emp_doj")
''       emp_cat = "W"
''    End If
''
    
     sql = ("Select * from  emp_mas where emp_company = '" & company_code & "'  and emp_status = 'A' and emp_code >= 1000 order by emp_code")
     
     sql = "Select * from  emp_mas ,pdept_mas where emp_dept = dept_code and emp_company = '" & company_code & "'   and emp_status = 'A' and emp_code >= 1000 order by dept_name,emp_code"
    
    
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              ''find_deptname (payrs.Fields("emp_dept"))
             
             .TextMatrix(.Rows - 1, 0) = payrs("emp_code")
             .TextMatrix(.Rows - 1, 1) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 2) = payrs("dept_name")
             .TextMatrix(.Rows - 1, 3) = payrs("emp_cat")
             
        End With
        payrs.MoveNext
    Wend

End Function

Private Sub cmb_deduction_Click()
    If cmb_deduction.Text = "MESS" Then
       cmd_import.Enabled = True
    Else
       cmd_import.Enabled = False
    End If
    edit_Click
    
    
End Sub

Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
    find_dates
End Sub


Private Sub cmd_import_Click()

    new_entry_chk = 1
    endrow = 0
    fillgrid
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    If cmb_deduction.Text = "" Then
       MsgBox ("Select Deduction Name...")
       Exit Sub
    End If
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = "select * from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'S' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
''    Else
''       sql = "select * from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'W' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
''    End If
    
    sql = " select emp_fpcode,emp_name,emp_cat,sum(messamt+cr_others) as messamt1 from ( " _
          & " select emp_fpcode,emp_name,emp_cat, case when cr_foodtype = 'B' or cr_foodtype = 'D' then 15 else 20 end as messamt, cr_others   from canteen_recovery , emp_mas where cr_empcode = emp_fpcode  and month(cr_date) = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and year(cr_date) = " & Val(cmb_year.Text) & " group by emp_fpcode,emp_name ,cr_foodtype,emp_cat, cr_others " _
          & " ) aa group by emp_fpcode,emp_name,emp_cat"


 sql = " select emp_fpcode,emp_name,emp_cat, cr_foodtype,case when cr_foodtype = 'B' or cr_foodtype = 'D' then 15 else 20 end as messamt, sum(nos) as nos , sum(cr_others) as cr_others   from ( " _
          & "select emp_fpcode,emp_name,emp_cat,cr_foodtype,count(*) as nos ,sum(cr_others) as cr_others  from canteen_recovery , emp_mas where cr_empcode = emp_fpcode  and month(cr_date) =  " & cmb_month.ItemData(cmb_month.ListIndex) & " and year(cr_date) = " & Val(cmb_year.Text) & "  group by emp_fpcode,emp_name,emp_cat,cr_foodtype ) a group by emp_fpcode,emp_name,emp_cat,cr_foodtype "


 sql = " select emp_fpcode,emp_name,emp_cat, cr_foodtype,case when cr_foodtype = 'B' or cr_foodtype = 'D' then 15 else case when cr_foodtype = 'L' then 20 else 0 end end as messamt, sum(nos) as nos , sum(cr_others) as cr_others   from ( " _
          & "select emp_fpcode,emp_name,emp_cat,cr_foodtype,count(*) as nos ,sum(cr_others) as cr_others  from canteen_recovery , emp_mas where cr_empcode = emp_fpcode  and month(cr_date) =  " & cmb_month.ItemData(cmb_month.ListIndex) & " and year(cr_date) = " & Val(cmb_year.Text) & "  group by emp_fpcode,emp_name,emp_cat,cr_foodtype ) a group by emp_fpcode,emp_name,emp_cat,cr_foodtype "





    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                      If att_flex.TextMatrix(i, 0) = payrs.Fields("emp_fpcode") Then
                            att_flex.TextMatrix(i, 4) = Val(att_flex.TextMatrix(i, 4)) + (payrs.Fields("messamt") * payrs.Fields("nos")) + payrs.Fields("cr_others")
                      End If
                End If
             Next
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
     sumvalue

End Sub

''Private Sub attn_dt_Change()
''      sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(attn_dt, "mm/dd/yyyy") & "'"
''      Set paydb = New ADODB.Connection
''      Set payrs = New ADODB.Recordset
''      paydb.Open pay
''      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''      If Not payrs.EOF Then
''         attstatus = payrs(1)
''      Else
''         attstatus = "EARNED LEAVE (FULL DAY)"
''      End If
''      endrow = 0
''      fillgrid
''      Set paydb = New ADODB.Connection
''      Set payrs = New ADODB.Recordset
''      If emptype_chk = 0 Then
''         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
''      Else
''         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
''      End If
''      paydb.Open pay
''      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''      If payrs.EOF Then
''         MsgBox ("Pervious date details are missing. First enter for previous date & continue")
''         attn_dt = attn_dt - 1
''      End If
''      i = 1
''      If emptype_chk = 0 Then
''         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
''      Else
''         sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
''      End If
''      Set paydb = New ADODB.Connection
''      Set payrs = New ADODB.Recordset
''      paydb.Open pay
''      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''      If Not payrs.EOF Then
''         While Not payrs.EOF
''              With att_flex
''                   .Rows = .Rows + 1
''                   find_empdetails (payrs.Fields("attn_empcode"))
''                   find_attnstatus (payrs.Fields("attn_status"))
''                   .TextMatrix(i, 0) = dept_name
''                   .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
''                   .TextMatrix(i, 2) = employee_name
''                    att_dat = attn_dt
''                    find_present_status (payrs(0))
''                   .TextMatrix(i, 3) = attstatus
''                   .TextMatrix(i, 4) = payrs(5)
''                   i = i + 1
''                   endrow = endrow + 1
''              End With
''              payrs.MoveNext
''         Wend
''      Else
''         Set paydb = New ADODB.Connection
''         Set payrs = New ADODB.Recordset
''         If emptype_chk = 0 Then
''            sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 0 or emp_type = 1) order by emp_name")
''         Else
''            sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and (emp_type = 2 or emp_type = 3) order by emp_name")
''         End If
''         paydb.Open pay
''         payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''         payrs.MoveFirst
''         While Not payrs.EOF
''               With att_flex
''                   .Rows = .Rows + 1
''                    find_deptname (payrs.Fields("emp_dept"))
''                   .TextMatrix(.Rows - 1, 0) = dname
''                   .TextMatrix(.Rows - 1, 1) = payrs(0)
''                   .TextMatrix(.Rows - 1, 2) = payrs(5)
''                   If Trim((payrs.Fields("emp_holiday"))) = UCase(RTrim(Format(attn_dt, "dddd"))) Then
''                      .TextMatrix(.Rows - 1, 3) = "WEEKLY OFF"
''                   End If
''                    att_dat = attn_dt
''                    find_present_status (payrs(0))
''                   .TextMatrix(.Rows - 1, 3) = attstatus
''                   endrow = endrow + 1
''               End With
''               payrs.MoveNext
''          Wend
''      End If
'''      lst_code.Clear
'''      Set paydb = New ADODB.Connection
'''      Set payrs = New ADODB.Recordset
'''      sql = ("Select * from attn_status_mas")
'''      paydb.Open pay
'''      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
'''      While Not payrs.EOF
'''            lst_code.AddItem payrs(1)
'''            lst_code.ItemData(lst_code.ListCount - 1) = payrs(0)
'''            payrs.MoveNext
'''      Wend
'''      lst_code.ListIndex = -1
''' Sub

''Private Sub attn_dt_Click()
''   Set paydb = New ADODB.Connection
''   Set payrs = New ADODB.Recordset
''   If emptype_chk = 0 Then
''      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
''   Else
''      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt - 1, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
''   End If
''   paydb.Open pay
''   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''   If payrs.EOF Then
''      MsgBox ("Pervious date details are missing. First enter for previous date & continue")
''      attn_dt = attn_dt - 1
''   End If
''   If emptype_chk = 0 Then
''      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 0 or attn_emptype = 1)")
''   Else
''      sql = ("Select * from  attn_entry where attn_date = '" & Format(attn_dt, "mm/dd/yyyy") & "' and attn_company = '" & company_code & "' and (attn_emptype = 2 or attn_emptype = 3)")
''   End If
''   Set paydb = New ADODB.Connection
''   Set payrs = New ADODB.Recordset
''   paydb.Open pay
''   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''   If Not payrs.EOF Then
''      While Not payrs.EOF
''            With att_flex
''                 .Rows = .Rows + 1
''                 find_empdetails (payrs.Fields("attn_empcode"))
''                 find_attnstatus (payrs.Fields("attn_status"))
''                 .TextMatrix(i, 0) = dept_name
''                 .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
''                 .TextMatrix(i, 2) = employee_name
''                  att_dat = attn_dt
''                  find_present_status (payrs(0))
''                 .TextMatrix(i, 3) = attstatus
''                 .TextMatrix(i, 4) = payrs(5)
''                 i = i + 1
''                 endrow = endrow + 1
''            End With
''            payrs.MoveNext
''      Wend
''    Else
''       MsgBox ("Details not available for the date ")
''    End If
''    new_entry_chk = 0
''End Sub

Private Sub NEW_Click()
  new_entry_chk = 0
  fillgrid
  sumvalue
 ''ttn_dt.SetFocus
End Sub

Private Sub edit_Click()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    If cmb_deduction.Text = "" Then
       MsgBox ("Select Deduction Name...")
       Exit Sub
    End If
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = "select * from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'S' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
''    Else
''       sql = "select * from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'W' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
''    End If
    
    sql = "select * from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
    
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                      If att_flex.TextMatrix(i, 0) = payrs.Fields("e_emp_code") Then
                            att_flex.TextMatrix(i, 4) = payrs.Fields("e_ded_amount")
                      End If
                End If
             Next
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
     sumvalue
        
''              With att_flex
''                  .Rows = .Rows + 1
''                  find_empdetails (payrs.Fields("attn_empcode"))
''                  .TextMatrix(i, 0) = dept_name
''                  .TextMatrix(i, 1) = payrs.Fields("attn_empcode")
''                  .TextMatrix(i, 2) = employee_name
''                  .TextMatrix(i, 3) = payrs.Fields("attn_act_wdays")
''                  .TextMatrix(i, 4) = payrs.Fields("attn_work_days")
''                  .TextMatrix(i, 5) = payrs.Fields("attn_el")
''                  .TextMatrix(i, 6) = payrs.Fields("attn_pl")
''                  .TextMatrix(i, 7) = payrs.Fields("attn_abs")
''                  .TextMatrix(i, 8) = payrs.Fields("attn_layoff")
''                  .TextMatrix(i, 9) = payrs.Fields("attn_dec_holiday")
''                  .TextMatrix(i, 10) = payrs.Fields("attn_salary_days")
''                  i = i + 1
''                  endrow = endrow + 1
''             End With
''             payrs.MoveNext
''        Wend
''     Else
''        MsgBox ("Details not available for the date ")
''     End If
End Sub


Private Sub exit_Click()
   Unload Me
End Sub
 
Private Sub Form_Load()
    cmd_import.Enabled = False
    Dim payrs As New ADODB.Recordset
    sql = "Select * from pdedu_mas"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_deduction.AddItem payrs(1)
        cmb_deduction.ItemData(cmb_deduction.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
    With cmb_month
        .AddItem "January"
        .ItemData(.NewIndex) = 1
        .AddItem "February"
        .ItemData(.NewIndex) = 2
        .AddItem "March"
        .ItemData(.NewIndex) = 3
        .AddItem "April"
        .ItemData(.NewIndex) = 4
        .AddItem "May"
        .ItemData(.NewIndex) = 5
        .AddItem "June"
        .ItemData(.NewIndex) = 6
        .AddItem "July"
        .ItemData(.NewIndex) = 7
        .AddItem "August"
        .ItemData(.NewIndex) = 8
        .AddItem "September"
        .ItemData(.NewIndex) = 9
        .AddItem "October"
        .ItemData(.NewIndex) = 10
        .AddItem "November"
        .ItemData(.NewIndex) = 11
        .AddItem "December"
        .ItemData(.NewIndex) = 12
    End With
''    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''
''    End With
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
   new_entry_chk = 0
   lbl_emp.Caption = "DEDUCTION ENTRY"


    
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
  fillgrid
''  lst_code.Visible = False
''  lst_name.Visible = False
''  txt_itemname.Visible = False
''txt.Visible = False
End Sub

Private Sub att_flex_KeyPress(KeyAscii As Integer)
 If cmb_month.Text = "" Or cmb_year.Text = "" Then
    MsgBox ("Select Month / Year....")
    Exit Sub
 End If
 On Error GoTo err_handler
 Dim layoffdays As Double
 Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
 Select Case fin_selcol
        Case 4
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 8, 5, 2)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              KeyAscii = 0
            End If
        End If
    End Select
 End With
 sumvalue
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
End Sub
Private Sub refresh_Click()
    fillgrid
    sumvalue
    new_entry_chk = 0
End Sub
Private Sub SAVE_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
       MsgBox "Out Of Financial Year", vbInformation, "Message"
       Exit Sub
    End If
    Dim attn_Days As Double
    attn_Days = 0
    On Error GoTo err_handler
''    If endrow = 0 Then
''         MsgBox (" Details not available ")
''         Exit Sub
''    End If
    If cmb_month.Text = "" Then
         MsgBox (" Select Month ")
         Exit Sub
    End If
    If cmb_year.Text = "" Then
         MsgBox (" Select Year ")
         Exit Sub
    End If
    If cmb_deduction.Text = "" Then
         MsgBox (" Select Deduction ")
         Exit Sub
    End If
    Me.MousePointer = 11
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    paydb.BeginTrans
''    If emptype_chk = 0 Then
''       sql = "delete from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'S' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & " " & loc & ""
''       sql = "delete from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'S' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & " "
''    Else
''       sql = "delete from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'W' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & " " & loc & "  "
''       sql = "delete from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_emp_cat = 'W' and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & " "
''    End If
''
    sql = "delete from monthly_Deduction where e_company = " & company_code & " and e_finyear = " & finyear & " and e_ded_year = " & Val(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and e_ded_code = " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & ""
    paydb.Execute sql
    For i = 1 To att_flex.Rows - 1
          If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
''             If emptype_chk = 0 Then
''                sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_empcat = 'S' and attn_year = " & Val(cmb_year.Text) & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcode = '" & att_flex.TextMatrix(i, 1) & "'"
''             Else
''                sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_empcat = 'S' and attn_year = " & Val(cmb_year.Text) & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcode = '" & att_flex.TextMatrix(i, 1) & "'"
''             End If
             sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and  attn_year = " & Val(cmb_year.Text) & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_empcode = '" & att_flex.TextMatrix(i, 0) & "'"
             
             payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
             While Not payrs.EOF
                   attn_Days = payrs.Fields("attn_act_wdays")
                   payrs.MoveNext
             Wend
             payrs.Close
             
''             If emptype_chk = 0 Then
''                sql2 = "insert into monthly_Deduction values ( " & company_code & ", " & finyear & ", '" & att_flex.TextMatrix(i, 1) & "' ,'S', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & "," & Val(att_flex.TextMatrix(i, 3)) & ",  " & attn_Days & " )"
''             Else
''                sql2 = "insert into monthly_Deduction values ( " & company_code & ", " & finyear & ", '" & att_flex.TextMatrix(i, 1) & "' ,'W', " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & "," & Val(att_flex.TextMatrix(i, 3)) & ", " & attn_Days & ")"
''             End If
             
             sql2 = "insert into monthly_Deduction values ( " & company_code & ", " & finyear & ", '" & att_flex.TextMatrix(i, 0) & "' ,'" & att_flex.TextMatrix(i, 3) & "' , " & cmb_year.Text & ",  " & cmb_month.ItemData(cmb_month.ListIndex) & " ,  " & cmb_deduction.ItemData(cmb_deduction.ListIndex) & "," & Val(att_flex.TextMatrix(i, 4)) & ", " & attn_Days & ")"
             paydb.Execute sql2
          End If
    Next
    MsgBox ("Records are saved")
    paydb.CommitTrans
    fillgrid
    sumvalue
    Me.MousePointer = 1
    Exit Sub
err_handler:
    paydb.RollbackTrans
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
Dim ret%, pst_row%, pst_rawname$
 Static PrevIndex%
lst_code.Tag = "Keypress"
    Select Case KeyAscii
        Case 8
            If Trim(fst_item) <> "" Then fst_item = Mid(fst_item, 1, Len(fst_item) - 1)
        Case 13
             pbl_status = True
             If lst_code.ListIndex <> -1 Then pst_rawname = lst_code.Text
             For pin_cnt = 1 To att_flex.Rows - 1
                If pin_cnt <> att_flex.Row Then If LCase(att_flex.TextMatrix(pin_cnt, 3)) = LCase(pst_rawname) Then pbl_status = False
             Next
                pst_row = att_flex.Row
                If lst_code.ListIndex <> -1 Then
                    att_flex.TextMatrix(pst_row, 3) = lst_code.Text
                 att_flex.Col = 1
                 att_flex.SetFocus
                 lst_code.Tag = ""
                 Exit Sub
               End If
        Case Else
            fst_item = txt & Chr(KeyAscii)
            If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    End Select
    
    PrevIndex = lst_code.ListIndex
    If Trim(fst_item) = "" Then
        lst_code.ListIndex = -1
    Else
        ret = SendMessage(lst_code.hwnd, LB_FINDSTRING, -1, fst_item)
        If ret = -1 Then
            lst_code.ListIndex = PrevIndex
        Else
            lst_code.ListIndex = ret
        End If
    End If
    lst_code.Tag = ""
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If
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
        Case 5, 1, 2
'            txt.Visible = False
'            lst_code.Visible = False
    End Select
    sumvalue
    Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

End Sub

Private Sub lst_code_DblClick()
On Error GoTo err_handler
     If lst_code.ListIndex <> -1 And lst_code.Tag = "" Then txt_KeyPress 13
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
    If cmb_year.Text = "" Then Exit Sub
    Dim mdays, diff As Integer
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


