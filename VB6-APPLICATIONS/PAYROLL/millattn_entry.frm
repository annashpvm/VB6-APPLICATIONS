VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form millattn_entry 
   Caption         =   "MILL ATTENDANCE ENTRY"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   4440
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txt_filename 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton cmd_select 
         Caption         =   "Select file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_upload 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "FILE PATH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   2400
      TabIndex        =   14
      Top             =   9240
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56426497
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56426497
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   9480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   7200
      Width           =   3855
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "millattn_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "millattn_entry.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "millattn_entry.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "millattn_entry.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "millattn_entry.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
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
      Left            =   4080
      TabIndex        =   2
      Top             =   600
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
      Left            =   8400
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   5610
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   9895
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   5
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl_disp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   8400
      Width           =   3975
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE ATTENDANCE ENTRY"
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
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   10695
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
      Left            =   7245
      TabIndex        =   4
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "MONTH"
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
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   13215
   End
End
Attribute VB_Name = "millattn_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim xl As New Excel.Application
''Dim xlsheet As Excel.Worksheet
''Dim xlwbook As Excel.Workbook
Dim pst_qry As String

Dim mdays, diff As Integer
Dim new_entry_chk As Integer
Dim fst_item$
Dim endrow As Byte
Dim emp_cat As String
Dim loadchk As Integer

Function fillgrid()
    With att_flex
        .Clear
        .Cols = 17
        .Rows = 2
        .TextMatrix(1, 0) = "Department"
        .TextMatrix(1, 1) = "Emp code"
        .TextMatrix(1, 2) = "Name"
        .TextMatrix(0, 3) = "Total "
        .TextMatrix(1, 3) = "EL  "
        .TextMatrix(0, 4) = "Upto  "
        .TextMatrix(1, 4) = "Leave"
        
        .TextMatrix(0, 5) = "Actual"
        .TextMatrix(1, 5) = "Work.Days"
        .TextMatrix(0, 6) = "Worked"
        .TextMatrix(1, 6) = " Days "
        .TextMatrix(0, 7) = "Eligible"
        .TextMatrix(1, 7) = " Leave "
        .TextMatrix(0, 8) = "Permiss"
        .TextMatrix(1, 8) = " Leave "
        .TextMatrix(0, 9) = "       "
        .TextMatrix(1, 9) = "Absent "
        .TextMatrix(0, 10) = "       "
        .TextMatrix(1, 10) = "Lay Off"
        .TextMatrix(0, 11) = "Decl.  "
        .TextMatrix(1, 11) = "Holiday"
        .TextMatrix(0, 12) = "Medical"
        .TextMatrix(1, 12) = "Leave"
        .TextMatrix(0, 13) = "Salary "
        .TextMatrix(1, 13) = " Days "
        .TextMatrix(1, 14) = "Cat"
        .TextMatrix(1, 15) = "FPCODE"
        .TextMatrix(1, 16) = "Location"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 800
        .ColWidth(13) = 800
        .ColWidth(14) = 800
    End With
End Function

Function filldata()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    Dim rmon As Integer
    If cmb_month.ListIndex <> -1 Then
       rmon = cmb_month.ItemData(cmb_month.ListIndex)
    Else
       rmon = 1
    End If
'----------------------
loc = ""
'----------------------
    
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
        If emptype_chk = 0 Then
           If rmon = 1 Then
              sql = ("Select cast(emp_code as int) as ecode, 0 as el ,* from  emp_mas a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
           Else
              sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
           End If
           emp_cat = "S"
        ElseIf emptype_chk = 1 Then
           If rmon = 1 Then
              sql = ("Select cast(emp_code as int) as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & " and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
           Else
              sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "   and   emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
           End If
           emp_cat = "W"
        ElseIf emptype_chk = 2 Then
           If rmon = 1 Then
               sql = ("Select cast(emp_code as int) as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
           Else
               sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
           End If
           emp_cat = "M"
        End If
    
''        If emptype_chk = 0 Then
''           If rmon = 1 Then
''              sql = ("Select cast(emp_code as int) as ecode, 0 as el ,* from  emp_mas a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & " and s_finyear = " & finyear & "    and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
''           Else
''              sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and s_finyear = " & finyear & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
''           End If
''           emp_cat = "S"
''        ElseIf emptype_chk = 1 Then
''           If rmon = 1 Then
''              sql = ("Select cast(emp_code as int) as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & " and s_finyear = " & finyear & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
''           Else
''              sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and s_finyear = " & finyear & "    and   emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
''           End If
''           emp_cat = "W"
''        ElseIf emptype_chk = 2 Then
''           If rmon = 1 Then
''               sql = ("Select cast(emp_code as int) as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "   and s_finyear = " & finyear & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
''           Else
''               sql = ("Select cast(emp_code as int) as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
''           End If
''           emp_cat = "M"
''        End If
    
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("s_el")
             .TextMatrix(.Rows - 1, 4) = payrs("el")
             If mdays = 28 Then
                .TextMatrix(.Rows - 1, 5) = 24
             ElseIf mdays = 29 Then
                .TextMatrix(.Rows - 1, 5) = 25
             Else
                .TextMatrix(.Rows - 1, 5) = 26
             End If
             .TextMatrix(.Rows - 1, 11) = 0
             If emptype_chk = 2 Then
                If mdays = 28 Then
                   .TextMatrix(.Rows - 1, 6) = 24
                   .TextMatrix(.Rows - 1, 13) = 24
                ElseIf mdays = 29 Then
                   .TextMatrix(.Rows - 1, 6) = 25
                   .TextMatrix(.Rows - 1, 13) = 25
                Else
                   .TextMatrix(.Rows - 1, 6) = 26
                   .TextMatrix(.Rows - 1, 13) = 26
                End If
             End If
             .TextMatrix(.Rows - 1, 14) = payrs("emp_cat")
             .TextMatrix(.Rows - 1, 15) = payrs("emp_fpcode")
             .TextMatrix(.Rows - 1, 16) = payrs("emp_workplace")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
'----------------------
loc = ""
'----------------------
    
''    If emptype_chk = 2 Then Exit Function
    
    
''    If emptype_chk = 0 Then
''       sql = ("Select emp_code as ecode,* from  emp_mas a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' " & loc & "  and EMP_CODE  like '%A'")
''       emp_cat = "S"
''    ElseIf emptype_chk = 1 Then
''       sql = ("Select emp_code as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c where attn_empcode = emp_code and emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A'  and EMP_CODE  like '%A'")
''       emp_cat = "W"
''    End If
    
''        If emptype_chk = 0 Then
''           If rmon = 1 Then
''              sql = ("Select emp_code  as ecode, 0 as el ,* from  emp_mas a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
''           Else
''              sql = ("Select emp_code as ecode,* from  emp_mas a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
''           End If
''           emp_cat = "S"
''        ElseIf emptype_chk = 1 Then
''           If rmon = 1 Then
''              sql = ("Select emp_code as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE")
''           Else
''              sql = ("Select emp_code as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE ")
''           End If
''           emp_cat = "W"
''        ElseIf emptype_chk = 2 Then
''           If rmon = 1 Then
''              sql = ("Select emp_code as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'M' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE")
''           Else
''              sql = ("Select emp_code as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE ")
''           End If
''           emp_cat = "M"
''
''        End If
    
        If emptype_chk = 0 Then
           If rmon = 1 Then
              sql = ("Select emp_code  as ecode, 0 as el ,* from  emp_mas a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
           Else
              sql = ("Select emp_code as ecode,* from  emp_mas a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
           End If
           emp_cat = "S"
        ElseIf emptype_chk = 1 Then
           If rmon = 1 Then
              sql = ("Select emp_code as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE")
           Else
              sql = ("Select emp_code as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE ")
           End If
           emp_cat = "W"
        ElseIf emptype_chk = 2 Then
           If rmon = 1 Then
              sql = ("Select emp_code as ecode, 0 as el,* from  emp_mas a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'M' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE")
           Else
              sql = ("Select emp_code as ecode,* from  emp_mas a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE ")
           End If
           emp_cat = "M"
        
        End If
    
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("s_el")
             .TextMatrix(.Rows - 1, 4) = payrs("el")
             .TextMatrix(.Rows - 1, 5) = 26
             .TextMatrix(.Rows - 1, 13) = 0
             .TextMatrix(.Rows - 1, 14) = payrs("emp_cat")
             .TextMatrix(.Rows - 1, 15) = payrs("emp_fpcode")
             .TextMatrix(.Rows - 1, 16) = payrs("emp_workplace")
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    If cmb_month.ListIndex = -1 Then Exit Function
''uploding data for bio_attendlogs
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          layoffdays = 0
          For i = 2 To att_flex.Rows - 1
            
              If att_flex.TextMatrix(i, 15) = payrs.Fields("a_fpcode") Then

'''modified by Devaraj on 06.01.15
            If emptype_chk = 1 Then
           If mdays = 31 Then
              att_flex.TextMatrix(i, 5) = 31 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
           ElseIf mdays = 30 Then
              att_flex.TextMatrix(i, 5) = 30 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
           ElseIf mdays = 29 Then
              att_flex.TextMatrix(i, 5) = 29 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
           ElseIf mdays = 27 Then
              att_flex.TextMatrix(i, 5) = 27 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
           End If
           
           Else
              att_flex.TextMatrix(i, 5) = "26"
           End If
                 ''att_flex.TextMatrix(i, 5) = "26"
                 If payrs.Fields("a_present") + payrs.Fields("a_el") + payrs.Fields("a_ch") + payrs.Fields("a_absent") + (payrs.Fields("a_pl") + payrs.Fields("a_layoff") + payrs.Fields("a_ml")) > att_flex.TextMatrix(i, 5) Then
                    att_flex.TextMatrix(i, 6) = payrs.Fields("a_present") + payrs.Fields("a_ch") - ((payrs.Fields("a_present") + payrs.Fields("a_ch") + payrs.Fields("a_el") + payrs.Fields("a_absent") + (payrs.Fields("a_pl") + payrs.Fields("a_layoff") + payrs.Fields("a_ml")) - att_flex.TextMatrix(i, 5)))
                 Else
                    att_flex.TextMatrix(i, 6) = payrs.Fields("a_present") + payrs.Fields("a_ch")
''
                 End If
                 If payrs.Fields("a_el") > Val(att_flex.TextMatrix(i, 3) - att_flex.TextMatrix(i, 4)) Then
                    att_flex.TextMatrix(i, 7) = Val(att_flex.TextMatrix(i, 3) - att_flex.TextMatrix(i, 4))
                    att_flex.TextMatrix(i, 8) = payrs.Fields("a_el") - Val(att_flex.TextMatrix(i, 7))
                 Else
                    att_flex.TextMatrix(i, 7) = IIf(payrs.Fields("a_el") > 0, payrs.Fields("a_el"), "")
                End If
''                 att_flex.TextMatrix(i, 8) = IIf(payrs.Fields("a_pl") > 0, payrs.Fields("a_pl"), "")
                 att_flex.TextMatrix(i, 9) = IIf(payrs.Fields("a_absent") > 0, payrs.Fields("a_absent"), "")
                 att_flex.TextMatrix(i, 10) = IIf(payrs.Fields("a_layoff") > 0, payrs.Fields("a_layoff"), "")
                 att_flex.TextMatrix(i, 11) = IIf(payrs.Fields("a_hop") > 0, payrs.Fields("a_hop"), "")
                 att_flex.TextMatrix(i, 12) = IIf(payrs.Fields("a_ml") > 0, payrs.Fields("a_ml"), "")
                 
                 If Val(att_flex.TextMatrix(i, 10)) > 0 Then
                    layoffdays = Val(att_flex.TextMatrix(i, 10)) / 2
                 End If
                 att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + layoffdays
                 
                 
              End If
          Next
        payrs.MoveNext
    Wend
    payrs.Close



End Function



Private Sub cmb_month_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
  '----------------------
loc = ""
'----------------------
    
     find_dates
     load_data
  End If
End Sub

Private Sub cmb_year_Click()
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     find_dates
     load_data
  End If
End Sub

''Private Sub cmd_attend_Click()
''        Dim dsnmdb As String
''        Dim mdbrs As New ADODB.Recordset
''        pst_qry = "delete from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.NewIndex)
''        paydb.Execute pst_qry
'''''---select MSACESS MDB FILE
''''        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\eTimeTrackLite1.mdb"
''
''        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"
''''        mdb_qry = "Select a.EmployeeId,b.employeecode from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# and b.Status = 'Working'  and b.employeecode = '5012'  group by a.EmployeeId,b.employeecode"
''        mdb_qry = "Select a.EmployeeId,b.employeecode from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# and b.Status = 'Working'  group by a.EmployeeId,b.employeecode"
''        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''        While Not mdbrs.EOF
''''            If mdbrs!employeeid = 2643 Then
''''            MsgBox ("DEVARAJ")
''''            End If
''             pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!employeeid & "," & mdbrs!employeecode & ", " & cmb_month.ItemData(cmb_month.NewIndex) & " , " & Val(cmb_year.Text) & " )"
''             paydb.Execute pst_qry
''             mdbrs.MoveNext
''        Wend
''        mdbrs.Close
''        Dim aday As String
''''        mdb_qry = "Select * from attendancelogs as a, employees as b where a.EmployeeId =  b.EmployeeId  and b.employeecode = '1042' and a.Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#  order by a.Attendancedate"
''''        mdb_qry = "Select * from attendancelogs where EmployeeId = 2642 and   Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#  order by Attendancedate"
''''        mdb_qry = "Select * from attendancelogs where EmployeeId = 2643 and  Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#   order by Attendancedate"
''        mdb_qry = "Select * from attendancelogs where Attendancedate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "#   order by Attendancedate"
''        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
''        While Not mdbrs.EOF
''             aday = Trim(Str(Day(mdbrs!Attendancedate)))
''             pst_qry = "update bio_attendlogs set a_day" & aday & " = '" & mdbrs!StatusCode & "',a_in_day" & aday & " = '" & mdbrs!intime & "' ,a_out_day" & aday & " = '" & mdbrs!outtime & "' where a_bioid = " & mdbrs!employeeid
''             paydb.Execute pst_qry
''             mdbrs.MoveNext
''        Wend
''        mdbrs.Close
''        Dim dayfind, dayfind_intime, dayfind_outtime As String
''        Dim present, absent, hop, wop, cl, sl, h, ch, layoff, wo, pl As Single
''        Dim intime, outtime, difftime As Integer
''        Set paydb = New ADODB.Connection
''        Set payrs = New ADODB.Recordset
''        sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.NewIndex)
''        paydb.Open pay
''        payrs.Open sql, paydb, 1, 2
''        If Not payrs.EOF Then
''           While Not payrs.EOF
''                 For i = 1 To 31
''                    dayfind = "a_day" & i
''                    dayfind_intime = "a_in_day" & i
''                    dayfind_outtime = "a_out_day" & i
''                    If IsNull(payrs.Fields(dayfind_intime)) = False And IsNull(payrs.Fields(dayfind_outtime)) = False Then
''                        intime = Hour(payrs.Fields(dayfind_intime)) * 60 + Minute(payrs.Fields(dayfind_intime))
''                        outtime = Hour(payrs.Fields(dayfind_outtime)) * 60 + Minute(payrs.Fields(dayfind_outtime))
''                        difftime = outtime - intime
''                        If payrs.Fields(dayfind) = "P" And difftime > 180 And difftime < 420 Then
''                            payrs.Fields(dayfind) = "½P"
''                            payrs.Update
''                        End If
''                    End If
''                 Next
''                 payrs.MoveNext
''            Wend
''         Else
''            MsgBox ("Details not available for the date ")
''         End If
''         payrs.Close
''''        sql = "select * from bio_attendlogs where a_fpcode = 5012"
''        sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.NewIndex)
''        payrs.Open sql, paydb, 1, 2
''        If Not payrs.EOF Then
''           While Not payrs.EOF
''            present = 0
''            absent = 0
''            hop = 0
''            wop = 0
''            cl = 0
''            sl = 0
''            h = 0
''            ch = 0
''            layoff = 0
''            wo = 0
''            pl = 0
''
''                 For i = 1 To 31
''                    dayfind = "a_day" & i
''                    If payrs.Fields(dayfind) = "P" Or payrs.Fields(dayfind) = "P(OD)" Or payrs.Fields(dayfind) = "½P(OD)" Or payrs.Fields(dayfind) = "A(OD)" Then
''                        present = present + 1
''                    ElseIf payrs.Fields(dayfind) = "A" Then
''                        absent = absent + 1
''                    ElseIf payrs.Fields(dayfind) = "PL" Or payrs.Fields(dayfind) = "PLP" Then
''                        pl = pl + 1
''                    ElseIf payrs.Fields(dayfind) = "½PL" Then
''                        pl = pl + 0.5
''                        absent = absent + 0.5
''                    ElseIf payrs.Fields(dayfind) = "½PLP" Then
''                        pl = pl + 0.5
''                        present = present + 0.5
''                    ElseIf payrs.Fields(dayfind) = "CL" Or payrs.Fields(dayfind) = "CL½P" Or payrs.Fields(dayfind) = "CLP" Then
''                        cl = cl + 1
''                    ElseIf payrs.Fields(dayfind) = "½CL" Then
''                        absent = absent + 0.5
''                        cl = cl + 0.5
''                    ElseIf payrs.Fields(dayfind) = "½CLP" Or payrs.Fields(dayfind) = "½CL½P" Then
''                        present = present + 0.5
''                        cl = cl + 0.5
''                    ElseIf payrs.Fields(dayfind) = "SL" Or payrs.Fields(dayfind) = "SLP" Then
''                        sl = sl + 1
''                    ElseIf payrs.Fields(dayfind) = "½SLP" Then
''                        sl = sl + 0.5
''                        present = present + 0.5
''                    ElseIf payrs.Fields(dayfind) = "H" Then
''                        h = h + 1
''                    ElseIf payrs.Fields(dayfind) = "½P" Then
''                        present = present + 0.5
''                        absent = absent + 0.5
''                    ElseIf payrs.Fields(dayfind) = "Layoff" Or payrs.Fields(dayfind) = "LayoffP" Then
''                        layoff = layoff + 1
''                    ElseIf payrs.Fields(dayfind) = "C.H" Or payrs.Fields(dayfind) = "C.H½P" Or payrs.Fields(dayfind) = "C.HP" Or payrs.Fields(dayfind) = "C.HP(OD)" Then
''                        ch = ch + 1
''                    ElseIf payrs.Fields(dayfind) = "HOP" Or payrs.Fields(dayfind) = "H½P(OD)" Then
''                        hop = hop + 1
''                    ElseIf payrs.Fields(dayfind) = "WOP" Or payrs.Fields(dayfind) = "WOP(OD)" Or payrs.Fields(dayfind) = "WO(OD)" Then
''                        wop = wop + 1
''                    ElseIf payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "WO½P" Then
''                        wo = wo + 1
''                    ElseIf payrs.Fields(dayfind) = "½C.H" Then
''                        ch = ch + 0.5
''                        absent = absent + 0.5
''                    End If
''                 Next
''
''                 payrs("a_present") = present
''                 payrs("a_hop") = hop
''                 payrs("a_wop") = wop
''                 payrs("a_el") = cl
''                 payrs("a_pl") = pl
''                 payrs("a_ml") = sl
''                 payrs("a_holiday") = h
''                 payrs("a_ch") = ch
''                 payrs("a_layoff") = layoff
''                 payrs("a_absent") = absent
''                 payrs("a_wo") = wo
''                 payrs.Update
''
''                 payrs.MoveNext
''            Wend
''         Else
''            MsgBox ("Details not available for the date ")
''         End If
''         payrs.Close
''         MsgBox ("Updated...")
''End Sub

Private Sub cmd_select_Click()
    Dim ff As Integer
    ff = FreeFile 'Sets to next available file number
    With CommonDialog1
        .FileName = ""
        .Filter = "All files (*.xls) |*.xls|" 'Sets the filter
        .ShowOpen
    End With
    txt_filename.Text = CommonDialog1.FileName
End Sub

Private Sub cmd_upload_Click()
     
    If att_flex.Rows = 2 Then
        MsgBox ("Details not selected... ")
        Exit Sub
    End If
    If Trim(txt_filename.Text) = "" Then
       MsgBox ("Select File Name...")
       Exit Sub
    End If
    Dim ecode As String
    Dim excelrows As Integer
    Dim wdays, el, pl, absent, dh, ml, layoffdays As Double
    Set xlwbook = xl.Workbooks.Open("" & txt_filename.Text & "")
    Set xlsheet = xlwbook.Sheets.item(1)
    excelrows = ActiveSheet.UsedRange.Rows.Count
    
''        pst_qry = "select * from emp_mas where emp_company = " & company_code & " and emp_cat = 'S' and emp_no = '" & xlsheet.Cells(i, 1) & "'"
    
    Dim j As Integer
    For j = 2 To excelrows
        ecode = Val(xlsheet.Cells(j, 3))
        wdays = Val(xlsheet.Cells(j, 43)) + Val(xlsheet.Cells(j, 49))
        If wdays > 26 Then
           wdays = 26
        End If
        
        el = Val(xlsheet.Cells(j, 45))
        dh = Val(xlsheet.Cells(j, 47))
      ''  ml = Val(xlsheet.Cells(j, 46))
      ''  pl = Val(xlsheet.Cells(j, 44))
        
        absent = Val(xlsheet.Cells(j, 44))
        For i = 2 To att_flex.Rows - 1
            If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                If att_flex.TextMatrix(i, 15) = ecode Then
                   att_flex.TextMatrix(i, 6) = wdays
                   att_flex.TextMatrix(i, 7) = IIf(el > 0, el, "")
                   att_flex.TextMatrix(i, 8) = IIf(pl > 0, pl, "")
                   att_flex.TextMatrix(i, 9) = IIf(absent > 0, absent, "")
                   att_flex.TextMatrix(i, 11) = IIf(dh > 0, dh, "")
                   att_flex.TextMatrix(i, 12) = IIf(ml > 0, ml, "")
                   layoffdays = 0
                   If Val(att_flex.TextMatrix(i, 10)) > 0 Then
                       layoffdays = Val(att_flex.TextMatrix(i, 10)) / 2
                   End If
                   att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 11)) + layoffdays
                   If Val(att_flex.TextMatrix(i, 6)) > Val(att_flex.TextMatrix(i, 5)) Then
                       MsgBox ("Worked days error ..")
                       att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + Val(att_flex.TextMatrix(i, 12)) + layoffdays
                       att_flex.TextMatrix(i, 6) = ""
''                       Exit Sub
                   End If
                   att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + Val(att_flex.TextMatrix(i, 12)) + layoffdays
                   If Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 8)) + Val(att_flex.TextMatrix(i, 9)) + Val(att_flex.TextMatrix(i, 10)) > Val(att_flex.TextMatrix(i, 5)) Then
                          MsgBox ("Total days not tallied.. check worked days..")
                          att_flex.TextMatrix(i, fin_selcol) = ""
                          layoffdays = 0
                          If Val(att_flex.TextMatrix(i, 10)) > 0 Then
                             layoffdays = Val(att_flex.TextMatrix(i, 10)) / 2
                          End If
                          att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 11)) + Val(att_flex.TextMatrix(i, 12)) + layoffdays
                   End If
                End If
            End If
        Next
    Next
    xl.ActiveWorkbook.Close False, "" & txt_filename.Text & ""
    xl.Quit
    Set xlwbook = Nothing
    Set xl = Nothing
    MsgBox ("Data uploaded from Excel file...")
    att_flex.Enabled = True
End Sub

Private Sub Command1_Click()
  Dim ecode1, ecode2 As String
  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  sql = "select * from old_empmas"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  While Not payrs.EOF
     ecode1 = payrs("emp_code")
     ecode2 = payrs("emp_idcode")
''     sql2 = "update attn_entry set attn_empcode = '" & ecode1 & "'  where attn_company = 1 and attn_finyear = " & finyear & " and attn_empcode = '" & ecode2 & "'  "
''     paydb.Execute sql2
''     sql2 = "update attn_entry set attn_empcode = '" & ecode1 & "'  where attn_company = 2 and attn_finyear = " & finyear & " and attn_empcode = '" & ecode2 & "'  "
''     paydb.Execute sql2
''     sql2 = "update attn_entry set attn_empcode = '" & ecode1 & "'  where attn_company = 3 and attn_finyear = " & finyear & " and attn_empcode = '" & ecode2 & "'  "
''     paydb.Execute sql2
''     sql2 = "update attn_entry set attn_empcode = '" & ecode1 & "'  where attn_company = 5 and attn_finyear = " & finyear & " and attn_empcode = '" & ecode2 & "'  "
''     paydb.Execute sql2
''     sql2 = "update attn_entry set attn_empcode = '" & ecode1 & "'  where attn_company = 8 and attn_finyear = " & finyear & " and attn_empcode = '" & ecode2 & "'  "
''     paydb.Execute sql2
     
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 1 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 2 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 3 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 4 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 5 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     sql2 = "update emp_salary set s_empcode = '" & ecode1 & "'  where s_company = 8 and s_finyear = " & finyear & " and s_empcode = '" & ecode2 & "'  "
     paydb.Execute sql2
     
     
     payrs.MoveNext
  Wend
  MsgBox ("Records are saved")
  payrs.Close
  paydb.Close

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
  filldata
 ''ttn_dt.SetFocus
End Sub

Private Sub edit_Click()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    filldata
    If cmb_month.Text = "" Then
       MsgBox ("Select Month...")
       Exit Sub
    End If
    If cmb_year.Text = "" Then
       MsgBox ("Select Year...")
       Exit Sub
    End If
    
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Then
       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'"
    ElseIf emptype_chk = 1 Then
       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'"
    ElseIf emptype_chk = 2 Then
       sql = "select * from attn_entry a , emp_mas b  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat"
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 2 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
                      If att_flex.TextMatrix(i, 1) = payrs.Fields("attn_empcode") Then
                            att_flex.TextMatrix(i, 5) = payrs.Fields("attn_act_wdays")
                            att_flex.TextMatrix(i, 6) = payrs.Fields("attn_work_days")
                            att_flex.TextMatrix(i, 7) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
                            att_flex.TextMatrix(i, 8) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
                            att_flex.TextMatrix(i, 9) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
                            att_flex.TextMatrix(i, 10) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
                            att_flex.TextMatrix(i, 11) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
                            att_flex.TextMatrix(i, 12) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
                            att_flex.TextMatrix(i, 13) = payrs.Fields("attn_salary_days")
                      End If
                End If
             Next
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
        
    payrs.Close
    sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       save.Enabled = False
    Else
       save.Enabled = True
    End If
    payrs.Close
    att_flex.Enabled = True
End Sub

Public Sub load_data()
    new_entry_chk = 1
    endrow = 0
    fillgrid
    
    i = 2
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  "
''    Else
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'"
''    End If
'----------------------
loc = ""
'----------------------
        
    
    If emptype_chk = 0 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  "
       sql = "select cast(emp_code as int) as ecode,* from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE)"
    ElseIf emptype_chk = 1 Then
''       sql = "select * from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'"
       sql = "select  cast(emp_code as int) as ecode,* from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE not like '%A' " & loc & " order by convert(int, EMP_CODE)"
    ElseIf emptype_chk = 2 Then
''       sql = "select * from attn_entry a , emp_mas b  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat"
       sql = "select * from attn_entry a , emp_mas b , pdept_mas c  where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and (emp_cat = 'M' or (emp_cat in ('S','W') and emp_status  ='B')) and attn_company = emp_company and attn_empcode = emp_code and attn_empcat = emp_cat and emp_dept = dept_code " & loc & "   order by attn_empcode  "
    End If
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
            att_flex.Rows = att_flex.Rows + 1
            att_flex.TextMatrix(att_flex.Rows - 1, 0) = payrs.Fields("dept_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 1) = payrs.Fields("attn_empcode")
            att_flex.TextMatrix(att_flex.Rows - 1, 2) = payrs.Fields("emp_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 5) = payrs.Fields("attn_act_wdays")
            att_flex.TextMatrix(att_flex.Rows - 1, 6) = payrs.Fields("attn_work_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 7) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 8) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 9) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 10) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 11) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 12) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 13) = payrs.Fields("attn_salary_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 14) = payrs.Fields("attn_empcat")
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
 ''   Exit Sub
    
    payrs.Close
    If emptype_chk = 0 Then
       sql = "select * from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE like '%A' "
    ElseIf emptype_chk = 1 Then
       sql = "select * from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE like '%A' "
    ElseIf emptype_chk = 2 Then
       sql = "select * from attn_entry a, emp_mas b , pdept_mas c where attn_company = " & company_code & " and attn_finyear = " & finyear & "  and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'M'  and attn_empcode = emp_code and attn_company = emp_company and emp_dept = dept_code  and EMP_CODE like '%A' "
    
    End If
''  paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
            att_flex.Rows = att_flex.Rows + 1
            att_flex.TextMatrix(att_flex.Rows - 1, 0) = payrs.Fields("dept_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 1) = payrs.Fields("attn_empcode")
            att_flex.TextMatrix(att_flex.Rows - 1, 2) = payrs.Fields("emp_name")
            att_flex.TextMatrix(att_flex.Rows - 1, 5) = payrs.Fields("attn_act_wdays")
            att_flex.TextMatrix(att_flex.Rows - 1, 6) = payrs.Fields("attn_work_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 7) = IIf(payrs.Fields("attn_el") > 0, payrs.Fields("attn_el"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 8) = IIf(payrs.Fields("attn_pl") > 0, payrs.Fields("attn_pl"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 9) = IIf(payrs.Fields("attn_abs") > 0, payrs.Fields("attn_abs"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 10) = IIf(payrs.Fields("attn_layoff") > 0, payrs.Fields("attn_layoff"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 11) = IIf(payrs.Fields("attn_dec_holiday") > 0, payrs.Fields("attn_dec_holiday"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 12) = IIf(payrs.Fields("attn_ml") > 0, payrs.Fields("attn_ml"), "")
            att_flex.TextMatrix(att_flex.Rows - 1, 13) = payrs.Fields("attn_salary_days")
            att_flex.TextMatrix(att_flex.Rows - 1, 14) = payrs.Fields("attn_empcat")
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
''    att_flex.Enabled = False
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
''''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''
''        .Text = "2015"
''    End With
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    new_entry_chk = 0
    If emptype_chk = 0 Then
       millattn_entry.Caption = "Attendacne Entry for STAFF"
       lbl_emp.Caption = "STAFF ATTENDANCE ENTRY"
    ElseIf emptype_chk = 1 Then
       millattn_entry.Caption = "Attendacne Entry for WORKER"
       lbl_emp.Caption = "WORKER ATTENDANCE ENTRY"
    Else
       millattn_entry.Caption = "Attendacne Entry for Retainer / Management"
       lbl_emp.Caption = "RETAINER / MANAGEMENT ATTENDANCE ENTRY"
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
'----------------------
loc = ""
'----------------------
      
  fillgrid
  filldata
  loadchk = 1
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
        Case 5
        If KeyAscii <> 13 Then
            If fin_selcol = 5 Then
                 KeyAscii = attndays_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
            Else
                 KeyAscii = Numeric_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
            End If
        End If
        If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
            att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
        ElseIf KeyAscii = 8 Then
            If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
            KeyAscii = 0
        End If
        layoffdays = 0
        If Val(att_flex.TextMatrix(fin_selrow, 7)) > (Val(att_flex.TextMatrix(fin_selrow, 3)) - Val(att_flex.TextMatrix(fin_selrow, 4))) Then
               MsgBox ("EL is beyond Total Eligible Leave..")
               att_flex.TextMatrix(fin_selrow, 7) = ""
               att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
               Exit Sub
        End If
        If Val(att_flex.TextMatrix(fin_selrow, 10)) > 0 Then
               layoffdays = Val(att_flex.TextMatrix(fin_selrow, 10)) / 2
        End If
        att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + layoffdays
''        ''''''''modified by devaraj
''
''          If Val(att_flex.TextMatrix(fin_selrow, 5)) > 26 Then
''            att_flex.TextMatrix(fin_selrow, 6)=val(att_flex.TextMatrix(fin_selrow ,6) + (val(att_flex.textmatrix
''''        If (Val(att_flex.TextMatrix(fin_selrow, 5)) + Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 11))) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
''                    att_flex.TextMatrix(fin_selrow, 6) = Val(att_flex.TextMatrix(fin_selrow, 5)) - (Val(att_flex.TextMatrix(fin_selrow, 5)) + Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 11))) - Val(att_flex.TextMatrix(fin_selrow, 5))
''                 Else
''                    att_flex.TextMatrix(fin_selrow, 6) = Val(att_flex.TextMatrix(fin_selrow, 5))
''
''                 End If
'''''''''''''''''''''
        If Val(att_flex.TextMatrix(fin_selrow, 6)) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
           MsgBox ("Worked days error ..")
           att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
        ''  .TextMatrix(fin_selrow, 6) = ""
           Exit Sub
        End If
        att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
        If Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 10)) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
               MsgBox ("Total days not tallied.. check worked days..")
               .TextMatrix(fin_selrow, fin_selcol) = ""
               layoffdays = 0
               If Val(att_flex.TextMatrix(fin_selrow, 10)) > 0 Then
                  layoffdays = Val(att_flex.TextMatrix(fin_selrow, 10)) / 2
               End If
               att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
               Exit Sub
        End If
            
        Case 6 To 12
        If att_flex.TextMatrix(fin_selrow, 16) <> "VPT1" Then
            If KeyAscii <> 13 Then
                If fin_selcol = 5 Then
                     KeyAscii = attndays_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
                Else
                     KeyAscii = Numeric_Chk(KeyAscii, att_flex.TextMatrix(fin_selrow, fin_selcol), 5, 2, 2)
                End If
            End If
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
                If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                KeyAscii = 0
            End If
            layoffdays = 0
            If Val(att_flex.TextMatrix(fin_selrow, 7)) > (Val(att_flex.TextMatrix(fin_selrow, 3)) - Val(att_flex.TextMatrix(fin_selrow, 4))) Then
                   MsgBox ("EL is beyond Total Eligible Leave..")
                   att_flex.TextMatrix(fin_selrow, 7) = ""
                   att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
                   Exit Sub
            End If
            If Val(att_flex.TextMatrix(fin_selrow, 10)) > 0 Then
                   layoffdays = Val(att_flex.TextMatrix(fin_selrow, 10)) / 2
            End If
''            att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + layoffdays
            
''            att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
            att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays

            If Val(att_flex.TextMatrix(fin_selrow, 6)) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
               MsgBox ("Worked days error ..")
               att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
              .TextMatrix(fin_selrow, 6) = ""
               Exit Sub
            End If
            att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
            If Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 10)) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
                   MsgBox ("Total days not tallied.. check worked days..")
                   .TextMatrix(fin_selrow, fin_selcol) = ""
                   layoffdays = 0
                   If Val(att_flex.TextMatrix(fin_selrow, 10)) > 0 Then
                      layoffdays = Val(att_flex.TextMatrix(fin_selrow, 10)) / 2
                   End If
                   att_flex.TextMatrix(fin_selrow, 13) = Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 11)) + Val(att_flex.TextMatrix(fin_selrow, 12)) + layoffdays
                   Exit Sub
            End If
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
    lbl_disp.Caption = ""
End Sub
Private Sub SAVE_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
        MsgBox "Out Of Financial Year", vbInformation, "Message"
        Exit Sub
    End If
    
    For i = 2 To att_flex.Rows - 1
        If Val(att_flex.TextMatrix(i, 6)) = 0 And (Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 10)) + Val(att_flex.TextMatrix(i, 11)) + Val(att_flex.TextMatrix(i, 12))) = 0 And Val(att_flex.TextMatrix(i, 13)) > 0 Then
           MsgBox (" Attendance Details are wrong for " & att_flex.TextMatrix(i, 2))
           Exit Sub
        End If
        ''pst_qry = "select * from emp_mas where emp_workplace='MILL'"
        
''        If Val(att_flex.TextMatrix(i, 15)) = 0 and  Then
''           MsgBox (" Finger Print code is missing for " & att_flex.TextMatrix(i, 2))
''           Exit Sub
''        End If
        If mdays = 28 Then
            If Val(att_flex.TextMatrix(i, 5)) < 24 Then
               MsgBox (" Actual working days wrong for " & att_flex.TextMatrix(i, 2))
               Exit Sub
            End If
        Else
            If Val(att_flex.TextMatrix(i, 5)) < 25 Then
               MsgBox (" Actual working days wrong for " & att_flex.TextMatrix(i, 2))
               Exit Sub
            End If
        End If
        If Val(att_flex.TextMatrix(i, 4)) + Val(att_flex.TextMatrix(i, 7)) > Val(att_flex.TextMatrix(i, 3)) Then
           MsgBox (" Eligible Leave is beyond the Total Eligible leave for " & att_flex.TextMatrix(i, 2))
           Exit Sub
        End If
        
    Next
On Error GoTo err_handler
  If att_flex.Rows < 3 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
  Me.MousePointer = 11

  Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  paydb.Open pay
  paydb.BeginTrans
  If new_entry_chk = 1 Then
     If emptype_chk = 0 Then
        sql = "delete from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'S' and attn_empcode in (select emp_code from emp_mas  where emp_company = " & company_code & "  " & loc & ")"
        paydb.Execute sql
     ElseIf emptype_chk = 1 Then
        sql = "delete from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = 'W' and attn_empcode in (select emp_code from emp_mas  where emp_company = " & company_code & "  " & loc & ")"
        paydb.Execute sql
     ElseIf emptype_chk = 2 Then
        For i = 2 To att_flex.Rows - 1
            If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
               sql = "delete from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and attn_year = " & Val(cmb_year.Text) & " and attn_empcat = '" & att_flex.TextMatrix(i, 14) & "' and attn_empcode =  '" & att_flex.TextMatrix(i, 1) & "' "
               paydb.Execute sql
            End If
        Next
     End If
  End If
  sql = "select * from attn_entry where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 2 To att_flex.Rows - 1

      If Trim(att_flex.TextMatrix(i, 1)) <> "" Then
            payrs.AddNew
            payrs.Fields("attn_company") = company_code
            payrs.Fields("attn_finyear") = finyear
            payrs.Fields("attn_month") = cmb_month.ItemData(cmb_month.ListIndex)
            payrs.Fields("attn_year") = Val(cmb_year.Text)
            payrs.Fields("attn_empcode") = att_flex.TextMatrix(i, 1)
''            find_empdetails (att_flex.TextMatrix(i, 1))
''            payrs.Fields("attn_empcat") = emp_cat
            payrs.Fields("attn_empcat") = att_flex.TextMatrix(i, 14)
            payrs.Fields("attn_act_wdays") = Val(att_flex.TextMatrix(i, 5))
            payrs.Fields("attn_work_days") = Val(att_flex.TextMatrix(i, 6))
            payrs.Fields("attn_el") = Val(att_flex.TextMatrix(i, 7))
            payrs.Fields("attn_pl") = Val(att_flex.TextMatrix(i, 8))
            payrs.Fields("attn_abs") = Val(att_flex.TextMatrix(i, 9))
            payrs.Fields("attn_layoff") = Val(att_flex.TextMatrix(i, 10))
            payrs.Fields("attn_dec_holiday") = Val(att_flex.TextMatrix(i, 11))
            payrs.Fields("attn_ml") = Val(att_flex.TextMatrix(i, 12))
            payrs.Fields("attn_salary_days") = Val(att_flex.TextMatrix(i, 13))
            payrs.Update
      End If
  Next
  MsgBox ("Records are saved")
  paydb.CommitTrans
  payrs.Close
  paydb.Close
  fillgrid
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
Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
    Select Case att_flex.Col
        Case 3
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
        Case 4, 1, 2
'            txt.Visible = False
'            lst_code.Visible = False
''        Case 5
''            If (Val(att_flex.TextMatrix(fin_selrow, 5)) + Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 11))) > Val(att_flex.TextMatrix(fin_selrow, 5)) Then
''                    att_flex.TextMatrix(fin_selrow, 6) = Val(att_flex.TextMatrix(fin_selrow, 5)) - (Val(att_flex.TextMatrix(fin_selrow, 5)) + Val(att_flex.TextMatrix(fin_selrow, 6)) + Val(att_flex.TextMatrix(fin_selrow, 7)) + Val(att_flex.TextMatrix(fin_selrow, 8)) + Val(att_flex.TextMatrix(fin_selrow, 9)) + Val(att_flex.TextMatrix(fin_selrow, 11))) - Val(att_flex.TextMatrix(fin_selrow, 5))
''                 Else
''                    att_flex.TextMatrix(fin_selrow, 6) = Val(att_flex.TextMatrix(fin_selrow, 5))
''''
''                 End If
    End Select
    End With
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
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    Dim rmon As Integer
    If cmb_month.ListIndex <> -1 Then
       rmon = cmb_month.ItemData(cmb_month.ListIndex)
    Else
       rmon = 1
    End If
'----------------------
loc = ""
'----------------------
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
        If emptype_chk = 0 Then
           If rmon = 1 Then
              sql = ("Select cast(emp_code as int) as ecode, 0 as el ,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
           Else
              sql = ("Select cast(emp_code as int) as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE not like '%A' " & loc & "  order by convert(int, EMP_CODE) ")
           End If
           emp_cat = "S"
        ElseIf emptype_chk = 1 Then
           If rmon = 1 Then
              sql = ("Select cast(emp_code as int) as ecode,dateadd(year,58,emp_dob) as emp_retire , 0 as el,* from  emp_voupay_mast a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & " and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
           Else
              sql = ("Select cast(emp_code as int) as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "   and   emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE not like '%A'  order by convert(int, EMP_CODE)")
           End If
           emp_cat = "W"
        ElseIf emptype_chk = 2 Then
           If rmon = 1 Then
               sql = ("Select cast(emp_code as int) as ecode,dateadd(year,58,emp_dob) as emp_retire , 0 as el,* from  emp_voupay_mast a,emp_eligible_leave b where emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
           Else
               sql = ("Select cast(emp_code as int) as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and emp_company = '" & company_code & "' and ((emp_cat in ('S','W') and emp_status = 'B') or emp_cat in ('M')) order by convert(int, EMP_CODE)")
           End If
           emp_cat = "M"
        End If
    
    
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("s_el")
             .TextMatrix(.Rows - 1, 4) = payrs("el")
             If mdays = 28 Then
                .TextMatrix(.Rows - 1, 5) = 24
             ElseIf mdays = 29 Then
                .TextMatrix(.Rows - 1, 5) = 25
             Else
                .TextMatrix(.Rows - 1, 5) = 26
             End If
             .TextMatrix(.Rows - 1, 11) = 0
             If emptype_chk = 2 Then
                If mdays = 28 Then
                   .TextMatrix(.Rows - 1, 6) = 24
                   .TextMatrix(.Rows - 1, 13) = 24
                ElseIf mdays = 29 Then
                   .TextMatrix(.Rows - 1, 6) = 25
                   .TextMatrix(.Rows - 1, 13) = 25
                Else
                   .TextMatrix(.Rows - 1, 6) = 26
                   .TextMatrix(.Rows - 1, 13) = 26
                End If
             End If
             .TextMatrix(.Rows - 1, 14) = payrs("emp_cat")
             .TextMatrix(.Rows - 1, 15) = payrs("emp_fpcode")
             .TextMatrix(.Rows - 1, 16) = payrs("emp_workplace")
             If Format(payrs.Fields("emp_retire"), "yyyy/MM/dd") >= Format(st_date.Value, "yyyy/MM/dd") And Format(payrs.Fields("emp_retire"), "yyyy/MM/dd") <= Format(end_date.Value, "yyyy/MM/dd") Then
                 .TextMatrix(.Rows - 1, 17) = Day(payrs("emp_retire"))
             End If
             endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    payrs.Close
'----------------------
loc = ""
'----------------------
    
        If emptype_chk = 0 Then
           If rmon = 1 Then
              sql = ("Select emp_code  as ecode, 0 as el ,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a ,emp_eligible_leave b where emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
           Else
              sql = ("Select emp_code as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a ,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & " and attn_company = '" & company_code & "'  and attn_empcat = 'S' and attn_month < " & rmon & "  group by attn_empcode) c  where attn_empcode = emp_code and emp_company = s_company  and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'S' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE like '%A' " & loc & "  order by EMP_CODE ")
           End If
           emp_cat = "S"
        ElseIf emptype_chk = 1 Then
           If rmon = 1 Then
              sql = ("Select emp_code as ecode,dateadd(year,58,emp_dob) as emp_retire , 0 as el,* from  emp_voupay_mast a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE like '%A'  order by EMP_CODE")
           Else
              sql = ("Select emp_code as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'W' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and (emp_status = 'A' or (emp_status = 'R' and  emp_resigneddate  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "')) and EMP_CODE like '%A'  order by EMP_CODE ")
           End If
           emp_cat = "W"
        ElseIf emptype_chk = 2 Then
           If rmon = 1 Then
              sql = ("Select emp_code as ecode,dateadd(year,58,emp_dob) as emp_retire , 0 as el,* from  emp_voupay_mast a,emp_eligible_leave b where  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'M' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE")
           Else
              sql = ("Select emp_code as ecode,dateadd(year,58,emp_dob) as emp_retire ,* from  emp_voupay_mast a,emp_eligible_leave b ,(select attn_empcode , sum(attn_el) as el from attn_entry where attn_year = " & Val(cmb_year) & "  and attn_company = '" & company_code & "'  and attn_empcat = 'M' and attn_month < " & rmon & "  group by attn_empcode) c   where attn_empcode = emp_code and  emp_company = s_company and emp_code = s_empcode and s_year = " & Val(cmb_year) & "  and  emp_company = '" & company_code & "' and emp_cat = 'W' and (emp_type = 0 or emp_type = 1 or emp_type = 2 or emp_type = 3) and emp_status = 'A' and EMP_CODE like '%A'  order by EMP_CODE ")
           End If
           emp_cat = "M"
        
        End If
    
    
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        With att_flex
             .Rows = .Rows + 1
              find_deptname (payrs.Fields("emp_dept"))
             .TextMatrix(.Rows - 1, 0) = dname
             .TextMatrix(.Rows - 1, 1) = payrs("ecode")
             .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
             .TextMatrix(.Rows - 1, 3) = payrs("s_el")
             .TextMatrix(.Rows - 1, 4) = payrs("el")
             .TextMatrix(.Rows - 1, 5) = 26
             .TextMatrix(.Rows - 1, 13) = 0
             .TextMatrix(.Rows - 1, 14) = payrs("emp_cat")
             .TextMatrix(.Rows - 1, 15) = payrs("emp_fpcode")
             .TextMatrix(.Rows - 1, 16) = payrs("emp_workplace")
             If Format(payrs.Fields("emp_retire"), "yyyy/MM/dd") >= Format(st_date.Value, "yyyy/MM/dd") And Format(payrs.Fields("emp_retire"), "yyyy/MM/dd") <= Format(end_date.Value, "yyyy/MM/dd") Then
                .TextMatrix(.Rows - 1, 17) = Day(payrs("emp_retire"))
             End If
             
             endrow = endrow + 1
        
        End With
        payrs.MoveNext
    Wend
    payrs.Close
    Dim wdays As Single
    wdays = 0
    If cmb_month.ListIndex = -1 Then Exit Function
''uploding data for bio_attendlogs
    sql = "select * from bio_attendlogs where a_fpcode <> 0 and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
''          If payrs.Fields("a_fpcode") = 2014 Then
''             MsgBox ("wait")
''          End If
          
          layoffdays = 0
          For i = 2 To att_flex.Rows - 1
              If att_flex.TextMatrix(i, 15) = payrs.Fields("a_fpcode") Then
              
  '''modified by Devaraj on 06.01.15
                If emptype_chk = 1 Then
                    If mdays = 31 Then
                       att_flex.TextMatrix(i, 5) = 31 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
                    ElseIf mdays = 30 Then
                       att_flex.TextMatrix(i, 5) = 30 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
                    ElseIf mdays = 29 Then
                       att_flex.TextMatrix(i, 5) = 29 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
                    ElseIf mdays = 27 Then
                       att_flex.TextMatrix(i, 5) = 27 - (payrs.Fields("a_wo") + payrs.Fields("a_wop"))
                    End If
                    
                    
                    If Val(att_flex.TextMatrix(i, 5)) > 26 And Val(att_flex.TextMatrix(i, 5)) <= 26.5 Then
                       att_flex.TextMatrix(i, 5) = 26
                    End If
                    
                    
                    If Val(att_flex.TextMatrix(i, 5)) > 27 Then
                       att_flex.TextMatrix(i, 5) = 27
                    End If
                Else
                   att_flex.TextMatrix(i, 5) = "26"
                End If
                If payrs.Fields("a_present") + payrs.Fields("a_el") + payrs.Fields("a_ch") + payrs.Fields("a_absent") + (payrs.Fields("a_pl") + payrs.Fields("a_layoff") + payrs.Fields("a_ml")) > att_flex.TextMatrix(i, 5) Then
                   att_flex.TextMatrix(i, 6) = payrs.Fields("a_present") + payrs.Fields("a_ch") - ((payrs.Fields("a_present") + payrs.Fields("a_ch") + payrs.Fields("a_el") + payrs.Fields("a_absent") + (payrs.Fields("a_pl") + payrs.Fields("a_layoff") + payrs.Fields("a_ml")) - att_flex.TextMatrix(i, 5)))
                Else
                   att_flex.TextMatrix(i, 6) = payrs.Fields("a_present") + payrs.Fields("a_ch")
                End If
                If (Val(att_flex.TextMatrix(i, 6)) > Val(att_flex.TextMatrix(i, 17))) And Val(att_flex.TextMatrix(i, 17)) > 0 Then
                      att_flex.TextMatrix(i, 6) = Val(att_flex.TextMatrix(i, 17))
                End If
                If payrs.Fields("a_el") > Val(att_flex.TextMatrix(i, 3) - att_flex.TextMatrix(i, 4)) Then
                      att_flex.TextMatrix(i, 7) = Val(att_flex.TextMatrix(i, 3) - att_flex.TextMatrix(i, 4))
                      att_flex.TextMatrix(i, 8) = payrs.Fields("a_el") - Val(att_flex.TextMatrix(i, 7))
                Else
                      att_flex.TextMatrix(i, 7) = IIf(payrs.Fields("a_el") > 0, payrs.Fields("a_el"), "")
                End If
                
                att_flex.TextMatrix(i, 8) = Val(att_flex.TextMatrix(i, 8)) + payrs.Fields("a_pl")
                att_flex.TextMatrix(i, 9) = IIf(payrs.Fields("a_absent") > 0, payrs.Fields("a_absent"), "")
                att_flex.TextMatrix(i, 10) = IIf(payrs.Fields("a_layoff") > 0, payrs.Fields("a_layoff"), "")
                att_flex.TextMatrix(i, 11) = IIf(payrs.Fields("a_holiday") > 0, payrs.Fields("a_holiday"), "")
                att_flex.TextMatrix(i, 12) = IIf(payrs.Fields("a_ml") > 0, payrs.Fields("a_ml"), "")
                att_flex.TextMatrix(i, 18) = IIf(payrs.Fields("a_hpe") > 0, payrs.Fields("a_hpe"), "")
                If Val(att_flex.TextMatrix(i, 10)) > 0 Then
                   layoffdays = Val(att_flex.TextMatrix(i, 10)) / 2
                End If
                If emptype_chk = 1 Then
                    att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + layoffdays + Val(att_flex.TextMatrix(i, 18))
                Else
                   wdays = 26 - (Val(att_flex.TextMatrix(i, 8)) + Val(att_flex.TextMatrix(i, 9)) + Val(att_flex.TextMatrix(i, 12)))
                   att_flex.TextMatrix(i, 13) = wdays + Val(att_flex.TextMatrix(i, 18))
''                   If Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + layoffdays + Val(att_flex.TextMatrix(i, 18)) > 26 Then
''                      att_flex.TextMatrix(i, 13) = 26 - Val(att_flex.TextMatrix(i, 9))
''                   Else
''                      att_flex.TextMatrix(i, 13) = Val(att_flex.TextMatrix(i, 6)) + Val(att_flex.TextMatrix(i, 7)) + Val(att_flex.TextMatrix(i, 11)) + layoffdays + Val(att_flex.TextMatrix(i, 18))
''                   End If
                End If
              End If
          Next
        payrs.MoveNext
    Wend
    payrs.Close
End Function


