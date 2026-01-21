VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form month_deduct 
   Caption         =   "DEDUCTION ENTRY "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122814465
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122814465
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2280
      TabIndex        =   19
      Top             =   6720
      Width           =   2055
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   825
         Left            =   1080
         Picture         =   "month_deduct.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   825
         Left            =   120
         Picture         =   "month_deduct.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE DEDUCTION DETAILS"
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
      Height          =   6585
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   10755
      Begin MSFlexGridLib.MSFlexGrid deduct_flex 
         Height          =   2970
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   5239
         _Version        =   393216
         Cols            =   3
         FixedCols       =   2
         BackColor       =   16776960
         BackColorFixed  =   -2147483624
         BackColorBkg    =   12632256
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
      Begin VB.Frame Frame2 
         Height          =   2370
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   10305
         Begin VB.TextBox Avilable_working_days 
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
            Left            =   2760
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1800
            Width           =   660
         End
         Begin VB.TextBox desi 
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
            Height          =   360
            Left            =   8115
            TabIndex        =   8
            Top             =   1335
            Width           =   2040
         End
         Begin VB.TextBox dept 
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
            Height          =   390
            Left            =   2820
            TabIndex        =   7
            Top             =   1290
            Width           =   3540
         End
         Begin VB.TextBox emp_idcode 
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
            Height          =   450
            Left            =   8100
            TabIndex        =   6
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox emptype 
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
            Height          =   390
            Left            =   2820
            TabIndex        =   5
            Top             =   750
            Width           =   3525
         End
         Begin VB.ComboBox empname_cmb 
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
            Left            =   2835
            TabIndex        =   4
            Top             =   285
            Width           =   7170
         End
         Begin VB.Label Label8 
            Caption         =   "Available Working Days"
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
            Height          =   285
            Left            =   225
            TabIndex        =   18
            Top             =   1920
            Width           =   2340
         End
         Begin VB.Label Label7 
            Caption         =   "Designation"
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
            Height          =   270
            Left            =   6540
            TabIndex        =   17
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Department"
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
            Height          =   390
            Left            =   225
            TabIndex        =   16
            Top             =   1335
            Width           =   2370
         End
         Begin VB.Label Label5 
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
            Height          =   360
            Left            =   6540
            TabIndex        =   15
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label Label4 
            Caption         =   "Employee Type"
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
            Height          =   270
            Left            =   225
            TabIndex        =   14
            Top             =   810
            Width           =   2145
         End
         Begin VB.Label Label3 
            Caption         =   "Employee Name"
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
            Height          =   300
            Left            =   225
            TabIndex        =   13
            Top             =   360
            Width           =   2280
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
         Left            =   3060
         TabIndex        =   1
         Top             =   360
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
         Left            =   7380
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10755
         Y1              =   840
         Y2              =   840
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
         Left            =   6330
         TabIndex        =   12
         Top             =   480
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
         Left            =   1140
         TabIndex        =   11
         Top             =   360
         Width           =   1710
      End
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
      Left            =   5040
      TabIndex        =   22
      Top             =   7200
      Width           =   4815
   End
End
Attribute VB_Name = "month_deduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim emp_chk As Integer
Dim emptypecode As String
Dim blank_rec_upd As Integer
Public nextname As String

Private Sub Command1_Click()
    fillgrid
End Sub

Private Sub cmb_month_Click()
  find_dates
  lbl_disp.Caption = ""
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     process_data
  End If
End Sub

Private Sub cmb_year_Click()
    If Trim(cmb_month.Text) = "" Then
       MsgBox ("Select Deduction month")
       Exit Sub
    End If
  find_dates
  lbl_disp.Caption = ""
  If Trim(cmb_month.Text) <> "" And Trim(cmb_year.Text) <> "" Then
     process_data
  End If
End Sub

Private Sub empname_cmb_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
        MsgBox "Out Of Financial Year", vbInformation, "Message"
        Exit Sub
    End If
        
    If Trim(cmb_month.Text) = "" Then
       MsgBox ("Select Deduction month")
       Exit Sub
    End If
    '' Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
''    paydb.Open pay
    If emptype_chk = 0 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('S','M') ")
    ElseIf emptype_chk = 1 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat = 'W' ")
    ElseIf emptype_chk = 2 Then
       sql = ("select * from emp_mas where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('M') ")
    ElseIf emptype_chk = 3 Then
       sql = ("select * from emp_voupay_mast where emp_name = '" & empname_cmb.Text & "' and  emp_company = '" & company_code & "' and emp_cat in ('R') ")
    
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
    Else
       emp_idcode = payrs.Fields("emp_code")
       find_deptname (payrs.Fields("emp_dept"))
       dept.Text = dname
       find_desiname (payrs.Fields("emp_design"))
       DESI.Text = dname
''       find_etypename (payrs.Fields("emp_type"))
''       emptype.Text = dname'
       emptypecode = payrs.Fields("emp_type")
       If payrs.Fields("emp_type") = "S" Then
          emptype.Text = "STAFF"
       Else
          emptype.Text = "WORKER"
       End If
       
    End If
    endrow = 0
    fillgrid
    '' Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Or emptype_chk = 3 Then
       sql = ("Select * from  pdedu_mas where pdedu_type in (1,2,4) order by pdedu_code")
    Else
       sql = ("Select * from  pdedu_mas where pdedu_type in (1,3,4) order by pdedu_code")
    End If
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        With deduct_flex
             .Rows = .Rows + 1
             .TextMatrix(.Rows - 1, 0) = payrs(0)
             .TextMatrix(.Rows - 1, 1) = payrs(1)
              endrow = endrow + 1
        End With
        payrs.MoveNext
    Wend
    '' Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    If emptype_chk = 0 Then
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat in ('S','M')")
    ElseIf emptype_chk = 3 Then
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat in ('R')")
    Else
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat = 'W'")
    End If
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
       Avilable_working_days = payrs.Fields("e_avail_workdays")
       With deduct_flex
            For i = 1 To endrow
                If .TextMatrix(i, 0) = payrs.Fields("e_ded_code") Then .TextMatrix(i, 2) = payrs.Fields("e_ded_amount")
            Next
            payrs.MoveNext
       End With
    Wend
    payrs.Close
End Sub

Private Sub exit_Click()
   Unload Me
End Sub
Private Sub Form_Load()
    If emptype_chk = 0 Then
       month_deduct.Caption = "Deduction Entry for STAFF"
    Else
       month_deduct.Caption = "Deduction Entry for WORLER"
    End If
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
''    End With
''    cmb_year.Text = "2015"
    
    With cmb_year
        .AddItem Left(fyear, 4)
        .AddItem Mid(fyear, 6, 4)
    End With
    
    '' Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset

'----------------------
loc = ""
'----------------------
    

    If emptype_chk = 0 Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat in ('S','M') and (emp_status = 'A' or emp_status = 'B')  " & loc & " order by emp_name")
       Avilable_working_days.Enabled = False
    ElseIf emptype_chk = 1 Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat = 'W' and emp_status = 'A' order by emp_name")
    ElseIf emptype_chk = 2 Then
       sql = ("Select * from  emp_mas where emp_company = '" & company_code & "' and emp_cat in ('M') and (emp_status = 'A' or emp_status = 'B') order by emp_name")
       Avilable_working_days.Enabled = False
    ElseIf emptype_chk = 3 Then
       sql = ("Select * from  emp_voupay_mast where emp_company = '" & company_code & "' and emp_cat = 'R' and emp_status = 'A' order by emp_name")
    End If
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
''    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    ''payrs.MoveFirst
    emp_chk = 0
    While Not payrs.EOF
        empname_cmb.AddItem payrs("emp_name")
    ''    empname_cmb.ItemData(empname_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
        emp_chk = emp_chk + 1
    Wend
    Avilable_working_days = 26
    blank_rec_upd = 0
    fillgrid
End Sub
Function fillgrid()
    With deduct_flex
        .Clear
        .Rows = 1
        .TextMatrix(0, 0) = "Deduct_Code"
        .TextMatrix(0, 1) = "Deduction Name"
        .TextMatrix(0, 2) = "Deduction Amount"
        .ColWidth(0) = 2000
        .ColWidth(1) = 6000
        .ColWidth(2) = 2000
     End With
End Function

Private Sub deduct_flex_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
 Dim fin_selrow%, fin_selcol%
 fin_selrow = deduct_flex.Row
 fin_selcol = deduct_flex.Col
 With deduct_flex
 Select Case fin_selcol
        Case 2
            If KeyAscii <> 13 Then
               KeyAscii = Numeric_Chk(KeyAscii, deduct_flex.TextMatrix(fin_selrow, fin_selcol), 7, 5, 2)
            End If
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                deduct_flex.TextMatrix(fin_selrow, fin_selcol) = deduct_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
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

Private Sub SAVE_Click()
    If st_date.Value < gdt_finsdate Or end_date.Value > gdt_finedate Then
        MsgBox "Out Of Financial Year", vbInformation, "Message"
        Exit Sub
    End If
      
  If endrow = 0 Then
     MsgBox (" Details not available ")
     Exit Sub
  End If
On Error GoTo err_handler
  '' Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
''  paydb.Open pay
  paydb.BeginTrans
'----------------------
loc = ""
'----------------------
  
  If emptype_chk = 0 Then
       sql = ("delete from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat ='S' ")
  ElseIf emptype_chk = 3 Then
       sql = ("delete from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat = 'R'  ")
  Else
       sql = ("delete from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat = 'W'  ")
  End If
  paydb.Execute sql
  
  If emptype_chk = 0 Then
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat ='S' ")
  ElseIf emptype_chk = 3 Then
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat ='R' ")
  Else
       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
            & " and e_emp_code = '" & emp_idcode.Text & "' and e_company = '" & company_code & "' and e_emp_cat = 'W'  ")
  End If
''  paydb.Open pay
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''  While Not payrs.EOF
''     payrs.Delete
''      payrs.Update
''     payrs.MoveNext
''  Wend
  ''payrs.Update
  blank_rec_upd = 1
  For i = 1 To endrow
      If Trim(deduct_flex.TextMatrix(i, 2)) <> "" Then
            payrs.AddNew
            payrs.Fields("e_company") = company_code
            payrs.Fields("e_finyear") = finyear
            payrs.Fields("e_emp_code") = emp_idcode.Text
            payrs.Fields("e_emp_cat") = IIf(emptype_chk = 0, "S", "W")
            If emptype_chk = 0 Then
               payrs.Fields("e_emp_cat") = "S"
            ElseIf emptype_chk = 1 Then
               payrs.Fields("e_emp_cat") = "W"
            ElseIf emptype_chk = 2 Then
               payrs.Fields("e_emp_cat") = "M"
            ElseIf emptype_chk = 3 Then
               payrs.Fields("e_emp_cat") = "R"
            
            End If
            
            payrs.Fields("e_ded_year") = Val(cmb_year)
            payrs.Fields("e_ded_month") = cmb_month.ItemData(cmb_month.ListIndex)
            payrs.Fields("e_ded_code") = Val(deduct_flex.TextMatrix(i, 0))
            payrs.Fields("e_ded_amount") = Val(deduct_flex.TextMatrix(i, 2))
            payrs.Fields("e_avail_workdays") = Val(Avilable_working_days.Text)
            payrs.Update
            blank_rec_upd = 0
      End If
  Next
  If blank_rec_upd = 1 Then
            payrs.AddNew
            payrs.Fields("e_company") = company_code
            payrs.Fields("e_finyear") = finyear
            payrs.Fields("e_emp_code") = emp_idcode.Text
''            payrs.Fields("e_emp_cat") = IIf(emptype_chk = 0, "S", "W")
            If emptype_chk = 0 Then
               payrs.Fields("e_emp_cat") = "S"
            ElseIf emptype_chk = 1 Then
               payrs.Fields("e_emp_cat") = "W"
            ElseIf emptype_chk = 2 Then
               payrs.Fields("e_emp_cat") = "M"
            ElseIf emptype_chk = 3 Then
               payrs.Fields("e_emp_cat") = "R"
               
            End If
            
            payrs.Fields("e_ded_year") = Val(cmb_year)
            payrs.Fields("e_ded_month") = cmb_month.ItemData(cmb_month.ListIndex)
            payrs.Fields("e_ded_code") = 1
            payrs.Fields("e_ded_amount") = 0#
            payrs.Fields("e_avail_workdays") = Val(Avilable_working_days.Text)
            payrs.Update
  End If
  MsgBox ("Records are saved")
  paydb.CommitTrans
  emptype.Text = ""
  emp_idcode.Text = ""
  dept.Text = ""
  DESI.Text = ""
  fillgrid
''  If Val(empname_cmb.ListCount) - 1 >= Val(empname_cmb.ListIndex) + 1 Then
''        nextname = Val(empname_cmb.ItemData(empname_cmb.ListIndex + 1))
''        '' Set paydb = New ADODB.Connection
''        Set payrs = New ADODB.Recordset
''        paydb.Open pay
''        sql = ("select * from emp_mas where emp_idcode = " & nextname & " and emp_company = '" & company_code & "'")
''        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''        If payrs.EOF Then
''           MsgBox ("Record over")
''        Else
''            empname_cmb.Text = payrs.Fields("emp_name")
''        End If
''        emp_idcode = payrs.Fields("emp_idcode")
''        find_deptname (payrs.Fields("emp_dept"))
''        dept.Text = dname
''        find_desiname (payrs.Fields("emp_design"))
''        DESI.Text = dname
''        find_etypename (payrs.Fields("emp_type"))
''        emptype.Text = dname
''''        endrow = 0
''''        fill_deduction
''  Else
''        MsgBox ("Data over")
''  End If
  Exit Sub
err_handler:
    paydb.RollbackTrans
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
  

End Sub


''Public Function fill_deduction()
''    fillgrid
''    '' Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    sql = ("Select * from  pdedu_mas order by pdedu_code")
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    payrs.MoveFirst
''    While Not payrs.EOF
''        With deduct_flex
''             .Rows = .Rows + 1
''             .TextMatrix(.Rows - 1, 0) = payrs(0)
''             .TextMatrix(.Rows - 1, 1) = payrs(1)
''              endrow = endrow + 1
''        End With
''        payrs.MoveNext
''    Wend
''    '' Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    If emptype_chk = 0 Then
''       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
''            & " and e_emp_code = " & nextname & " and e_company = '" & company_code & "' and (e_emp_type = 0 or e_emp_type = 1)")
''    Else
''       sql = ("Select * from  monthly_deduction where e_ded_year = " & Trim(cmb_year.Text) & " and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) _
''            & " and e_emp_code = " & nextname & " and e_company = '" & company_code & "' and (e_emp_type = 2 or e_emp_type = 3)")
''    End If
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''       Avilable_working_days = payrs.Fields("e_avail_workdays")
''       With deduct_flex
''            For i = 1 To endrow
''                If .TextMatrix(i, 0) = payrs.Fields("e_ded_code") Then .TextMatrix(i, 2) = payrs.Fields("e_ded_amount")
''            Next
''            payrs.MoveNext
''       End With
''    Wend
''
''End Function
Public Function process_data()
  Dim sql2 As String
  '' Set paydb = New ADODB.Connection
  Set payrs = New ADODB.Recordset
  ''paydb.Open pay
  sql = "select *  from attn_entry where attn_company = " & company_code & " and attn_finyear = " & finyear & " and attn_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and attn_year = " & Trim(cmb_year.Text) & " and attn_empcode not in (select e_emp_code  from monthly_deduction where e_company = " & company_code & " and e_finyear = " & finyear & "  and e_ded_month = " & cmb_month.ItemData(cmb_month.ListIndex) & "  and e_ded_year = " & Trim(cmb_year.Text) & ")"
  paydb.Execute sql
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  While Not payrs.EOF
        sql2 = "insert into monthly_deduction values ( " & payrs("attn_company") & " , " & payrs("attn_finyear") & " , '" & payrs("attn_empcode") & "' , '" & payrs("attn_empcat") & "' , " & payrs("attn_year") & " , " & payrs("attn_month") & " , 1, 0 , " & payrs("attn_act_wdays") & " )"
        paydb.Execute sql2
     payrs.MoveNext
  Wend
  payrs.Close
  
  sql = "select * from payroll_lock where pay_company = " & company_code & " and pay_finyear = " & finyear & " and pay_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and pay_year = " & Val(cmb_year.Text) & ""
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  If Not payrs.EOF Then
       If payrs("pay_dedu_lock") = "Y" Then
          save.Enabled = False
          lbl_disp.Caption = "Deduction Locked .. Can't Modify"
       End If
  Else
       lbl_disp.Caption = ""
       save.Enabled = True
  End If
  payrs.Close
End Function

Public Sub find_dates()
    Dim mdays, diff As Integer
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If cmb_year.Text = "" Then Exit Sub
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

