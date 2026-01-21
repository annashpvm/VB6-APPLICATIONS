VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_shift_schdule 
   Caption         =   "Weekly - Shift Schedule "
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   16245
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4680
      TabIndex        =   29
      Top             =   6840
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   240
      TabIndex        =   22
      Top             =   6960
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109248513
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   24
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109248513
         CurrentDate     =   39359
      End
      Begin VB.Label Label10 
         Caption         =   " From Date"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1455
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
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   6495
      Begin MSComCtl2.MonthView Calendar1 
         Height          =   2370
         Left            =   1680
         TabIndex        =   33
         Top             =   1080
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   109248513
         CurrentDate     =   43048
      End
      Begin VB.CommandButton cmd_go 
         Caption         =   "GO"
         Height          =   375
         Left            =   4920
         TabIndex        =   32
         Top             =   4080
         Width           =   975
      End
      Begin VB.ComboBox cmb_shift 
         Height          =   315
         Left            =   3000
         TabIndex        =   28
         Top             =   4080
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   1455
         Left            =   360
         TabIndex        =   21
         Top             =   4560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2566
         _Version        =   393216
      End
      Begin VB.ComboBox cmb_weekoff 
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
         Height          =   360
         ItemData        =   "frm_shift_schdule.frx":0AAC
         Left            =   3000
         List            =   "frm_shift_schdule.frx":0AAE
         TabIndex        =   20
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox txt_dept 
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt_fpcode 
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_empname2 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Shift Starting From"
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
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   27
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Selecte Weekly off"
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
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   19
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   17
         Top             =   240
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
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   1095
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
         Index           =   4
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Employee"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1320
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Shift Schedule - for Employees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_shift_schdule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim findrow As Integer
Private Sub Calendar1_Click()
    find_dates
End Sub



Private Sub cmb_weekoff_Click()
''    fillgrid
''    Dim rdate As Date
''
''    Dim i As Integer
''    i = 0
''    For rdate = st_date To end_date
''         findrow = flx_data.Rows - 1
''        If UCase(Format(rdate, "dddd")) <> cmb_weekoff.Text Then
''            flx_data.TextMatrix(findrow, 0) = findrow
''            If i = 0 Then
''               flx_data.TextMatrix(findrow, 1) = rdate
''               i = i + 1
''            Else
''               flx_data.TextMatrix(findrow, 2) = rdate
''
''            End If
''        Else
''            flx_data.Rows = flx_data.Rows + 1
''            i = 0
''        End If
''    Next
End Sub

Private Sub cmb_weekoff_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd_filter_Click()
    
    Dim payrs As New ADODB.Recordset
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
       sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
       sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          cmb_weekoff.Text = payrs!emp_holiday
          payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub cmd_go_Click()
    If st_date.Value = end_date.Value Then
       MsgBox ("Select day from Calender...")
       Exit Sub
    End If
    fillgrid
    Dim rdate As Date
  
    Dim i As Integer
    i = 0
    For rdate = st_date To end_date
         findrow = flx_data.Rows - 1
        If UCase(Format(rdate, "dddd")) <> cmb_weekoff.Text Then
            flx_data.TextMatrix(findrow, 0) = findrow
            If i = 0 Then
               flx_data.TextMatrix(findrow, 1) = rdate
               i = i + 1
            Else
               flx_data.TextMatrix(findrow, 2) = rdate
         
            End If
        Else
            flx_data.Rows = flx_data.Rows + 1
            i = 0
        End If
    Next
    
    If cmb_shift.Text = "" Then
       MsgBox ("Select Employee shift...")
       Exit Sub
    End If
    Dim sft As String
    sft = cmb_shift.Text
    For i = 1 To flx_data.Rows - 1
        flx_data.TextMatrix(i, 3) = sft
        If sft = "A SHIFT" Then
           sft = "B SHIFT"
        ElseIf sft = "B SHIFT" Then
           sft = "C SHIFT"
        ElseIf sft = "C SHIFT" Then
           sft = "A SHIFT"
        End If
    Next


End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fillgrid
    Calendar1.Value = Now
    cmb_weekoff.AddItem "SUNDAY"
    cmb_weekoff.AddItem "MONDAY"
    cmb_weekoff.AddItem "TUESDAY"
    cmb_weekoff.AddItem "WEDNESDAY"
    cmb_weekoff.AddItem "THURSDAY"
    cmb_weekoff.AddItem "FRIDAY"
    cmb_weekoff.AddItem "SATURDAY"
    
    cmb_shift.AddItem "A SHIFT"
    cmb_shift.AddItem "B SHIFT"
    cmb_shift.AddItem "C SHIFT"
    cmb_shift.AddItem "GS"
    cmb_shift.AddItem "6.00AM-6.00PM"
    cmb_shift.AddItem "6.00PM-6.00AM"
    cmb_shift.AddItem "A+6.00PM-6.00AM"
    cmb_shift.AddItem "8.00 PM to 8.00 AM"
    cmb_shift.AddItem "B+C"
    cmb_shift.AddItem "A+B+C"
''    cmb_shift.AddItem "WO"
    cmb_shift.AddItem "Unshift_neight"
    cmb_shift.AddItem "G+C"
    cmb_shift.AddItem "A+C"
    cmb_shift.AddItem "G1/2+C"
    cmb_shift.AddItem "G+NEXTDAY"
    
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear

    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close



End Sub

Private Sub lst_dept_Click()
    lst_employee.Clear
    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 4
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "From Date"
     .TextMatrix(0, 2) = "To Date"
     .TextMatrix(0, 3) = "Shift"
     .ColWidth(0) = 500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 2000
     
     .Redraw = True
   End With
End Sub
Public Sub find_dates()
    Dim d1 As Date
    Dim mdays As Integer
    
    mmon = Calendar1.Month
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And Calendar1.Year Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + Str(Calendar1.Year))
    st_date = end_date - Day(end_date) + 1
    end_date = DateValue(Str(12) + "/" + Str(31) + "/" + Str(Calendar1.Year))
End Sub


Private Sub save_Click()
   If flx_data.Rows - 1 = 0 Then Exit Sub
   If txt_fpcode.Text = "" Then
         MsgBox ("Select Employee Name")
         Exit Sub
   End If
 

paydb.BeginTrans
On Error GoTo err_handler
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset
    
    pst_qry = "delete from bio_shift_schedule where emps_fpcode = '" & txt_fpcode.Text & "' and emps_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "'"
    paydb.Execute pst_qry
    
    pst_qry = "select * from bio_shift_schedule"
    payrs.Open pst_qry, paydb, 1, 2
    Dim rdate, sdate, edate As Date
    Dim i As Integer
    
    For i = 1 To flx_data.Rows - 1
        sdate = Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy")
        edate = Format(flx_data.TextMatrix(i, 2), "dd/MM/yyyy")
        For rdate = sdate To edate
            sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date,emps_shift_alloted, emps_shift) values ( " & txt_fpcode.Text & ", '" & Format(rdate, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 3) & "','" & flx_data.TextMatrix(i, 3) & "')"
            paydb.Execute sql
        Next
        If i <> flx_data.Rows - 1 Then
          sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date,emps_shift_alloted, emps_shift) values ( " & txt_fpcode.Text & ", '" & Format(rdate, "MM/dd/yyyy") & "','WO','WO')"
          paydb.Execute sql
        End If
        
    Next
    
    
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
  
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    cmb_weekoff.Text = ""
    fillgrid
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

