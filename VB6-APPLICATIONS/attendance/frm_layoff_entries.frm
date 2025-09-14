VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_layoff_entries 
   Caption         =   "LAYOFF ENTRIES"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   20340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   360
      TabIndex        =   30
      Top             =   1080
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   31
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   3480
      TabIndex        =   18
      Top             =   1080
      Width           =   9015
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   6000
         TabIndex        =   24
         Top             =   360
         Width           =   2895
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   375
            Left            =   1200
            TabIndex        =   25
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   61276161
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dt_to 
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   61276161
            CurrentDate     =   42278
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
            Left            =   120
            TabIndex        =   28
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
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign Layoff"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   23
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "Modify Layoff"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   21
         Top             =   4560
         Width           =   2655
      End
      Begin VB.CommandButton cmd_view_leaves 
         Caption         =   "View Layoff details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete Layoff"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   19
         Top             =   3360
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   1455
         Left            =   240
         TabIndex        =   22
         Top             =   3720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2566
         _Version        =   393216
      End
      Begin VB.PictureBox lst_view 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   240
         ScaleHeight     =   2955
         ScaleWidth      =   5595
         TabIndex        =   29
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4440
      TabIndex        =   15
      Top             =   6840
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_layoff_entries.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_layoff_entries.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   6255
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
         Left            =   840
         TabIndex        =   12
         Top             =   120
         Width           =   3015
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
         Left            =   4680
         TabIndex        =   11
         Top             =   120
         Width           =   1335
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61276161
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61276161
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frame_group 
      Caption         =   "DETAILS FOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   480
      Width           =   5325
      Begin VB.Frame Frame8 
         Caption         =   "Frame5"
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   15
      End
      Begin VB.OptionButton opt_worker 
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
         Height          =   345
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton opt_staff 
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
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   6960
      TabIndex        =   41
      Top             =   6840
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Layoff Entries - for Employees"
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
      Left            =   720
      TabIndex        =   42
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_layoff_entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim fpcode As Integer
''Dim no As Integer
''Dim rdate As Date
''Dim del_leave As Integer
''Private Sub cmb_leave_KeyPress(KeyAscii As Integer)
''    KeyAscii = 0
''End Sub
''
''Private Sub cmb_month_Click()
''    find_dates
''End Sub
''
''Private Sub cmb_year_Click()
''   find_dates
''End Sub
''
''Private Sub cmd_Assign_Click()
''    Dim pst_qry As String
''    Dim payrs As New ADODB.Recordset
''
''
''
''    Dim iSelected As Integer
''    Dim item As ListItem
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''          iSelected = iSelected + 1
''        End If
''    Next
''    If iSelected = 0 Then
''       MsgBox ("Employee Not selected in the view...")
''       Exit Sub
''    End If
''
''
''    Dim idate As Date
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''           For idate = dt_from To dt_to
''               pst_qry = "select * from bio_empleave where emp_fpcode = " & lst_view.ListItems(i).Text & " and emp_leave_date = '" & Format(idate, "MM/dd/yyyy") & "'"
''               payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''               If Not payrs.EOF Then
''''                     MsgBox ("Already leave assigned for " + lst_view.ListItems(i).Text + " Date " + Format(idate, "dd/MM/yyyy"))
''                     MsgBox ("Already layoff assigned for " + lst_view.ListItems.item(i).SubItems(1) + " Date " + Format(idate, "dd/MM/yyyy"))
''
''                     payrs.Close
''                     Exit Sub
''               End If
''               payrs.Close
''           Next
''        End If
''    Next
''
''    Dim weekoff As String
''
''
''
''paydb.BeginTrans
''On Error GoTo err_handler
''    pst_qry = "select max(emp_leave_no)+1 as endno from bio_empleave"
''
''    payrs.Open pst_qry, paydb, 1, 2
''    no = 1
''    If Not IsNull(payrs!endno) Then
''        If Not payrs.EOF Then
''             no = payrs!endno
''        End If
''    End If
''    payrs.Close
''    Dim ltype, sql As String
''    weekoff = "SUNDAY"
''    For i = 1 To lst_view.ListItems.Count
''        pst_qry = "select * from emp_mas where emp_fpcode = " & lst_view.ListItems(i).Text & " and emp_status = 'A'"
''        payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''        If Not payrs.EOF Then
''           weekoff = payrs("emp_holiday")
''        End If
''        payrs.Close
''
''        If lst_view.ListItems(i).Checked = True Then
''           For idate = dt_from To dt_to
''               If UCase(WeekdayName(Weekday(idate))) <> weekoff Then
''                  sql = "Select * from emp_dec_holiday where emp_dec_holiday = '" & Format(idate, "mm/dd/yyyy") & "'"
''                  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''                  If payrs.EOF Then
''                     sql = "insert into bio_empleave (emp_leave_no , emp_fpcode, emp_leave_type, emp_leave_date, emp_leave_period) values (" & no & ", " & lst_view.ListItems(i).Text & ", 'LAYOFF','" & Format(idate, "MM/dd/yyyy") & "','F')"
''                     paydb.Execute sql
''                  End If
''                  payrs.Close
''               End If
''           Next
''        End If
''    Next
''    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
''
''    Exit Sub
''Exit Sub
''err_handler:
''        paydb.RollbackTrans
''        chk = gen_Validation(Err.Number, Err.Description)
''
''End Sub
''
''Private Sub cmd_delete_Click()
''    del_leave = 1
''    cmd_Assign.Enabled = False
''    cmd_modify.Enabled = True
''    flx_data.Enabled = True
''End Sub
''
''Private Sub cmd_filter_Click()
''    Dim chk As Integer
''    chk = 0
''    Refresh_Click
''    fillgrid
''    Dim payrs As New ADODB.Recordset
''    Dim itmx As ListItem
''    lst_view.ColumnHeaders.Clear
''    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
''    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
''    lst_view.ColumnHeaders.Add , , "Department ", 1500
''    lst_view.View = lvwReport
''    lst_view.ListItems.Clear
''
''    If txt_empcode.Text <> "" Then
''        sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
''    ElseIf txt_empname.Text <> "" Then
''        sql = "select * from bio_empmas where bioemp_name like  '%" & txt_empname.Text & "%' and bioemp_status = 'Working' order by bioemp_dept"
''    ElseIf lst_employee.Text <> "" Then
''        sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_dept = '" & lst_dept.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
''    End If
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''            Set itmx = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
''            itmx.SubItems(1) = payrs.Fields("bioemp_name")
''            itmx.SubItems(2) = payrs.Fields("bioemp_dept")
''            If chk = 0 Then
''               itmx.Checked = True
''            End If
''            chk = 1
''            payrs.MoveNext
''    Wend
''    payrs.Close
''
''    pst_qry = "select * from bio_empleave  a ,emp_mas b  where  a.emp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_empcode.Text & "' and  emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and emp_leave_type  ='LAYOFF' and emp_status = 'A' order by emp_leave_date"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''        flx_data.TextMatrix(i, 0) = i
''        flx_data.TextMatrix(i, 1) = payrs!emp_name
''        flx_data.TextMatrix(i, 2) = Format(payrs!emp_leave_date, "dd/MM/yyyy")
''        If payrs!emp_leave_period = "F" Then
''           flx_data.TextMatrix(i, 3) = "FULL DAY"
''        Else
''           flx_data.TextMatrix(i, 3) = "HALF DAY"
''        End If
''        flx_data.TextMatrix(i, 4) = payrs!emp_leave_no
''        flx_data.TextMatrix(i, 5) = payrs!emp_leave_type
''        flx_data.TextMatrix(i, 6) = payrs!emp_leave_period
''        flx_data.Rows = flx_data.Rows + 1
''
''        flx_dataold.TextMatrix(i, 0) = i
''        flx_dataold.TextMatrix(i, 1) = payrs!emp_name
''        flx_dataold.TextMatrix(i, 2) = Format(payrs!emp_leave_date, "dd/MM/yyyy")
''        If payrs!emp_leave_period = "F" Then
''           flx_dataold.TextMatrix(i, 3) = "FULL DAY"
''        Else
''           flx_dataold.TextMatrix(i, 3) = "HALF DAY"
''        End If
''        flx_dataold.TextMatrix(i, 4) = payrs!emp_leave_no
''        flx_dataold.TextMatrix(i, 5) = payrs!emp_leave_type
''        flx_dataold.TextMatrix(i, 6) = payrs!emp_leave_period
''
''        payrs.MoveNext
''        flx_dataold.Rows = flx_dataold.Rows + 1
''
''        i = i + 1
''    Wend
''    payrs.Close
''
''
''
''End Sub
''
''
''Private Sub cmd_modify_Click()
''   fpcode = lst_view.SelectedItem.Text
''
''
''paydb.BeginTrans
''On Error GoTo err_handler
''
''    Dim pst_qry As String
''    Dim payrs As New ADODB.Recordset
''
''    Dim rdate, sdate, edate As Date
''    For i = 1 To flx_dataold.Rows - 1
''        sdate = Format(flx_dataold.TextMatrix(i, 2), "dd/MM/yyyy")
''        sql = "delete from bio_empleave where emp_fpcode = " & fpcode & "  and emp_leave_date  = '" & Format(sdate, "MM/dd/yyyy") & "'  and emp_leave_type  ='LAYOFF'"
''        paydb.Execute sql
''    Next
''
''    For i = 1 To flx_data.Rows - 1
''        If Val(flx_data.TextMatrix(i, 0)) > 0 Then
''            sdate = Format(flx_data.TextMatrix(i, 2), "dd/MM/yyyy")
''            sql = "insert into bio_empleave (emp_leave_no , emp_fpcode, emp_leave_type, emp_leave_date, emp_leave_period) values (" & Val(flx_data.TextMatrix(i, 4)) & ", " & fpcode & ", '" & flx_data.TextMatrix(i, 5) & "','" & Format(sdate, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 6) & "')"
''            paydb.Execute sql
''        End If
''    Next
''
''    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
''
''    Exit Sub
''Exit Sub
''err_handler:
''        paydb.RollbackTrans
''        chk = gen_Validation(Err.Number, Err.Description)
''
''End Sub
''
''Private Sub Command1_Click()
''
''End Sub
''
''Private Sub cmd_view_leaves_Click()
''    fpcode = lst_view.SelectedItem.Text
''    fillgrid
''    Dim payrs As New ADODB.Recordset
''    pst_qry = "select * from bio_empleave  a ,emp_mas b  where  a.emp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and emp_leave_type  ='LAYOFF'  and emp_status = 'A' order by emp_leave_date"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''        flx_data.TextMatrix(i, 0) = i
''        flx_data.TextMatrix(i, 1) = payrs!emp_name
''        flx_data.TextMatrix(i, 2) = Format(payrs!emp_leave_date, "dd/MM/yyyy")
''        If payrs!emp_leave_period = "F" Then
''           flx_data.TextMatrix(i, 3) = "FULL DAY"
''        Else
''           flx_data.TextMatrix(i, 3) = "HALF DAY"
''        End If
''        flx_data.TextMatrix(i, 4) = payrs!emp_leave_no
''        flx_data.TextMatrix(i, 5) = payrs!emp_leave_type
''        flx_data.TextMatrix(i, 6) = payrs!emp_leave_period
''        flx_data.Rows = flx_data.Rows + 1
''
''        flx_dataold.TextMatrix(i, 0) = i
''        flx_dataold.TextMatrix(i, 1) = payrs!emp_name
''        flx_dataold.TextMatrix(i, 2) = Format(payrs!emp_leave_date, "dd/MM/yyyy")
''        If payrs!emp_leave_period = "F" Then
''           flx_dataold.TextMatrix(i, 3) = "FULL DAY"
''        Else
''           flx_dataold.TextMatrix(i, 3) = "HALF DAY"
''        End If
''        flx_dataold.TextMatrix(i, 4) = payrs!emp_leave_no
''        flx_dataold.TextMatrix(i, 5) = payrs!emp_leave_type
''        flx_dataold.TextMatrix(i, 6) = payrs!emp_leave_period
''
''        payrs.MoveNext
''        flx_dataold.Rows = flx_dataold.Rows + 1
''
''        i = i + 1
''    Wend
''    payrs.Close
''''  flx_data.Enabled = False
''
''End Sub
''
''Private Sub dt_from_Change()
''     dt_to.Value = dt_from.Value
''End Sub
''
''Private Sub exit_Click()
''     Unload Me
''End Sub
''
''Private Sub flx_data_DblClick()
''   If del_leave = 0 Then Exit Sub
''   flex_edit_row = 0
''   Dim fin_selrow As Integer
''   Dim pst_ans As String
''   fin_selrow = flx_data.Row
''
''   timchk = 0
''   With flx_data
''       pst_ans = MsgBox("Press YES-to DELETE  NO-to CANCEL", vbYesNo, "Confirmation")
''       If pst_ans = 6 Then
''               If .Rows < 2 Then
''                  MsgBox "No rows to remove"
''               Else
''                  If Val(flx_data.TextMatrix(.Row, 0)) > 0 Then
''                     flx_data.RemoveItem fin_selrow
''                  End If
''                  .Row = flx_data.Rows - 1
''               End If
''        End If
''   End With
''
''End Sub
''
''Private Sub Form_Load()
''    del_leave = 0
''    dt_from.Value = Now
''    dt_to.Value = Now
''
''    With cmb_month
''        .AddItem "January"
''        .ItemData(.NewIndex) = 1
''        .AddItem "February"
''        .ItemData(.NewIndex) = 2
''        .AddItem "March"
''        .ItemData(.NewIndex) = 3
''        .AddItem "April"
''        .ItemData(.NewIndex) = 4
''        .AddItem "May"
''        .ItemData(.NewIndex) = 5
''        .AddItem "June"
''        .ItemData(.NewIndex) = 6
''        .AddItem "July"
''        .ItemData(.NewIndex) = 7
''        .AddItem "August"
''        .ItemData(.NewIndex) = 8
''        .AddItem "September"
''        .ItemData(.NewIndex) = 9
''        .AddItem "October"
''        .ItemData(.NewIndex) = 10
''        .AddItem "November"
''        .ItemData(.NewIndex) = 11
''        .AddItem "December"
''        .ItemData(.NewIndex) = 12
''    End With
''    With cmb_year
''''        .AddItem finyear + 2000
''''        .AddItem "2012"
''''        .AddItem "2013"
''''        .AddItem "2014"
''''        .AddItem "2015"
''''        .AddItem "2016"
''''        .Text = "2015"
''      .AddItem Left(fyear, 4)
''      .AddItem Mid(fyear, 6, 4)
''      If Year(Date) = Int(Left(fyear, 4)) Then
''         cmb_year.Text = Left(fyear, 4)
''      Else
''          cmb_year.Text = Mid(fyear, 6, 4)
''      End If
''
''    End With
''    cmb_month.ListIndex = Month(Date) - 1
''
''    Dim payrs As New ADODB.Recordset
''    lst_dept.Clear
''
''    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        lst_dept.AddItem payrs("bioemp_dept")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''    Dim itmx As ListItem
''    lst_view.ColumnHeaders.Clear
''    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
''    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
''    lst_view.ColumnHeaders.Add , , "Department ", 1500
''    lst_view.ColumnHeaders.Add , , "Type ", 1000
''    lst_view.View = lvwReport
''    lst_view.ListItems.Clear
''
''    sql = "select * from bio_empmas where bioemp_status = 'Working' order by bioemp_dept"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''            Set itmx = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
''            itmx.SubItems(1) = payrs.Fields("bioemp_name")
''            itmx.SubItems(2) = payrs.Fields("bioemp_dept")
''            itmx.SubItems(3) = payrs.Fields("bioemp_team")
''            payrs.MoveNext
''    Wend
''
''    payrs.Close
''    fillgrid
''
''End Sub
''
''Private Sub lst_dept_Click()
''    lst_view.ListItems.Clear
''    lst_employee.Clear
''    If opt_all.Value = True Then
''       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name"
''    ElseIf opt_staff.Value = True Then
''       sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A'  order by bioemp_name"
''    Else
''       sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'W'  and emp_status = 'A'  order by bioemp_name"
''    End If
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        lst_employee.AddItem payrs("bioemp_name")
''        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''End Sub
''
''Private Sub opt_all_Click()
''    lst_dept_Click
''End Sub
''
''Private Sub opt_leave_full_Click()
''    lbl2.Visible = True
''    dt_to.Visible = True
''    frame_leavetype.Visible = False
''    lbl.Caption = "From Date"
''
''End Sub
''
''
''
''Private Sub opt_leave_half_Click()
''    lbl2.Visible = False
''    dt_to.Visible = False
''    frame_leavetype.Visible = True
''    lbl.Caption = "Leave for"
''End Sub
''
''Private Sub fillgrid()
''   With flx_data
''     .Redraw = False
''     .Clear
''     .Rows = 2
''     .Cols = 7
''     .TextMatrix(0, 0) = "S.No"
''     .TextMatrix(0, 1) = "Name"
''     .TextMatrix(0, 2) = "Leave Date"
''     .TextMatrix(0, 3) = "Period"
''     .TextMatrix(0, 4) = "NO"
''     .TextMatrix(0, 5) = "Type"
''     .TextMatrix(0, 6) = "Periodtype"
''
''     .ColWidth(0) = 500
''     .ColWidth(1) = 2000
''     .ColWidth(2) = 1500
''     .ColWidth(3) = 1200
''     .ColWidth(4) = 0
''     .ColWidth(5) = 1000
''     .ColWidth(6) = 0
''     .Redraw = True
''   End With
''   With flx_dataold
''     .Redraw = False
''     .Clear
''     .Rows = 2
''     .Cols = 7
''     .TextMatrix(0, 0) = "S.No"
''     .TextMatrix(0, 1) = "Name"
''     .TextMatrix(0, 2) = "Leave Date"
''     .TextMatrix(0, 3) = "Period"
''     .TextMatrix(0, 4) = "NO"
''     .TextMatrix(0, 5) = "Type"
''     .TextMatrix(0, 6) = "Periodtype"
''     .ColWidth(0) = 500
''     .ColWidth(1) = 2000
''     .ColWidth(2) = 1500
''     .ColWidth(3) = 1200
''     .ColWidth(4) = 0
''     .ColWidth(5) = 0
''     .ColWidth(6) = 0
''
''     .Redraw = True
''   End With
''
''End Sub
''
''
''Public Sub find_dates()
''    If cmb_month.ListIndex = -1 Then Exit Sub
''    Dim d1 As Date
''    mmon = cmb_month.ItemData(cmb_month.ListIndex)
''    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
''        mdays = 31
''    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
''        mdays = 30
''    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
''        mdays = 29
''    Else
''        mdays = 28
''    End If
''    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
''    st_date = end_date - Day(end_date) + 1
''End Sub
''
''Private Sub opt_staff_Click()
''   lst_dept_Click
''End Sub
''
''Private Sub opt_worker_Click()
''   lst_dept_Click
''End Sub
''
''Private Sub Refresh_Click()
''    flx_data.Enabled = True
''    cmd_Assign.Enabled = True
''    cmd_modify.Enabled = False
''    del_leave = 0
''End Sub
''
Private Sub cmd_Assign_Click()

End Sub
