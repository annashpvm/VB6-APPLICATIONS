VERSION 5.00
Begin VB.Form frm_dec_holiday_entry 
   Caption         =   "Declare holiday Entry"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ListView1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8040
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   3000
      TabIndex        =   11
      Top             =   1920
      Width           =   6015
      Begin VB.OptionButton opt_vou 
         Caption         =   "Voucher"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Width           =   1665
      End
      Begin VB.OptionButton opt_worker 
         Caption         =   "Worker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt_staff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   10935
      Begin VB.ComboBox cmb_year 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmb_holiday 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Select year"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Select Declare Holiday"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4560
      TabIndex        =   5
      Top             =   6960
      Width           =   2175
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_dec_holiday_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_dec_holiday_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Employees"
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
      Height          =   3975
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   11055
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign Eligibility"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9000
         TabIndex        =   3
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmd_allselect 
         Caption         =   "Select All"
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
         Left            =   9000
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_deselect 
         Caption         =   "Deselect All"
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
         Left            =   9000
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox lst_view 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   2640
         ScaleHeight     =   3555
         ScaleWidth      =   5595
         TabIndex        =   14
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE DECLARE HOLIDAY ENTRY"
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
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "frm_dec_holiday_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Sub cmb_year_Click()
''    Dim payrs As New ADODB.Recordset
''    Dim fdate, edata As Date
''    fdate = DateValue("01/01/" + Str(cmb_year.Text))
''    edate = DateValue("12/31/" + Str(cmb_year.Text))
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    sql = "Select * from emp_dec_holiday where emp_dec_holiday between '" & Format(fdate, "MM/dd/yyyy") & "' and '" & Format(edate, "MM/dd/yyyy") & "' order by emp_dec_holiday "
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF
''        cmb_holiday.AddItem Format(payrs!emp_dec_holiday, "dd/MM/yyyy")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''End Sub
''
''Private Sub cmd_allselect_Click()
''    For i = 1 To lst_view.ListItems.Count
''        lst_view.ListItems(i).Checked = True
''    Next
''End Sub
''
''Private Sub cmd_Assign_Click()
''
''    Dim pst_qry As String
''    Dim payrs As New ADODB.Recordset
''
''
''''    Dim iSelected As Integer
''''    Dim item As ListItem
''''    For i = 1 To lst_view.ListItems.Count
''''        If lst_view.ListItems(i).Checked = True Then
''''          iSelected = iSelected + 1
''''        End If
''''    Next
''''    If iSelected = 0 Then
''''       MsgBox ("Employee Not selected in the view...")
''''       Exit Sub
''''    End If
''
''
''    paydb.BeginTrans
''''     sql = " delete from  bio_empdh_eligible where empdh_date = '" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "' and empdh_fpcode in (select bioemp_fpcode from bio_empmas  where bioemp_dept = '" & lst_dept.Text & "') = " & lst_view.ListItems(i).Text & ""
''''    paydb.Execute sql
''''
''    For i = 1 To lst_view.ListItems.Count
''        sql = " delete from  emp_dec_holiday_empwise where emp_decholi_date = '" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "' and emp_decholi_fpcode = '" & lst_view.ListItems(i).Text & "' "
''        paydb.Execute sql
''    Next
''
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''           sql = "insert into emp_dec_holiday_empwise (emp_decholi_fpcode, emp_decholi_date) values ( " & lst_view.ListItems(i).Text & ",'" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "')"
''           paydb.Execute sql
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
''
''End Sub
''
''
''Private Sub cmd_deselect_Click()
''    For i = 1 To lst_view.ListItems.Count
''        lst_view.ListItems(i).Checked = False
''    Next
''End Sub
''
''Private Sub cmd_filter_Click()
''''    If cmb_holiday.Text = "" Then
''''       MsgBox ("Select Declare holiday date...")
''''       Exit Sub
''''    End If
''    Refresh_Click
''    Dim payrs As New ADODB.Recordset
''    Dim itmx As ListItem
''    lst_view.ColumnHeaders.Clear
''    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
''    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
''    lst_view.ColumnHeaders.Add , , "Department ", 1500
''    lst_view.View = lvwReport
''    lst_view.ListItems.Clear
'' ''   sql = "select * from bio_empmas where bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_dept"
''''    sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A'  order by bioemp_dept"
''''    sql = "select  bioemp_fpcode,bioemp_name,bioemp_dept,1 as checkedlist from bio_empmas a, emp_mas b,bio_empdh_eligible c where bioemp_fpcode = emp_fpcode and emp_fpcode=empdh_fpcode and convert(varchar(10),empdh_date,103)= '" & cmb_holiday.Text & "'  and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A' " _
''''           & " Union All " _
''''          & " select  bioemp_fpcode,bioemp_name,bioemp_dept,0 as checkedlist   from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A' and bioemp_fpcode not in (select  bioemp_fpcode from bio_empmas a, emp_mas b,bio_empdh_eligible c where bioemp_fpcode = emp_fpcode and emp_fpcode=empdh_fpcode and convert(varchar(10),empdh_date,103)= '" & cmb_holiday.Text & "'  and emp_classification = 'B' and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A')  order by bioemp_dept"
''''
''    If opt_staff.Value = True Then
''        sql = "select  bioemp_fpcode,bioemp_name,bioemp_dept,1 as checkedlist from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and emp_cat = 'S'  and emp_status = 'A' order by bioemp_dept "
''    ElseIf opt_worker.Value = True Then
''        sql = "select  bioemp_fpcode,bioemp_name,bioemp_dept,1 as checkedlist from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode  and bioemp_status = 'Working' and emp_cat = 'W'  and emp_status = 'A' order by bioemp_dept "
''    Else
''       sql = "select  bioemp_fpcode,bioemp_name,bioemp_dept,1 as checkedlist  from bio_empmas a, emp_voupay_mast b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and emp_cat = 'R'  and emp_status = 'A' order by bioemp_name"
''    End If
''
''
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''
''            Set itmx = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
''''            If payrs.Fields("checkedlist") = 1 Then
''''                itmx.Checked = True
''''            Else
''''                itmx.Checked = False
''''            End If
''            itmx.SubItems(1) = payrs.Fields("bioemp_name")
''            itmx.SubItems(2) = payrs.Fields("bioemp_dept")
''
''            payrs.MoveNext
''    Wend
''    payrs.Close
''    sql = " select * from  emp_dec_holiday_empwise where emp_decholi_date = '" & Format(cmb_holiday.Text, "MM/dd/yyyy") & "'"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''          For i = 1 To lst_view.ListItems.Count
''             If lst_view.ListItems(i).Text = payrs("emp_decholi_fpcode") Then
''                lst_view.ListItems(i).Checked = True
''                Exit For
''             End If
''          Next
''          payrs.MoveNext
''    Wend
''    payrs.Close
''
''
''
''End Sub
''
''Private Sub exit_Click()
''    Unload Me
''End Sub
''
''Private Sub Form_Load()
''
''    With cmb_year
''      .AddItem Left(fyear, 4)
''      .AddItem Mid(fyear, 6, 4)
''      If Year(Date) = Int(Left(fyear, 4)) Then
''         cmb_year.Text = Left(fyear, 4)
''      Else
''          cmb_year.Text = Mid(fyear, 6, 4)
''      End If
''    End With
''
''
''
''End Sub
''
''
''Private Sub opt_staff_Click()
''   lst_dept_Click
''End Sub
''
''Private Sub opt_worker_Click()
''   lst_dept_Click
''End Sub
''
''
''
''Private Sub lst_dept_Click()
''
''    lst_view.ListItems.Clear
''''    lst_view.ListItems.Clear
''''    lst_employee.Clear
''''    If opt_staff.Value = True Then
''''       sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A'  order by bioemp_name"
''''    Else
''''        sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'W'  and emp_status = 'A'  order by bioemp_name"
''''    End If
''''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''''    While Not payrs.EOF()
''''        lst_view.AddItem payrs("bioemp_name")
''''        lst_view.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
''''        payrs.MoveNext
''''    Wend
''''    payrs.Close
''End Sub
''
''
''
''
''Private Sub Refresh_Click()
''    lst_view.ListItems.Clear
''End Sub
''
''
