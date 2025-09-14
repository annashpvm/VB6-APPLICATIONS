VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_saturday 
   Caption         =   "STATURDAY ABSENT"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   18285
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flx_data 
      Height          =   5295
      Left            =   12840
      TabIndex        =   34
      Top             =   2520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "SAVE"
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
      Left            =   14400
      TabIndex        =   33
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   9480
      TabIndex        =   28
      Top             =   360
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22282241
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22282241
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frame_month 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   12255
      Begin VB.CommandButton cmd_ref 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   35
         Top             =   120
         Width           =   375
      End
      Begin VB.ComboBox cmb_saturday 
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
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   2535
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
         Left            =   6360
         TabIndex        =   23
         Top             =   360
         Width           =   1335
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
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "DATE"
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
         Index           =   5
         Left            =   8280
         TabIndex        =   27
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
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   1335
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
         Index           =   4
         Left            =   5400
         TabIndex        =   24
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4560
      TabIndex        =   14
      Top             =   7800
      Width           =   1815
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_saturday.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_saturday.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   3720
      TabIndex        =   12
      Top             =   2520
      Width           =   8535
      Begin ComctlLib.ListView lst_view 
         Height          =   2535
         Left            =   600
         TabIndex        =   36
         Top             =   1320
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4471
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame9 
         Caption         =   "EMP Type"
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
         Height          =   855
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   5775
         Begin VB.OptionButton opt_regular 
            Caption         =   "Regular"
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
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_vou 
            Caption         =   "Voucher"
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
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_cs 
            Caption         =   "CS"
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
            Left            =   3720
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign "
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
         Left            =   2400
         TabIndex        =   13
         Top             =   4560
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   4920
         Width           =   975
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
         TabIndex        =   11
         Top             =   720
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
         TabIndex        =   10
         Top             =   240
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
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "SATURDAY Entries - for Employees"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "frm_saturday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim fpcode As Integer
''Dim no As Integer
''Dim rdate As Date
''Dim del_leave As Integer
''Dim codelist, sel_codes As String
''
''
''Private Sub cmb_month_Change()
''  find_dates
''End Sub
''
''Private Sub cmb_month_Click()
''    find_dates
''End Sub
''
''Private Sub cmb_saturday_Change()
''    fillgrid
''End Sub
''
''Private Sub cmb_year_Change()
''    find_dates
''End Sub
''
''Private Sub cmb_year_Click()
''    find_dates
''End Sub
''
''Private Sub cmd_ref_Click()
''    sql = "select ds_date from bio_device_shiftlogs where ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' group by ds_date order by ds_date"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''         cmb_saturday.AddItem payrs!ds_date
''         payrs.MoveNext
''    Wend
''    payrs.Close
''End Sub
''
''Private Sub cmd_save_Click()
''    If cmb_month.Text = "" Then
''       MsgBox ("Select Month ..")
''    End If
''    If cmb_year.Text = "" Then
''       MsgBox ("Select year ..")
''    End If
''    If cmb_saturday.Text = "" Then
''       MsgBox ("Select Date ..")
''    End If
''    sql = "delete from bio_emp_saturday where empsat_date = '" & Format(cmb_saturday, "MM/dd/yyyy") & "'"
''    paydb.Execute sql
''    For i = 1 To flx_data.Rows - 1
''        If flx_data.TextMatrix(i, 2) = "Y" Then
''           sql = "insert into bio_emp_saturday  (empsat_fpcode,empsat_name,empsat_date) values (" & flx_data.TextMatrix(i, 0) & ", '" & flx_data.TextMatrix(i, 1) & "', '" & Format(cmb_saturday, "MM/dd/yyyy") & "')"
''           paydb.Execute sql
''        End If
''    Next
''
''    MsgBox ("Saturday Salary Cut Assigned Sucessfully...")
''
''
''End Sub
''
''Private Sub exit_Click()
''     Unload Me
''End Sub
''
''Private Sub flx_data_Click()
''   fin_selrow = flx_data.Row
''    findatacol = flx_data.Col
''    Select Case flx_data.Col
''       Case 2
''          If flx_data.TextMatrix(flx_data.Row, 2) = "Y" Then
''             flx_data.TextMatrix(flx_data.Row, 2) = "N"
''          Else
''             flx_data.TextMatrix(flx_data.Row, 2) = "Y"
''          End If
''    End Select
''    Exit Sub
''
''
''End Sub
''
''Private Sub Form_Load()
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
''    Refresh_Click
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
''    fillgrid
''
''
''    sql = "select empsat_fpcode,empsat_name from bio_emp_saturday group by empsat_fpcode,empsat_name order by empsat_fpcode,empsat_name"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''          flx_data.TextMatrix(flx_data.Rows - 1, 0) = payrs.Fields("empsat_fpcode")
''          flx_data.TextMatrix(flx_data.Rows - 1, 1) = payrs.Fields("empsat_name")
''          flx_data.TextMatrix(flx_data.Rows - 1, 2) = "Y"
''          flx_data.Rows = flx_data.Rows + 1
''
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''End Sub
''
''Private Sub lst_dept_Click()
''    lst_employee.Clear
''    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        lst_employee.AddItem payrs("bioemp_name")
''        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
''        payrs.MoveNext
''    Wend
''    payrs.Close
''
''End Sub
''
''
''
''
''Private Sub cmd_filter_Click()
''    Dim chk As Integer
''    chk = 0
''     Refresh_Click
''
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
''      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
''    ElseIf txt_empname.Text <> "" Then
''      sql = "select * from bio_empmas where bioemp_name like  '%" & txt_empname.Text & "%' and bioemp_status = 'Working' order by bioemp_dept"
''    ElseIf lst_employee.Text <> "" Then
''          sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
''    Else
''         MsgBox ("Employee code Not found ....")
''         Exit Sub
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
''''    If opt_vou.Value = True Then
''''       pst_qry = "select * from bio_emp_oddetails   a ,emp_voupay_mast b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and emp_status = 'A' order by empod_date desc"
''''    ElseIf opt_cs.Value = True Then
''''           pst_qry = "select * from bio_emp_oddetails   a ,mas_caemp b  where  a.empod_fpcode = b.ca_fpcode and b.ca_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ca_status = 'A' order by empod_date desc"
''''    Else
''''       pst_qry = "select * from bio_emp_oddetails   a ,emp_mas b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and emp_status = 'A' order by empod_date desc"
''''    End If
''''
''''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''''    i = 1
''''    While Not payrs.EOF
''''        flx_data.TextMatrix(i, 0) = i
''''        If opt_cs.Value = True Then
''''           flx_data.TextMatrix(i, 1) = payrs!ca_empname
''''        Else
''''           flx_data.TextMatrix(i, 1) = payrs!emp_name
''''        End If
''''        flx_data.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
''''        flx_data.TextMatrix(i, 3) = payrs!empod_fromtime
''''        flx_data.TextMatrix(i, 4) = payrs!empod_totime
''''        flx_data.TextMatrix(i, 5) = payrs!empod_location
''''        flx_data.TextMatrix(i, 6) = payrs!empod_purpose
''''        flx_data.TextMatrix(i, 7) = payrs!empod_no
''''        flx_data.Rows = flx_data.Rows + 1
''''
''''        flx_dataold.TextMatrix(i, 0) = i
''''''        flx_dataold.TextMatrix(i, 1) = payrs!emp_name
''''        If opt_cs.Value = True Then
''''           flx_dataold.TextMatrix(i, 1) = payrs!ca_empname
''''        Else
''''           flx_dataold.TextMatrix(i, 1) = payrs!emp_name
''''        End If
''''
''''        flx_dataold.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
''''        flx_dataold.TextMatrix(i, 3) = payrs!empod_fromtime
''''        flx_dataold.TextMatrix(i, 4) = payrs!empod_totime
''''        flx_dataold.TextMatrix(i, 5) = payrs!empod_location
''''        flx_dataold.TextMatrix(i, 6) = payrs!empod_purpose
''''        flx_dataold.TextMatrix(i, 7) = payrs!empod_no
''''        payrs.MoveNext
''''        flx_dataold.Rows = flx_dataold.Rows + 1
''''
''''        i = i + 1
''''    Wend
''''    payrs.Close
''
''
''
''End Sub
''
''Private Sub cmd_Assign_Click()
''    Dim pst_qry As String
''    Dim payrs As New ADODB.Recordset
''
''
''    codelist = ""
''    sel_codes = ""
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
''    codelist = "(10"
''    Dim idate As Date
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''          codelist = codelist + ", " + lst_view.ListItems(i).Text
''          If sel_codes = "" Then
''             sel_codes = "{bio_emp_saturday.empsat_fpcode} = " & lst_view.ListItems(i).Text
''          Else
''             sel_codes = sel_codes + " or {bio_emp_saturday.empsat_fpcode} = " & lst_view.ListItems(i).Text
''          End If
''        End If
''    Next
''    codelist = codelist + ")"
''    For i = 1 To lst_view.ListItems.Count
''        If lst_view.ListItems(i).Checked = True Then
''               pst_qry = "select * from bio_emp_saturday  where empsat_fpcode in " & codelist
''               payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''               If Not payrs.EOF Then
''''                     MsgBox ("Already leave assigned for " + lst_view.ListItems(i).Text + " Date " + Format(idate, "dd/MM/yyyy"))
''                     MsgBox ("Already assigned ")
''
''                     payrs.Close
''                     Exit Sub
''               End If
''               payrs.Close
''        End If
''    Next
''
''paydb.BeginTrans
''On Error GoTo err_handler
''''    If opt_leave_half.Value = True Then
''''       dt_to.Value = dt_from.Value
''''    End If
''      sql = "insert into bio_emp_saturday  (empsat_fpcode,empsat_name,empsat_date) values (" & lst_view.ListItems(1).Text & ", '" & lst_employee.Text & "', '" & Format(cmb_saturday, "MM/dd/yyyy") & "')"
''      paydb.Execute sql
''    paydb.CommitTrans
''    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
''    Exit Sub
''Exit Sub
''err_handler:
''        paydb.RollbackTrans
''        chk = gen_Validation(Err.Number, Err.Description)
''
''End Sub
''
''Private Sub Refresh_Click()
''    cmd_Assign.Enabled = True
''''    cmd_modify.Enabled = False
''''    del_leave = 0
''''    txt_pw.Visible = False
''''    txt_pw.Text = ""
''End Sub
''
''
''
''Public Sub find_dates()
''    If cmb_month.ListIndex = -1 Then Exit Sub
''    If cmb_year.Text = "" Then Exit Sub
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
''    st_date = end_date - Day(end_date - 1)
''
''    cmb_saturday.Clear
''
''    Dim daycount As Integer
''    daycount = 1
''    sql = "select ds_date from bio_device_shiftlogs where ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "'  and datepart(dw,ds_date) =  7  group by ds_date"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''         If daycount = 2 Or daycount = 4 Then
''            cmb_saturday.AddItem payrs!ds_date
''         End If
''         daycount = daycount + 1
''         payrs.MoveNext
''    Wend
''    payrs.Close
''
''
''End Sub
''
''
''Function fillgrid()
''    With flx_data
''        .Clear
''        .Cols = 3
''        .TextMatrix(0, 0) = "Emp code"
''        .TextMatrix(0, 1) = "Name"
''        .TextMatrix(0, 2) = "Y / N "
''
''        .ColWidth(0) = 1000
''        .ColWidth(1) = 1500
''        .ColWidth(2) = 1000
''    End With
''End Function
''
