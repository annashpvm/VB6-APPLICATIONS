VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_rep_bio_metrics_view 
   Caption         =   "BIO - METRIC OUTPUT"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11415
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView lst_view 
      Height          =   4815
      Left            =   3600
      TabIndex        =   31
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8493
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56098817
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56098817
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Select Mill"
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
      Left            =   3360
      TabIndex        =   20
      Top             =   120
      Width           =   7575
      Begin VB.OptionButton opt_cogen 
         Caption         =   "COGEN"
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
         Height          =   195
         Left            =   6120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt_vjpm 
         Caption         =   "VJPM"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "All"
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
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opt_dpm2 
         Caption         =   "DPM-2"
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
         Height          =   195
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_dpm1 
         Caption         =   "DPM-1"
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
         Height          =   195
         Left            =   1320
         TabIndex        =   21
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
      Left            =   360
      TabIndex        =   12
      Top             =   0
      Width           =   3045
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   495
         Left            =   3000
         TabIndex        =   19
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
         Left            =   1560
         TabIndex        =   14
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
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   120
         Picture         =   "frm_rep_bio_metrics_view.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton EXIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&XIT"
         Height          =   825
         Left            =   1080
         Picture         =   "frm_rep_bio_metrics_view.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2655
         Begin VB.ComboBox cmb_year 
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
            Left            =   720
            TabIndex        =   18
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cmb_month 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Month"
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
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Year"
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
            TabIndex        =   6
            Top             =   840
            Width           =   495
         End
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
         TabIndex        =   11
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
   End
End
Attribute VB_Name = "frm_rep_bio_metrics_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcode As String
Dim emp_type As String
Dim mdays As Integer
Private Sub cmb_month_Click()
    finddates
End Sub

Public Sub finddates()
    Dim SDATE, EDATE As Date
    SDATE = CDate("01/" & cmb_month.ItemData(cmb_month.ListIndex) & "/" & cmb_year.Text & "")
    EDATE = DateAdd("m", 1, SDATE) - 1
    mmon = Right("0" + Trim(Str(cmb_month.ItemData(cmb_month.ListIndex))), 2)
    If mmon = "00" Then mmon = "12"
    If mmon = "01" Or mmon = "03" Or mmon = "05" Or mmon = "07" Or mmon = "08" Or mmon = "10" Or mmon = "12" Then
      mdays = "31"
    ElseIf mmon = "04" Or mmon = "06" Or mmon = "09" Or mmon = "11" Then
       mdays = "30"
    Else
        If Val(cmb_year.Text) / 4 = Int(Val(cmb_year.Text) / 4) Then
           mdays = "29"
        Else
          mdays = "28"
        End If
    End If
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text)
    st_date = end_date - Day(end_date) + 1

End Sub


Private Sub cmb_year_Click()
   finddates
End Sub

Private Sub cmd_filter_Click()
    If lst_employee.ListIndex = -1 Then Exit Sub
    If cmb_month.ListIndex = -1 Then Exit Sub
    Dim itmx As ListItem
    Dim payrs As New ADODB.Recordset
    lst_view.ColumnHeaders.Clear
    lst_view.ColumnHeaders.ADD , , "FP Code ", 700
    lst_view.ColumnHeaders.ADD , , "Date ", 1100
    lst_view.ColumnHeaders.ADD , , "In Time ", 1100
    lst_view.ColumnHeaders.ADD , , "Out Time ", 2100
    lst_view.ColumnHeaders.ADD , , "Status", 1000
    lst_view.View = lvwReport
    lst_view.ListItems.Clear
''    If opt_all.Value = True Then
''       sql = "select * from bio_attendlogs where a_fpcode = " & lst_employee.ItemData(lst_employee.ListIndex) & " and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_year = " & Val(cmb_year.Text)
''    Else
''       sql = "select * from bio_attendlogs  , emp_mas where emp_company = " & mcode & " and emp_fpcode = a_fpcode and a_fpcode = " & lst_employee.ItemData(lst_employee.ListIndex) & " and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_year = " & Val(cmb_year.Text)
''    End If
    
    If opt_all.Value = True Then
       sql = "select * from bio_device_shiftlogs where ds_fpcode = " & lst_employee.ItemData(lst_employee.ListIndex) & " and ds_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' order by ds_date "
    Else
       sql = "select * from bio_device_shiftlogs  , bio_empmas where bioemp_company = '" & mcode & "' and bioemp_fpcode = ds_fpcode and ds_fpcode = " & lst_employee.ItemData(lst_employee.ListIndex) & "  and ds_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' order by ds_date "
    End If
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    Dim rdate As Date
    Dim dt As String
    While Not payrs.EOF()
            Set itmx = lst_view.ListItems.ADD(, , CStr(payrs("ds_fpcode")))
            itmx.SubItems(1) = Format(payrs.Fields("ds_date"), "dd/MM/yyyy")
            If IsNull(payrs.Fields("ds_shift_in")) = True Then
               itmx.SubItems(2) = ""
            Else
               itmx.SubItems(2) = Format(payrs.Fields("ds_shift_in"), "HH:MM:SS")
            End If
            If IsNull(payrs.Fields("ds_shift_out")) = True Then
               itmx.SubItems(3) = ""
            Else
''               itmx.SubItems(3) = payrs.Fields("ds_shift_out")
               itmx.SubItems(3) = Format(payrs.Fields("ds_shift_out"), "HH:MM:SS")
            End If
            itmx.SubItems(4) = payrs.Fields("ds_status")
            payrs.MoveNext
    Wend
    
''    While Not payrs.EOF()
''          For i = 1 To mdays
''              dt = monthname(cmb_month.ItemData(cmb_month.ListIndex)) + " " + Trim(Str(i)) + " , " + Trim(cmb_year.Text)
''              rdate = CDate(dt)
''              dayfind = "a_day" & i
''              dayfind_intime = "a_in_day" & i
''              dayfind_outtime = "a_out_day" & i
''              If IsNull(payrs.Fields(dayfind_intime)) = False And IsNull(payrs.Fields(dayfind_outtime)) = False Then
''                  If payrs.Fields(dayfind) <> "" Then
''
''                     Set itmx = lst_view.ListItems.ADD(, , CStr(payrs("a_fpcode")))
''                     itmx.SubItems(1) = Format(rdate, "dd/MM/yyyy")
''                     If payrs.Fields(dayfind_intime) = "01/01/1900" Then
''                        itmx.SubItems(2) = ""
''                     Else
''                        itmx.SubItems(2) = Format(payrs.Fields(dayfind_intime), "HH:MM:SS")
''                     End If
''                     If payrs.Fields(dayfind_intime) = "01/01/1900" Then
''                        itmx.SubItems(3) = ""
''                     Else
''                        itmx.SubItems(3) = payrs.Fields(dayfind_outtime)
''                     End If
''                     itmx.SubItems(4) = payrs.Fields(dayfind)
''                  End If
''              End If
''          Next
''          payrs.MoveNext
''    Wend
    payrs.Close
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
''    lst_view.ColumnHeaders.ADD , , "FP Code "
''    lst_view.ColumnHeaders.ADD , , "In Time "
''    lst_view.ColumnHeaders.ADD , , "Out Time "
''    lst_view.ColumnHeaders.ADD , , "Status"
''    lst_view.View = lvwReport
''    dt_from.MaxDate = Format(Now, "dd/mm/yyyy")
''    dt_to.MaxDate = Format(Now, "dd/mm/yyyy")
''    dt_to.Value = Format(Now, "dd/mm/yyyy")
''    dt_from.Value = Format(Now - Day(Now) + 1, "dd/mm/yyyy")
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
    With cmb_year
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
      .AddItem Left(fyear, 4)
      .AddItem Mid(fyear, 6, 4)
      If year(DATE) = Int(Left(fyear, 4)) Then
         cmb_year.Text = Left(fyear, 4)
      Else
          cmb_year.Text = Mid(fyear, 6, 4)
      End If

    End With
''    cmb_year.Text = "2015"
    cmb_month.ListIndex = Month(DATE) - 1

    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
''    sql = "Select * from  pdept_mas order by dept_name"
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    While Not payrs.EOF()
''        lst_dept.AddItem payrs("dept_name")
''        lst_dept.ItemData(lst_dept.NewIndex) = payrs("dept_code")
''        payrs.MoveNext
''    Wend
''    payrs.Close
    sql = "select bioemp_dept  from bio_empmas group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    
    emp_type = "S"
End Sub

Private Sub lst_dept_Click()
    Dim payrs As New ADODB.Recordset
    lst_employee.Clear
''    If opt_all.Value = True Then
''       sql = "Select * from  emp_mas where emp_status = 'A'  and emp_cat = '" & emp_type & "' and emp_workplace  = 'MILL' and emp_dept =  " & lst_dept.ItemData(lst_dept.ListIndex)
''    Else
''       sql = "Select * from  emp_mas where emp_company = " & mcode & " and  emp_status = 'A'  and emp_cat = '" & emp_type & "'  and emp_workplace  = 'MILL' and emp_dept =  " & lst_dept.ItemData(lst_dept.ListIndex)
''    End If
    If opt_all.Value = True Then
       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    Else
       sql = "Select * from  bio_empmas where bioemp_company = '" & mcode & "' and bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    End If
    
    
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub opt_all_Click()
    mcode = 0
End Sub

Private Sub opt_cogen_Click()
   mcode = "COGEN"
End Sub

Private Sub opt_dpm1_Click()
   mcode = "DPM"
End Sub

Private Sub opt_dpm2_Click()
  mcode = "VJPM"
End Sub

Private Sub opt_dpm3_Click()
  mcode = 2
End Sub

Private Sub opt_staff_Click()
   emp_type = "S"
   sw_click
End Sub
Public Sub sw_click()
    Dim payrs As New ADODB.Recordset
    lst_employee.Clear
    If opt_all.Value = True Then
       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    Else
       sql = "Select * from  bio_empmas where bioemp_company = '" & mcode & "' and bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub opt_vjpm_Click()
   mcode = "VJPM"

End Sub

Private Sub opt_worker_Click()
   emp_type = "W"
   sw_click
End Sub

