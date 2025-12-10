VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_shift_schdule_random_new 
   Caption         =   "Shift Schedule "
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   16275
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10080
      TabIndex        =   37
      Top             =   5640
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule_random_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule_random_new.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   0
      TabIndex        =   32
      Top             =   7320
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130809857
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   34
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130809857
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   3360
      TabIndex        =   17
      Top             =   480
      Width           =   6615
      Begin VB.ComboBox cmb_weekoff 
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
         ItemData        =   "frm_shift_schdule_random_new.frx":0AAC
         Left            =   3120
         List            =   "frm_shift_schdule_random_new.frx":0AAE
         TabIndex        =   26
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_dept 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txt_fpcode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt_empname2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   6255
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
            TabIndex        =   20
            Top             =   120
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
            Left            =   840
            TabIndex        =   19
            Top             =   120
            Width           =   3015
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
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   855
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
            TabIndex        =   21
            Top             =   240
            Width           =   555
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   4935
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8705
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Select Weekly off"
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
         Left            =   1080
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
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
         TabIndex        =   30
         Top             =   120
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
         TabIndex        =   29
         Top             =   120
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
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmb_shift 
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
      Left            =   10440
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   10680
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   10080
      TabIndex        =   1
      Top             =   720
      Width           =   2415
      Begin VB.OptionButton opt_sw 
         Caption         =   "SW"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opt_c 
         Caption         =   "C"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_wo_change 
      Caption         =   "Change WO"
      Height          =   495
      Left            =   10080
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   40
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu mnu_shift 
      Caption         =   "shift"
      Visible         =   0   'False
      Begin VB.Menu mnu_blank 
         Caption         =   ""
      End
      Begin VB.Menu mnu_gs 
         Caption         =   "GS"
      End
      Begin VB.Menu mnu_a 
         Caption         =   "A SHIFT"
      End
      Begin VB.Menu mnu_b 
         Caption         =   "B SHIFT"
      End
      Begin VB.Menu mnu_c 
         Caption         =   "C SHIFT"
      End
      Begin VB.Menu mnuAB 
         Caption         =   "A+ B"
      End
      Begin VB.Menu mnu_abc 
         Caption         =   "A+B+C"
      End
      Begin VB.Menu mnu_ab_cpartical 
         Caption         =   "A+B+C(Partical)"
      End
      Begin VB.Menu mnu_ac 
         Caption         =   "A+C"
      End
      Begin VB.Menu mnu_a_AND_6pm_6am 
         Caption         =   "A+6.00PM-6.00AM"
      End
      Begin VB.Menu mnu_6to6_day 
         Caption         =   "6.00AM-6.00PM"
      End
      Begin VB.Menu mnu_6to6_night 
         Caption         =   "6.00PM-6.00AM"
      End
      Begin VB.Menu mnu_7pm_7am 
         Caption         =   "7.00PM-7.00AM"
      End
      Begin VB.Menu mnu_8pm_8am 
         Caption         =   "8.00 PM to 8.00 AM"
      End
      Begin VB.Menu mnu_bc 
         Caption         =   "B+C"
      End
      Begin VB.Menu mnu_unshift 
         Caption         =   "Unshift_neight"
      End
      Begin VB.Menu mnu_gc 
         Caption         =   "G+C"
      End
      Begin VB.Menu mnu_ghalf_c 
         Caption         =   "G1/2+C"
      End
      Begin VB.Menu mnu_Unshift_2dayworked 
         Caption         =   "Unshift_2daysworked"
      End
      Begin VB.Menu mnu_6to6_day_Cshift 
         Caption         =   "6.00AM-6.00PM+ C SHIFT"
      End
      Begin VB.Menu mnu_gnextday 
         Caption         =   "G+NEXTDAY"
      End
   End
End
Attribute VB_Name = "frm_shift_schdule_random_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim findrow As Integer
Dim fin_selrow As Integer


Private Sub cmb_month_Change()
   find_dates
End Sub

Private Sub cmb_month_Click()
   find_dates
End Sub

Public Sub flex_refresh()
    fillgrid
    Dim rdate As Date
  
    For rdate = st_date To end_date
         findrow = flx_data.Rows - 1
        If UCase(Format(rdate, "dddd")) = cmb_weekoff.Text Then
           flx_data.TextMatrix(findrow, 2) = "WO"
        End If
        flx_data.TextMatrix(findrow, 0) = findrow
        flx_data.TextMatrix(findrow, 1) = rdate
        flx_data.Rows = flx_data.Rows + 1
    Next
End Sub

Private Sub cmb_shift_Click()
On Error GoTo err_handler
    If cmb_shift.Text = "" Then Exit Sub
    fin_selrow = flx_data.Row
    findatacol = flx_data.Col
    If flx_data.Rows > 2 Then
       flx_data.TextMatrix(fin_selrow, 2) = cmb_shift.Text
       cmb_shift.Visible = False
       Me.MousePointer = 1
    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
    Me.MousePointer = 1

End Sub

Private Sub cmb_shift_GotFocus()
On Error GoTo err_handler
Dim pin_ret As Integer
    pin_ret = FindItem(cmb_shift.hwnd, CB_FINDSTRING, -1, flx_data.TextMatrix(fin_selrow, 2))
    If pin_ret >= 0 Then
        cmb_shift.ListIndex = pin_ret
    Else
        cmb_shift.ListIndex = -1
    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then
        Resume
    End If
End Sub

Private Sub cmb_shift_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    If KeyAscii = 27 Then
        cmb_shift.Visible = False
        Exit Sub
    ElseIf KeyAscii = 13 Then
        'cmb_unit_LostFocus
        Dim pin_ret As Integer
        pin_ret = FindItem(cmb_shift.hwnd, CB_FINDSTRING, -1, cmb_shift.Text)
        If pin_ret >= 0 Then
            cmb_shift.ListIndex = pin_ret
        Else
            cmb_unit.ListIndex = -1
            Exit Sub
        End If
    '    put_unit_grid
        If cmb_shift.ListIndex = -1 Then
          MsgBox "Select Shift ", vbInformation, "Error"
          cmb_shift.SetFocus
          Exit Sub
        End If
        flx_data.TextMatrix(fin_selrow, 2) = cmb_shift.Text
        cmb_shift.Visible = False
    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then
        Resume
    End If

End Sub

Private Sub cmb_weekoff_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmb_year_Change()
    find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
End Sub

Private Sub cmd_filter_Click()
    fillgrid
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    Dim payrs As New ADODB.Recordset
    If opt_sw.Value = True Then
        If txt_empcode.Text <> "" Then
          sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        ElseIf txt_empname.Text <> "" Then
           sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        ElseIf lst_employee.Text <> "" Then
           sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        Else
           MsgBox ("Employee Not selected...")
           Exit Sub
        End If
    Else
        If txt_empcode.Text <> "" Then
          sql = "select * from bio_empmas a, mas_caemp b where a.bioemp_fpcode = b.ca_fpcode and bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        ElseIf txt_empname.Text <> "" Then
           sql = "select * from bio_empmas a, mas_caemp b where a.bioemp_fpcode = b.ca_fpcode and bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        ElseIf lst_employee.Text <> "" Then
           sql = "select * from bio_empmas a, mas_caemp b where a.bioemp_fpcode = b.ca_fpcode and bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
        Else
           MsgBox ("Employee Not selected...")
           Exit Sub
        End If
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          If opt_sw.Value = True Then
             cmb_weekoff.Text = payrs!emp_holiday
          Else
          End If
          payrs.MoveNext
    Wend
    payrs.Close
    
    flex_refresh
    
    Dim payrs2 As New ADODB.Recordset

    pst_qry = "select *  from bio_shift_schedule where  emps_fpcode = '" & txt_fpcode.Text & "' and emps_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "'order by emps_date"
    payrs2.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''     If payrs2.EOF Then
''''        MsgBox "SHIFT NOT ALLOCATED FOR THIS EMPLOYEE "
''        MousePointer = vbDefault
''        Exit Sub
''    End If
    
       While Not payrs2.EOF
           For i = 1 To flx_data.Rows - 1
               With flx_data
                    If .TextMatrix(i, 1) = Format(payrs2("emps_date"), "dd/MM/yyyy") Then
                       ''.TextMatrix(i, 2) = payrs2("emps_shift_alloted")
                       .TextMatrix(i, 2) = payrs2("emps_shift")
                    End If
''                    i = i + 1
               End With
           Next
          payrs2.MoveNext
        Wend
   payrs2.Close
End Sub




Private Sub cmd_wo_change_Click()
    For i = 1 To flx_data.Rows - 1
        If flx_data.TextMatrix(i, 2) = "WO" Then
           flx_data.TextMatrix(i, 2) = "GS"
        End If
        If UCase(Format(flx_data.TextMatrix(i, 1), "dddd")) = cmb_weekoff.Text Then
           flx_data.TextMatrix(i, 2) = "WO"
        End If

    Next
End Sub

Private Sub Command1_Click()

MsgBox "Rows " & flx_data.Row & " through " & flx_data.RowSel & " are selected"
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_data_Click()
On Error GoTo err_handler
    '''''''assign row clicked by user to a variable to facilitate easy removal of that row
''    mouse_row = flx_data.Row
   '' flx_data_EnterCell
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then
        Resume
    End If
End Sub

Private Sub flx_data_EnterCell()
''On Error GoTo err_handler
''    If flx_data.Rows < 2 Then Exit Sub
''    flx_data.CellForeColor = vbRed
''
''    fin_selrow = flx_data.Row
''    findatacol = flx_data.Col
''    With flx_data
''    Select Case findatacol
''        Case 2
''            cmb_shift.Visible = True
''            cmb_shift.Left = .Left + .CellLeft
''            cmb_shift.Left = 6000
''            cmb_shift.Top = .Top + .CellTop
''            cmb_shift.Top = 2600
''            cmb_shift.Width = .CellWidth
''            cmb_shift.Text = IIf(flx_data.TextMatrix(fin_selrow, 2) = "", cmb_shift.Text, flx_data.TextMatrix(fin_selrow, 2))
''
''            cmb_shift.Visible = True
''            cmb_shift.SetFocus
''    End Select
''    End With
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''      cmb_shift.Text = flx_data.TextMatrix(1, findatacol)
End Sub

Private Sub flx_data_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo err_handler
    If KeyCode = vbKeyDelete Then
        If flx_data.Rows <> 2 Then
            'flx_details.RemoveItem fin_row
            flx_data.RemoveItem mouse_row
        Else
            flx_data.Rows = 1
            flx_data.Rows = 2
        End If
    End If
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then
        Resume
    End If
    
End Sub

Private Sub flx_data_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
    End If
End Sub

Private Sub flx_data_MouseDown(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
''Text1.Enabled = False
PopupMenu mnu_shift
''PopupMenu mnu_a
''Text1.Enabled = True
End If
End Sub
Private Sub Form_Load()
    fillgrid
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
''        .AddItem finyear + 2000
''        .AddItem "2012"
''        .AddItem "2013"
''        .AddItem "2014"
''        .AddItem "2015"
''        .AddItem "2016"
''
''        .Text = "2015"
      .AddItem Left(fyear, 4)
      .AddItem Mid(fyear, 6, 4)
      If Year(Date) = Int(Left(fyear, 4)) Then
         cmb_year.Text = Left(fyear, 4)
      Else
          cmb_year.Text = Mid(fyear, 6, 4)
      End If

    End With
    cmb_month.ListIndex = Month(Date) - 1
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
    
    cmb_shift.AddItem "8.00 PM to 8.00 AM"
    cmb_shift.AddItem "B+C"
    cmb_shift.AddItem "A+B+C"
    cmb_shift.AddItem "A+6.00PM-6.00AM"
''    cmb_shift.AddItem "WO"
    cmb_shift.AddItem "Unshift_neight"
    cmb_shift.AddItem "G+C"
    cmb_shift.AddItem "A+C"
    cmb_shift.AddItem "A+B"
    cmb_shift.AddItem "G1/2+C"
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear

    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    find_dates
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
     .Cols = 3
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = " Date"
     .TextMatrix(0, 2) = "Shift"
     .ColWidth(0) = 500
     .ColWidth(1) = 1500
     .ColWidth(2) = 2000
     
     .Redraw = True
   End With
End Sub
Public Sub find_dates()
     
    If cmb_month.ListIndex = -1 Then Exit Sub
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
''    mmon = 10
    
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
    flex_refresh
End Sub

Private Sub mnu_gs_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "GS"
Next
End With
End Sub
Private Sub mnu_unshift_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "Unshift_neight"
Next
End With

End Sub
    
Private Sub mnu_ghalf_c_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "G1/2+C"
Next
End With
End Sub
    
    
Private Sub mnu_a_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A SHIFT"
Next
End With
End Sub
Private Sub mnu_wo_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "WO"
Next
End With
End Sub
Private Sub mnu_a_AND_6pm_6am_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A+6.00PM-6.00AM"
Next
End With
End Sub
Private Sub mnu_gnextday_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "G+NEXTDAY"
Next
End With
End Sub

Private Sub mnu_b_Click()
With flx_data
''If flx_data.RowSel = flx_data.Rows Then

For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "B SHIFT"
Next
End With
End Sub
Private Sub mnu_Unshift_2dayworked_click()
With flx_data
''If flx_data.RowSel = flx_data.Rows Then

For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "Unshift_2daysworked"
Next
End With
End Sub

Private Sub mnu_c_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "C SHIFT"
Next
End With
End Sub
Private Sub mnu_6to6_day_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "6.00AM-6.00PM"
Next
End With
End Sub
Private Sub mnu_6to6_night_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "6.00PM-6.00AM"
Next
End With
End Sub


Private Sub mnu_abc_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A+B+C"
Next
End With
End Sub
Private Sub mnu_ab_cpartical_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A+B+C(Partical)"
Next
End With
End Sub

Private Sub mnu_blank_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = ""
Next
End With
End Sub

Private Sub mnu_bc_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "B+C"
Next
End With
End Sub
Private Sub mnu_ac_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A+C"
Next
End With
End Sub
Private Sub mnu_8pm_8am_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "8.00 PM to 8.00 AM"
Next
End With
End Sub

Private Sub mnu_gc_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "G+C"
Next
End With
End Sub

Private Sub mnu_6to6_day_Cshift_click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "6.00AM-6.00PM+ C SHIFT"
Next
End With
End Sub

Private Sub mnuAB_Click()
With flx_data
For i = flx_data.Row To flx_data.RowSel
.TextMatrix(i, 2) = "A+B"
Next
End With
End Sub

Private Sub save_Click()
   If flx_data.Rows - 1 = 1 Then Exit Sub
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
        If Trim(flx_data.TextMatrix(i, 2)) = "" Then
           sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date,emps_shift_alloted, emps_shift) values ( " & txt_fpcode.Text & ", '" & Format(sdate, "MM/dd/yyyy") & "','GS','GS')"
        Else
           sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date,emps_shift_alloted, emps_shift) values ( " & txt_fpcode.Text & ", '" & Format(sdate, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 2) & "','" & flx_data.TextMatrix(i, 2) & "')"
        End If
        paydb.Execute sql
    Next
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    fillgrid
    txt_empcode.Text = ""
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub



