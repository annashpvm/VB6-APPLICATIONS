VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_shift_schdule_modification 
   Caption         =   "SHIFT SHEDULE MODIFICATION"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   3720
      TabIndex        =   15
      Top             =   720
      Width           =   6495
      Begin VB.CommandButton cmd_go 
         Caption         =   "GO"
         Height          =   735
         Left            =   5280
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Shift Change range"
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
         Height          =   1095
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   4815
         Begin MSComCtl2.DTPicker st_date 
            Height          =   375
            Left            =   960
            TabIndex        =   25
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   109248513
            CurrentDate     =   39359
         End
         Begin MSComCtl2.DTPicker end_date 
            Height          =   375
            Left            =   2880
            TabIndex        =   26
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   109248513
            CurrentDate     =   39359
         End
         Begin VB.Label Label9 
            Caption         =   "To Date"
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
            Left            =   2760
            TabIndex        =   28
            Top             =   360
            Width           =   1095
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
            Left            =   480
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmb_shift 
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox txt_empname2 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txt_fpcode 
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_dept 
         Height          =   285
         Left            =   4560
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   2415
         Left            =   360
         TabIndex        =   29
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         _Version        =   393216
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Index           =   6
         Left            =   4680
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Shift Changed as"
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
         Left            =   600
         TabIndex        =   19
         Top             =   5160
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4080
      TabIndex        =   11
      Top             =   6960
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule_modification.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_shift_schdule_modification.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   480
      TabIndex        =   0
      Top             =   720
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
      Caption         =   "Shift Modification - for Employees"
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
      TabIndex        =   14
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frm_shift_schdule_modification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim findrow As Integer


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
          payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub cmd_go_Click()
    fillgrid
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset
    
    pst_qry = "select * from bio_shift_schedule where emps_fpcode = '" & txt_fpcode.Text & "' and emps_date between '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' order by emps_date"
    payrs.Open pst_qry, paydb, 1, 2
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = payrs!emps_date
        flx_data.TextMatrix(i, 2) = payrs!emps_shift
        payrs.MoveNext
        flx_data.Rows = flx_data.Rows + 1
        i = i + 1
    Wend
    payrs.Close


End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    st_date.Value = Now
    end_date.Value = Now
    fillgrid
    cmb_shift.AddItem "GS"
    cmb_shift.AddItem "WO"
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
    cmb_shift.AddItem "WO"
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

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lst_dept_Click()
    lst_employee.Clear
    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_employee.AddItem payrs("bioemp_name")
        lst_employee.ItemData(lst_employee.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
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
    
    Dim rdate, sdate, edate As Date
    Dim i As Integer
    
    For i = 1 To flx_data.Rows - 1
        rdate = Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy")
        sql = "update bio_shift_schedule set emps_shift = '" & cmb_shift.Text & "' where emps_fpcode = " & txt_fpcode.Text & " and emps_date = '" & Format(rdate, "MM/dd/yyyy") & "'"
        paydb.Execute sql
    Next
    
    
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
  
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 3
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Shift"
     .ColWidth(0) = 500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     
     .Redraw = True
   End With
End Sub

