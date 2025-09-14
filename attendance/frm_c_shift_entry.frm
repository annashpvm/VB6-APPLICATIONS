VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_shift_c_entry 
   Caption         =   "C SHIFT ENTRY"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   15435
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   10440
      TabIndex        =   35
      Top             =   720
      Width           =   2175
      Begin VB.CommandButton PROCESS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&VIEW"
         Height          =   825
         Left            =   600
         Picture         =   "frm_c_shift_entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1680
         Width           =   945
      End
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   120848385
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   120848385
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
         Left            =   480
         TabIndex        =   39
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      TabIndex        =   26
      Top             =   7560
      Width           =   2175
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_c_shift_entry.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_c_shift_entry.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   6855
      Begin VB.Frame Frame6 
         Caption         =   "EMPLOYEE SEARCH"
         Height          =   2535
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   6615
         Begin VB.CommandButton cmd_show 
            Caption         =   "SHOW"
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
            Left            =   5160
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_name 
            Height          =   405
            Left            =   2280
            TabIndex        =   31
            Top             =   240
            Width           =   2655
         End
         Begin MSFlexGridLib.MSFlexGrid flx_empdata 
            Height          =   1575
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
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
            Caption         =   "Enter Emp.Short Name"
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
            Index           =   7
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.CommandButton cmd_add 
         Caption         =   "ADD"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   6615
         Begin VB.TextBox txt_empname2 
            Height          =   285
            Left            =   1560
            TabIndex        =   20
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txt_fpcode 
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txt_dept 
            Height          =   285
            Left            =   4440
            TabIndex        =   18
            Top             =   480
            Width           =   1815
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
            Left            =   1560
            TabIndex        =   23
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
            Left            =   240
            TabIndex        =   22
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
            Left            =   4560
            TabIndex        =   21
            Top             =   240
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
         Left            =   2400
         TabIndex        =   15
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Caption         =   "Shift Date"
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
         Height          =   615
         Left            =   720
         TabIndex        =   12
         Top             =   120
         Width           =   4815
         Begin MSComCtl2.DTPicker r_date 
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   120848385
            CurrentDate     =   39359
         End
         Begin VB.Label Label10 
            Caption         =   "Shift Date"
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
            Left            =   600
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   2175
         Left            =   240
         TabIndex        =   24
         Top             =   4920
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
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
      Begin VB.Label Label1 
         Caption         =   "Shift"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   4560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox lst_dept 
         Height          =   2010
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox lst_employee 
         Height          =   2400
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   2655
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "REFRESH"
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
         TabIndex        =   2
         Top             =   6600
         Width           =   1215
      End
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
         Left            =   1560
         TabIndex        =   1
         Top             =   6600
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   3600
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_blink 
      Caption         =   " <----- To Modify Click the grid and change the shift and give modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   975
      Left            =   10560
      TabIndex        =   41
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Shift C && 12 Hrs Entry - for Employees"
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
      Left            =   960
      TabIndex        =   29
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frm_shift_c_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pst_respo As String
Dim flx_edit_row As Integer
Dim fin_selrow As String

Private Sub cmd_add_Click()
        
Dim pst_ans As String
On Error GoTo err_handler

If Trim(txt_fpcode.Text) = "" Then
    MsgBox "Select Employee .", vbOKOnly + vbExclamation, "vbInformation"
    lst_employee.SetFocus
    Exit Sub
End If

''If cmb_shift.Text = "" Then
''    MsgBox "Select Employee Shift ", vbOKOnly + vbExclamation, "vbInformation"
''    cmb_shift.SetFocus
''    Exit Sub
''End If
 
 With flx_data
 If flx_edit_row = 0 Then
        
        .TextMatrix(.Rows - 1, 1) = txt_fpcode.Text
        .TextMatrix(.Rows - 1, 2) = txt_empname2.Text
        .TextMatrix(.Rows - 1, 3) = cmb_shift.Text
        .TextMatrix(.Rows - 1, 4) = txt_dept.Text
        .Rows = .Rows + 1
    Else
        .TextMatrix(flx_edit_row, 1) = txt_fpcode.Text
        .TextMatrix(flx_edit_row, 2) = txt_empname2.Text
        .TextMatrix(flx_edit_row, 3) = cmb_shift.Text
        .TextMatrix(flx_edit_row, 4) = txt_dept.Text
        flx_edit_row = 0
    End If

End With
For i = 1 To flx_data.Rows - 1
   flx_data.TextMatrix(i, 0) = i
Next



cmd_add.Caption = "ADD"

Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
        
        
        
End Sub

Private Sub cmd_filter_Click()
     
    If txt_empcode.Text = "" And txt_empname.Text = "" And lst_employee.Text = "" Then
        MsgBox ("Select Employee Name / Code...")
        Exit Sub
    End If
 
    Dim payrs As New ADODB.Recordset
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    
    ElseIf txt_empname.Text <> "" Then
       sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
       sql = "select * from bio_empmas where bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          
          payrs.MoveNext
    Wend
    payrs.Close
    
     
    fpcode = txt_fpcode.Text
    fillgrid
    
    If fpcode = "" Then Exit Sub
    pst_qry = "select * from bio_shift_schedule where emps_fpcode = " & fpcode & " and emps_date = '" & Format(r_date, "MM/dd/yyyy") & "'"
    
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        
        flx_data.TextMatrix(i, 0) = 1
        flx_data.TextMatrix(i, 1) = txt_fpcode.Text
        flx_data.TextMatrix(i, 2) = txt_empname2.Text
        flx_data.TextMatrix(i, 3) = payrs!emps_shift
        flx_data.TextMatrix(i, 4) = txt_dept.Text
        
        
        payrs.MoveNext
        flx_data.Rows = flx_data.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close
''  flx_data.Enabled = False

    
End Sub

Private Sub cmd_show_Click()
    fillgrid2
    Dim payrs As New ADODB.Recordset
    Dim i As Integer
    i = 1
    sql = "select * from bio_empmas where bioemp_status = 'Working' and bioemp_name like '%" + Trim(txt_name.Text) + "%'"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        With flx_empdata
        .TextMatrix(.Rows - 1, 0) = i
        .TextMatrix(.Rows - 1, 1) = payrs!bioemp_fpcode
        .TextMatrix(.Rows - 1, 2) = payrs!bioemp_name
        .TextMatrix(.Rows - 1, 3) = payrs!bioemp_dept
        .Rows = .Rows + 1
        i = i + 1
        End With
        
        payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_data_DblClick()
fin_selrow = flx_data.Row
pst_ans = MsgBox("Press YES-to Modify NO-to Cancel", vbYesNo, "Confirmation")
    If pst_ans = vbYes Then
        
        cmd_add.Caption = "Modify"
        
        
        txt_fpcode.Text = flx_data.TextMatrix(fin_selrow, 1)
        txt_empname2.Text = flx_data.TextMatrix(fin_selrow, 2)
        txt_dept.Text = flx_data.TextMatrix(fin_selrow, 4)
        cmb_shift.Text = flx_data.TextMatrix(fin_selrow, 3)
        flx_edit_row = fin_selrow
    ElseIf pst_ans = vbNo Then
        If flx_data.Rows <= 2 Then
            flx_data.Rows = 1
            MsgBox "No rows to remove"
        Else
           flx_data.RemoveItem fin_selrow
           flx_data.Row = flx_data.Rows - 1
        End If
    End If

End Sub

Private Sub flx_empdata_DblClick()
        fin_selrow = flx_empdata.Row
        txt_fpcode.Text = flx_empdata.TextMatrix(fin_selrow, 1)
        txt_empname2.Text = flx_empdata.TextMatrix(fin_selrow, 2)
        txt_dept.Text = flx_empdata.TextMatrix(fin_selrow, 3)
End Sub

Private Sub Form_Load()
    r_date.Value = Now
    st_date.Value = Now
    end_date.Value = Now
    fillgrid
    fillgrid2
''    cmb_shift.AddItem "A SHIFT"
''    cmb_shift.AddItem "B SHIFT"
    cmb_shift.AddItem ""
    cmb_shift.AddItem "GS"
    cmb_shift.AddItem "C SHIFT"
    cmb_shift.AddItem "6.00AM-6.00PM"
    cmb_shift.AddItem "6.00PM-6.00AM"
    cmb_shift.AddItem "A+6.00PM-6.00AM"
    cmb_shift.AddItem "8.00 PM to 8.00 AM"
    cmb_shift.AddItem "B+C"
    cmb_shift.AddItem "A+B+C"
    cmb_shift.AddItem "A+B+C(Partical)"
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


Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 5
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Emp code"
     .TextMatrix(0, 2) = "Emp Name"
     .TextMatrix(0, 3) = "Shift"
     .TextMatrix(0, 4) = "Department"
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 2200
     .ColWidth(3) = 2000
     .ColWidth(4) = 2000
     
     .Redraw = True
   End With
   
   
End Sub

Private Sub fillgrid2()
   
   With flx_empdata
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 4
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Emp code"
     .TextMatrix(0, 2) = "Emp Name"
     .TextMatrix(0, 3) = "Department"
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 2200
     .ColWidth(3) = 2000
     
     .Redraw = True
   End With
   
   
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

Private Sub lst_employee_Click()
    Dim payrs As New ADODB.Recordset
    sql = "select * from bio_empmas a, emp_mas b where a.bioemp_fpcode = b.emp_fpcode and bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "'  and bioemp_dept = '" & lst_dept.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          
          payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub PROCESS_Click()
   Dim date1 As Date
   date1 = 1 & "/" & 1 & " /" & 1900
   
   MousePointer = vbDefault
   gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
   cry_rep1.Formulas(0) = "sdate = '" & Format(st_date.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(1) = "edate = '" & Format(end_date.Value, "dd/mm/yyyy") & "'"
   cry_rep1.Formulas(2) = ("millname= '" & mname & "'")
   
   cry_rep1.PrinterSelect
   Dim sw, ds, emp, mill As String
   
    
   

   ds = " and  ({bio_device_shiftlogs.emps_shift} =  '6.00 to 6.00 (Night)' or   {bio_device_shiftlogs.emps_shift} =  '6.00PM-6.00AM' or {bio_device_shiftlogs.emps_shift} = '8.00 PM to 8.00 AM' or  {bio_device_shiftlogs.emps_shift} = 'B+C' or {bio_device_shiftlogs.emps_shift} = 'A+B+C' or {bio_device_shiftlogs.emps_shift} = 'Unshift_neight' or     {bio_device_shiftlogs.emps_shift} = 'G+C' or {bio_device_shiftlogs.emps_shift} = 'A+C' or {bio_device_shiftlogs.emps_shift} = 'G1/2+C' or {bio_device_shiftlogs.emps_shift} = 'C SHIFT' or {bio_device_shiftlogs.emps_shift} = 'G+NEXTDAY'  )"
   cry_rep1.ReportFileName = "\\10.0.0.252\vbcryrep\payroll\attendance_shift_details.rpt"
   cry_rep1.ReplaceSelectionFormula ("{bio_device_shiftlogs.emps_date} >=  date(" & Format$(st_date, "yyyy,mm,dd") & ") and {bio_device_shiftlogs.emps_date} <=  date(" & Format$(end_date, "yyyy,mm,dd") & ")  " & ds & " ")
    
   
   cry_rep1.WindowState = crptMaximized
   cry_rep1.Connect = gst_repconnect
   cry_rep1.Action = 1

End Sub

Private Sub save_Click()
   If flx_data.Rows - 1 = 0 Then Exit Sub
   
''   If Format(r_date.Value, "MM/dd/yyyy") < Format(Now - 2, "MM/dd/yyyy") Then
''         MsgBox ("You can't enter data  for this day ....")
''         Exit Sub
''   End If
   
   If txt_fpcode.Text = "" Then
         MsgBox ("Select Employee Name")
         Exit Sub
   End If
 

paydb.BeginTrans
On Error GoTo err_handler
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset

    Dim i As Integer
    
    For i = 1 To flx_data.Rows - 1
        If Val(flx_data.TextMatrix(i, 1)) > 0 Then
           pst_qry = "delete from bio_shift_schedule where emps_fpcode = '" & Val(flx_data.TextMatrix(i, 1)) & "'  and emps_date = '" & Format(r_date.Value, "MM/dd/yyyy") & "'"
           paydb.Execute pst_qry
           sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date,emps_shift_alloted, emps_shift) values ( " & Val(flx_data.TextMatrix(i, 1)) & ", '" & Format(r_date.Value, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 3) & "','" & flx_data.TextMatrix(i, 3) & "')"
''           sql = "insert into bio_shift_schedule ( emps_fpcode, emps_date, emps_shift) values ( " & Val(flx_data.TextMatrix(i, 1)) & ", '" & Format(r_date.Value, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 3) & "')"
           paydb.Execute sql
        End If
    Next
       
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
  
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    fillgrid
    fillgrid2
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

