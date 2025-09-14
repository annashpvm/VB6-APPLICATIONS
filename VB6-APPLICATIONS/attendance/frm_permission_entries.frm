VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_permission_entries 
   Caption         =   "PERMISSION ENTRY"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_pw 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "#"
      TabIndex        =   50
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmd_r 
      Caption         =   "R"
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   6600
      Width           =   135
   End
   Begin VB.Frame Frame9 
      Caption         =   "Permission Type"
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
      Height          =   735
      Left            =   6960
      TabIndex        =   46
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton opt_personnal 
         Caption         =   "Personnal"
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
         TabIndex        =   48
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt_mill 
         Caption         =   "Mill"
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
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   360
      TabIndex        =   39
      Top             =   6960
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   126025729
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   126025729
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   480
      TabIndex        =   34
      Top             =   840
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4440
      TabIndex        =   21
      Top             =   7680
      Width           =   2175
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   240
         MaskColor       =   &H000000FF&
         Picture         =   "frm_permission_entries.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_permission_entries.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   9375
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
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
         Left            =   5640
         TabIndex        =   45
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
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
         Left            =   5640
         TabIndex        =   44
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "&Assign "
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
         Left            =   5640
         TabIndex        =   33
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   5295
         Begin VB.TextBox txt_purpose 
            Height          =   405
            Left            =   1800
            TabIndex        =   30
            Top             =   1800
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   375
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   126025729
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dt_fromtime 
            Height          =   375
            Left            =   1800
            TabIndex        =   24
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm:ss"
            Format          =   126025731
            CurrentDate     =   41387.375
         End
         Begin MSComCtl2.DTPicker dt_totime 
            Height          =   375
            Left            =   1800
            TabIndex        =   25
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm:ss"
            Format          =   126025731
            CurrentDate     =   41387.75
         End
         Begin VB.Label Label2 
            Caption         =   "Purpose"
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
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "From time"
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
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "To time"
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
            TabIndex        =   26
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   "Permission Date"
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
            TabIndex        =   17
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.TextBox txt_empname2 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txt_fpcode 
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txt_dept 
         Height          =   285
         Left            =   4560
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   2055
         Left            =   240
         TabIndex        =   28
         Top             =   3720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3625
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
         TabIndex        =   20
         Top             =   360
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
         TabIndex        =   19
         Top             =   360
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
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   1440
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
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
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
         Top             =   3120
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   6840
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Permission Entries - for Employees"
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
      TabIndex        =   23
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frm_permission_entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim del_permission As Integer
Dim docno As Integer
Dim pertype As String

Private Sub cmb_month_Change()
   find_dates
End Sub

Private Sub cmb_month_Click()
  find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
End Sub

Private Sub cmd_Assign_Click()
    Dim min1, min2 As Integer
    Dim permin As Integer
    
     min1 = (DatePart("h", dt_fromtime.Value) * 60) + DatePart("n", dt_fromtime.Value)
     min2 = (DatePart("h", dt_totime.Value) * 60) + DatePart("n", dt_totime.Value)
     
''     If min2 - min1 > 120 Then
     permins = 0
     If opt_personnal.Value = True Then
        permins = 120
     Else
        permins = 240
     End If
     If min2 - min1 > permins Then
         MsgBox ("Permission hours should not exceed " & permins & " Minutes. ")
         Exit Sub
     End If
      
     
     If min1 > min2 Then
         MsgBox ("Error in FROM and TO times ")
         Exit Sub
     End If
     
     
''    min1 = (Int(Val(txt_total_hrs.Text)) * 60) + ((Val(txt_total_hrs.Text) - Int(Val(txt_total_hrs.Text)))) * 100
  ''  min2 = (Int(Val(txt_runhrs.Text)) * 60) + ((Val(txt_runhrs.Text) - Int(Val(txt_runhrs.Text)))) * 100
   
   
   If txt_purpose.Text = "" Then
         MsgBox ("Enter data in Purpose ")
         txt_purpose.SetFocus
         Exit Sub
   End If
   



   If txt_fpcode.Text = "" Then
         MsgBox ("Select Employee Code ")
         txt_fpcode.SetFocus
         Exit Sub
   End If


''   If Format(dt_from.Value, "MM/dd/yyyy") < Format(Now - 5, "MM/dd/yyyy") Then
''      If Trim(txt_pw.Text) <> "ATTN" Then
''         MsgBox ("You can't give permission for this day / period....")
''         Exit Sub
''      End If
''   End If
   
    
    Dim payrs As New ADODB.Recordset

    Dim pst_qry As String
    
    permin = min2 - min1
    
    Dim smin, emin As Double
    smin = 0
    emin = 0
    If opt_personnal.Value = True Then
        pst_qry = "select * from bio_emp_permissions where emp_per_type = 'P' and  empp_fpcode = " & txt_fpcode.Text & " and empp_date >= '" & Format(st_date.Value, "MM/dd/yyyy") & "' and empp_date <= '" & Format(end_date.Value, "MM/dd/yyyy") & "'"
        payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs.EOF
              ''permin = permin + Int(payrs("empp_fromtime")) * 60 + (payrs("empp_fromtime") - Int(payrs("empp_fromtime"))) * 100
              smin = smin + Int(payrs("empp_fromtime")) * 60 + (payrs("empp_fromtime") - Int(payrs("empp_fromtime"))) * 100
              emin = emin + Int(payrs("empp_totime")) * 60 + (payrs("empp_totime") - Int(payrs("empp_totime"))) * 100
              payrs.MoveNext
        Wend
        payrs.Close
        permin = permin + emin - smin
        If permin > 120 Then
           MsgBox ("Already utilized this employee his monthly eligible permission time")
           Exit Sub
        End If
    End If
   
    Dim idate As Date
    pst_qry = "select * from bio_emp_permissions where empp_fpcode = " & txt_fpcode.Text & " and empp_date = '" & Format(dt_from.Value, "MM/dd/yyyy") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
          MsgBox ("Already  Permission assigned to " + txt_empname2.Text + " on Date :" + Format(dt_from.Value, "dd/MM/yyyy"))
          payrs.Close
          Exit Sub
    End If
    payrs.Close


paydb.BeginTrans
On Error GoTo err_handler
''    If opt_leave_half.Value = True Then
''       dt_to.Value = dt_from.Value
''    End If
    
    pst_qry = "select max(empp_no)+1 as endno from bio_emp_permissions "
    
    payrs.Open pst_qry, paydb, 1, 2
    no = 1
    If Not IsNull(payrs!endno) Then
        If Not payrs.EOF Then
             no = payrs!endno
        End If
    End If
    payrs.Close
    Dim sql As String

    
    sql = "insert into bio_emp_permissions (empp_no,empp_fpcode,empp_date,emp_per_type,empp_fromtime,empp_totime,empp_purpose) values (" & no & ", " & txt_fpcode.Text & ", '" & Format(dt_from.Value, "MM/dd/yyyy") & "' ,'" & pertype & "'," & Format(dt_fromtime, "HH.MM") & "," & Format(dt_totime, "HH.MM") & ",'" & txt_purpose.Text & "')"
    paydb.Execute sql
    
    
    
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    
    txt_pw.Visible = False
    txt_pw.Text = ""
  
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub cmd_delete_Click()
    flx_data.Enabled = True
    del_permission = 1
    cmd_Assign.Enabled = False
    cmd_modify.Enabled = True
End Sub

Private Sub cmd_filter_Click()
    fillgrid
    Dim payrs As New ADODB.Recordset
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name like  '" & txt_empname.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' and bioemp_dept =  '" & lst_dept.Text & "' order by bioemp_dept"
    Else
       MsgBox ("Employee Not selected..")
       Exit Sub
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          payrs.MoveNext
    Wend
    payrs.Close
    
''    pst_qry = "select * from bio_emp_permissions  a ,emp_mas b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_empcode.Text & "' and  empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    
''    If opt_vou.Value = True Then
''       pst_qry = "select * from bio_emp_permissions  a ,emp_voupay_mast b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    ElseIf opt_casuals.Value = True Then
''       pst_qry = "select * from bio_emp_permissions  a ,mas_caemp  b  where  a.empp_fpcode = b.ca_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    Else
''       pst_qry = "select * from bio_emp_permissions  a ,emp_mas b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
''    End If
    
''    If opt_vou.Value = True Then
''       pst_qry = "select * from bio_emp_permissions  a ,emp_voupay_mast b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date =  '" & Format(dt_from, "MM/dd/yyyy") & "'"
''    ElseIf opt_casuals.Value = True Then
''       pst_qry = "select * from bio_emp_permissions  a ,mas_caemp  b  where  a.empp_fpcode = b.ca_fpcode and a.empp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date ='" & Format(dt_from, "MM/dd/yyyy") & "'"
''    Else
''       pst_qry = "select * from bio_emp_permissions  a ,emp_mas b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date = '" & Format(dt_from, "MM/dd/yyyy") & "'"
''    End If
      pst_qry = "select * from bio_emp_permissions  a ,emp_mas b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date = '" & Format(dt_from, "MM/dd/yyyy") & "'"
    
    pst_qry = "select * from bio_emp_permissions  a ,emp_mas b  where  a.empp_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & txt_fpcode.Text & "' and  empp_date >= '" & Format(st_date, "MM/dd/yyyy") & "'  and  empp_date <= '" & Format(end_date, "MM/dd/yyyy") & "'"
    
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = Format(payrs!empp_date, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 2) = payrs!empp_fromtime
        flx_data.TextMatrix(i, 3) = payrs!empp_totime
        flx_data.TextMatrix(i, 4) = payrs!empp_purpose
        flx_data.TextMatrix(i, 5) = payrs!empp_no
        flx_data.TextMatrix(i, 6) = payrs!emp_per_type
        flx_data.Rows = flx_data.Rows + 1
        
        flx_dataold.TextMatrix(i, 0) = i
        flx_dataold.TextMatrix(i, 1) = Format(payrs!empp_date, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 2) = payrs!empp_fromtime
        flx_dataold.TextMatrix(i, 3) = payrs!empp_totime
        flx_dataold.TextMatrix(i, 4) = payrs!empp_purpose
        flx_dataold.TextMatrix(i, 5) = payrs!empp_no
        payrs.MoveNext
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close
End Sub

Private Sub cmd_modify_Click()
Dim i As Integer
paydb.BeginTrans
On Error GoTo err_handler
    
    For i = 1 To flx_dataold.Rows - 1
        Dim sdate As Date
        If Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
            sdate = Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy")
            sql = "delete from bio_emp_permissions where empp_fpcode = " & txt_fpcode.Text & "  and empp_date  = '" & Format(sdate, "MM/dd/yyyy") & "'"
            paydb.Execute sql
        End If
    Next


    For i = 1 To flx_data.Rows - 1
        If Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
           sdate = Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy")
           If flx_data.TextMatrix(i, 3) <> "" Then
              sql = "insert into bio_emp_permissions (empp_no,empp_fpcode,empp_date,empp_fromtime,empp_totime,empp_purpose) values (" & Val(flx_data.TextMatrix(i, 5)) & ", " & txt_fpcode.Text & ", '" & Format(sdate, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 2) & "','" & flx_data.TextMatrix(i, 3) & "','" & flx_data.TextMatrix(i, 4) & "')"
              paydb.Execute sql
           End If
        End If
    Next


 ''   sql = "insert into bio_emp_permissions (empp_no,empp_fpcode,empp_date,empp_fromtime,empp_totime,empp_purpose) values (" & docno & ", " & txt_fpcode.Text & ", '" & Format(dt_from.Value, "MM/dd/yyyy") & "'," & Format(dt_fromtime, "HH.MM") & "," & Format(dt_totime, "HH.MM") & ",'" & txt_purpose.Text & "')"
   '' paydb.Execute sql

    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    
    txt_pw.Visible = False
    txt_pw.Text = ""
    Refresh_Click
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub cmd_r_Click()
    txt_pw.Visible = True
    txt_pw.Text = ""
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
pertype = "P"
dt_from.Value = Now
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
  
    Dim payrs As New ADODB.Recordset
    lst_dept.Clear
    fillgrid
    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close
    Refresh_Click
End Sub

Private Sub lst_dept_Click()
  lst_employee.Clear
     
    If opt_personnal.Value = True Then
       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'    order by bioemp_name"
    Else
       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'  and bioemp_team  ='STAFF'   order by bioemp_name"
    End If
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
     .Cols = 7
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "From time"
     .TextMatrix(0, 3) = "To time"
     .TextMatrix(0, 4) = "Purpose"
     .TextMatrix(0, 5) = "Docno"
     .TextMatrix(0, 6) = "type"
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1500
     .ColWidth(5) = 1000
     .ColWidth(6) = 1000
     .Redraw = True
   End With
   With flx_dataold
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 6
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "From time"
     .TextMatrix(0, 3) = "To time"
     .TextMatrix(0, 4) = "Purpose"
     .TextMatrix(0, 5) = "Docno"
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1500
     .ColWidth(5) = 1000
     .Redraw = True
   End With

End Sub


Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
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

Private Sub opt_mill_Click()
    pertype = "M"
End Sub

Private Sub opt_personnal_Click()
    pertype = "P"
End Sub

Private Sub Refresh_Click()
    del_permission = 0
    fillgrid
    cmd_Assign.Enabled = True
    cmd_modify.Enabled = False
    flx_data.Enabled = False
    txt_pw.Visible = False
    txt_pw.Text = ""
End Sub
Private Sub flx_data_DblClick()
   If del_permission = 0 Then Exit Sub
   flex_edit_row = 0
   Dim fin_selrow As Integer
   Dim pst_ans As String
   fin_selrow = flx_data.Row

   timchk = 0
   With flx_data
       pst_ans = MsgBox("Press YES-to DELETE  NO-to CANCEL", vbYesNo, "Confirmation")
       If pst_ans = 6 Then
               If .Rows < 2 Then
                  MsgBox "No rows to remove"
               Else
                  If Val(flx_data.TextMatrix(.Row, 0)) > 0 Then
                     docno = Val(flx_data.TextMatrix(.Row, 5))
                     flx_data.RemoveItem fin_selrow
                  End If
                  .Row = flx_data.Rows - 1
               End If
        End If
   End With
End Sub

