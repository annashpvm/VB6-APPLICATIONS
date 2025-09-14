VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_manual_inout_punch 
   Caption         =   "MANUAL IN/OUT PUNCH"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9720
      TabIndex        =   55
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame9 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   52
      Top             =   9120
      Width           =   2895
      Begin VB.CommandButton cmd_ok 
         Caption         =   "OK"
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
         Left            =   2280
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txt_pw 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "$"
         TabIndex        =   53
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   9015
      Left            =   11520
      TabIndex        =   50
      Top             =   480
      Width           =   8535
      Begin MSFlexGridLib.MSFlexGrid flx_data_summary 
         Height          =   8295
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   14631
         _Version        =   393216
         RowHeightMin    =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Height          =   375
      Left            =   6960
      TabIndex        =   45
      Top             =   960
      Visible         =   0   'False
      Width           =   885
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
         Left            =   2760
         TabIndex        =   48
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
         Left            =   1560
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   9240
      Visible         =   0   'False
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
         Format          =   130940929
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
         Format          =   130940929
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
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   720
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5280
      TabIndex        =   20
      Top             =   9120
      Width           =   1815
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_manual_inout_punch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_manual_inout_punch.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   7935
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   5895
         Begin VB.OptionButton opt_out_punch 
            Caption         =   "OUT Punch"
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
            Left            =   2160
            TabIndex        =   49
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_in_punch 
            Caption         =   "IN Punch"
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
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt_double 
            Caption         =   "Double Punch"
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
            Left            =   3840
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dt_date_in 
            Height          =   375
            Left            =   2640
            TabIndex        =   40
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
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
            CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
            Format          =   130940931
            CurrentDate     =   42294
         End
         Begin MSComCtl2.DTPicker dt_date_out 
            Height          =   375
            Left            =   2640
            TabIndex        =   41
            Top             =   1440
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
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
            CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
            Format          =   130940931
            CurrentDate     =   42294
         End
         Begin VB.Label lbl_out 
            Caption         =   "Manual out Punch"
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
            Left            =   120
            TabIndex        =   43
            Top             =   1560
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lbl_in 
            Caption         =   "Manual In Punch"
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
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.TextBox txt_idcode 
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
         Height          =   405
         Left            =   480
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete "
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
         Left            =   6360
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "Modify"
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
         Height          =   375
         Left            =   6360
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign"
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
         Left            =   6360
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txt_empname2 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   15
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txt_fpcode 
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
         Height          =   405
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_dept 
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
         Height          =   405
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   3375
         Left            =   240
         TabIndex        =   16
         Top             =   3840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5953
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
         Left            =   1920
         TabIndex        =   19
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
         Left            =   600
         TabIndex        =   18
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
         Left            =   5280
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   6120
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   2595
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   2
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   10320
      TabIndex        =   44
      Top             =   9720
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Manual IN/OUT Entries - for Employees"
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frm_manual_inout_punch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim del_inout As Integer
Private Sub cmd_Assign_Click()
    If opt_in_punch.Value = False And opt_out_punch.Value = False And opt_double.Value = False Then
       MsgBox ("Select IN or OUT punch option...")
       Exit Sub
    End If
    
    txt_fpcode.Text = txt_empcode.Text
    txt_idcode.Text = txt_empcode.Text
    If txt_fpcode.Text = "" Then
       MsgBox ("Select Employee Code / Name.. Please UPLOAD USERS from ESSL")
       
       Exit Sub
    End If

paydb.BeginTrans
On Error GoTo err_handler

    Dim sql As String

    pst_qry = "select max(bio_logid)+1 as endno from bio_manual_logs"
    payrs.Open pst_qry, paydb, 1, 2
    no = 1
    If Not IsNull(payrs!endno) Then
        If Not payrs.EOF Then
             no = payrs!endno
        End If
    End If
    payrs.Close

    pst_qry = "update bio_manual_logs set bio_logid = " & no
    paydb.Execute pst_qry
    
    If opt_in_punch.Value = True Then
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_idcode.Text & ",'" & no & "', '" & Format(dt_date_in.Value, "MM/dd/yyyy") & "', '" & Format(dt_date_in.Value, "MM/dd/yyyy HH:MM:SS") & "','M','N','in')"
        paydb.Execute sql
    ElseIf opt_out_punch.Value = True Then
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_idcode.Text & ",'" & no & "', '" & Format(dt_date_out.Value, "MM/dd/yyyy") & "', '" & Format(dt_date_out.Value, "MM/dd/yyyy HH:MM:SS") & "','M','N','out')"
        paydb.Execute sql
    ElseIf opt_double.Value = True Then
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_idcode.Text & ",'" & no & "', '" & Format(dt_date_in.Value, "MM/dd/yyyy") & "', '" & Format(dt_date_in.Value, "MM/dd/yyyy HH:MM:SS") & "','M','N','in')"
        paydb.Execute sql
        no = no + 1
        sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_idcode.Text & ",'" & no & "', '" & Format(dt_date_out.Value, "MM/dd/yyyy") & "', '" & Format(dt_date_out.Value, "MM/dd/yyyy HH:MM:SS") & "','M','N','out')"
        paydb.Execute sql
        pst_qry = "update bio_manual_logs set bio_logid = " & no
        paydb.Execute pst_qry
    End If
    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
  
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub cmd_clear_Click()
    Dim i As Long
    For i = 0 To lst_employee.ListCount - 1
        lst_employee.Selected(i) = False
    Next
End Sub

Private Sub cmd_delete_Click()
    del_inout = 1
    cmd_Assign.Enabled = False
    cmd_modify.Enabled = True
    flx_data.Enabled = True
End Sub

Private Sub cmd_filter_Click()
    Refresh_Click
    Dim payrs As New ADODB.Recordset
    
    
    
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "'  and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name like  '" & txt_empname.Text & "' and bioemp_dept = '" & lst_dept.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
       sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_dept = '" & lst_dept.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    Else
       MsgBox ("Employee not selected...")
       Exit Sub
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
          txt_fpcode.Text = payrs!bioemp_fpcode
          txt_idcode.Text = payrs!bioemp_id
          txt_empname2.Text = payrs!bioemp_name
          txt_dept.Text = payrs!bioemp_dept
          payrs.MoveNext
    Wend
    payrs.Close
    
    If txt_empcode.Text = "" And txt_fpcode.Text <> "" Then
      txt_empcode.Text = txt_fpcode.Text
    End If
    pst_qry = "select * from bio_devicelogs where ad_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ad_fpcode = '" & txt_empcode.Text & "' order by ad_date,ad_logdate"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = Format(payrs!ad_date, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 2) = Format(payrs!ad_logdate, "dd/MM/yyyy HH:MM:SS")
        flx_data.TextMatrix(i, 3) = payrs!ad_auto
        flx_data.TextMatrix(i, 4) = payrs!ad_logslno
        flx_data.TextMatrix(i, 5) = payrs!ad_punch
        flx_data.Rows = flx_data.Rows + 1
        flx_dataold.TextMatrix(i, 0) = i
        flx_dataold.TextMatrix(i, 1) = Format(payrs!ad_date, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 2) = Format(payrs!ad_logdate, "dd/MM/yyyy HH:MM:SS")
        flx_dataold.TextMatrix(i, 3) = payrs!ad_auto
        flx_dataold.TextMatrix(i, 4) = payrs!ad_logslno
        flx_dataold.TextMatrix(i, 5) = payrs!ad_punch
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        payrs.MoveNext
        i = i + 1
    Wend
    payrs.Close

    pst_qry = "select * from bio_device_shiftlogs where ds_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = '" & txt_empcode.Text & "' order by ds_date"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data_summary.TextMatrix(i, 0) = i
        flx_data_summary.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
        flx_data_summary.TextMatrix(i, 2) = payrs!ds_sft_hrs
        flx_data_summary.TextMatrix(i, 3) = payrs!ds_status
        flx_data_summary.TextMatrix(i, 4) = Format(payrs!ds_shift_in, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 5) = Format(payrs!ds_shift_out, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 6) = Format(payrs!ds_shift_in2, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 7) = Format(payrs!ds_shift_out2, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 8) = Format(payrs!ds_shift_in3, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 9) = Format(payrs!ds_shift_out3, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 10) = Format(payrs!ds_shift_in4, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 11) = Format(payrs!ds_shift_out4, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 12) = Format(payrs!ds_shift_in5, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 13) = Format(payrs!ds_shift_out5, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 14) = Format(payrs!ds_shift_in6, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.TextMatrix(i, 15) = Format(payrs!ds_shift_out6, "dd/MM/yyyy HH:MM:SS")
        flx_data_summary.Rows = flx_data_summary.Rows + 1

        
        payrs.MoveNext
        i = i + 1
    Wend
    payrs.Close
    

End Sub

Private Sub cmd_modify_Click()
Dim i As Integer
paydb.BeginTrans
On Error GoTo err_handler
    Dim sdate, inout_time As Date
    For i = 1 To flx_dataold.Rows - 1
        If Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
            If flx_dataold.TextMatrix(i, 3) = "M" Then
               sdate = Format(flx_dataold.TextMatrix(i, 1), "dd/MM/yyyy")
               inout_time = Format(flx_dataold.TextMatrix(i, 2), "dd/MM/yyyy HH:MM:SS")
               sql = "delete from bio_devicelogs where ad_fpcode =  " & txt_fpcode.Text & "  and ad_date = '" & Format(sdate, "yyyy/MM/dd") & "' and ad_logdate =  '" & Format(inout_time, "yyyy/MM/dd HH:MM:SS") & "' and ad_auto = 'M'"
               paydb.Execute sql
            End If
        End If
    Next

''    pst_qry = "select max(bio_logid)+1 as endno from bio_manual_logs"
''    payrs.Open pst_qry, paydb, 1, 2
''    no = 1
''    If Not IsNull(payrs!endno) Then
''        If Not payrs.EOF Then
''             no = payrs!endno
''        End If
''    End If
''    payrs.Close


    For i = 1 To flx_data.Rows - 1
        If Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy") <> "" Then
            If flx_data.TextMatrix(i, 3) = "M" Then
               sdate = Format(flx_data.TextMatrix(i, 1), "dd/MM/yyyy")
               inout_time = Format(flx_data.TextMatrix(i, 2), "dd/MM/yyyy HH:MM:SS")
               
               sql = "insert into bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_upd,ad_punch) values (" & txt_fpcode.Text & ", " & txt_idcode.Text & ",'" & flx_data.TextMatrix(i, 4) & "', '" & Format(sdate, "MM/dd/yyyy") & "', '" & Format(inout_time, "MM/dd/yyyy HH:MM:SS") & "','M','N','" & flx_data.TextMatrix(i, 5) & "')"
               paydb.Execute sql

            End If
        End If
    Next



    paydb.CommitTrans
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    
    Refresh_Click
    Exit Sub
Exit Sub
err_handler:
        paydb.RollbackTrans
        chk = gen_Validation(Err.Number, Err.Description)

End Sub

Private Sub cmd_single_punch_Click()

End Sub

Private Sub cmd_ok_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    
    Dim superuser, adminuser   As String
    superuser = ""
    adminuser = ""

    sql = "select * from mas_users  where usr_code =65"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF() Then
           adminuser = payrs.Fields("usr_pwd")
    End If
    payrs.Close
    
    If Trim(adminuser) = Trim(txt_pw.Text) Then
       cmd_Assign.Enabled = True
       cmd_delete.Enabled = True
       cmd_modify.Enabled = True
    Else
       cmd_Assign.Enabled = False
       cmd_delete.Enabled = False
       cmd_modify.Enabled = False
    End If
    
    
End Sub

Private Sub Command1_Click()
    Dim sql As String
    Dim no As Long
    
    pst_qry = "select * from bio_devicelogs where ad_auto = 'M'"
    payrs.Open pst_qry, paydb, 1, 2
    no = 50001
    While Not payrs.EOF
        payrs!ad_logslno = no
        payrs.Update
        no = no + 1
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub dt_date_in_Change()
      dt_date_out.Value = dt_date_in.Value
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_data_DblClick()
   If del_inout = 0 Then Exit Sub
   flex_edit_row = 0
   Dim fin_selrow As Integer
   Dim pst_ans As String
   fin_selrow = flx_data.Row

   timchk = 0
   With flx_data
       If flx_data.TextMatrix(.Row, 3) = "A" Then Exit Sub
       pst_ans = MsgBox("Press YES-to DELETE  NO-to CANCEL", vbYesNo, "Confirmation")
       If pst_ans = 6 Then
               If .Rows < 2 Then
                  MsgBox "No rows to remove"
               Else
                  If Val(flx_data.TextMatrix(.Row, 0)) > 0 Then
                     flx_data.RemoveItem fin_selrow
                  End If
                  .Row = flx_data.Rows - 1
               End If
        End If
   End With
End Sub
Private Sub Form_Load()
    dt_date_in.Value = Now
    dt_date_out.Value = Now
    st_date.Value = Now
    end_date.Value = Now
    
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

    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close

    
    fillgrid
''    If adminpw = 0 Then
       cmd_Assign.Enabled = False
       cmd_delete.Enabled = False
       cmd_modify.Enabled = False
''    End If

End Sub

Private Sub lst_dept_Click()
    lst_employee.Clear
''    sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "'"
    If opt_all.Value = True Then
       sql = "Select * from  bio_empmas where bioemp_status = 'Working'  and bioemp_dept = '" & lst_dept.Text & "' order by bioemp_name"
    ElseIf opt_staff.Value = True Then
       sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'S'  and emp_status = 'A'  order by bioemp_name"
    Else
       sql = "select  * from bio_empmas a, emp_mas b where bioemp_fpcode = emp_fpcode and bioemp_status = 'Working' and bioemp_dept = '" & lst_dept.Text & "' and emp_cat = 'W'  and emp_status = 'A'  order by bioemp_name"
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
     .Cols = 6
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Log Time"
     .TextMatrix(0, 3) = "Log"
     .TextMatrix(0, 4) = "Logid"
     .TextMatrix(0, 5) = "Punch"
     .ColWidth(0) = 500
     .ColWidth(1) = 1500
     .ColWidth(2) = 2000
     .ColWidth(3) = 500
     .ColWidth(4) = 0
     .ColWidth(5) = 1000
     .Redraw = True
   End With

   With flx_dataold
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 6
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Log Time"
     .TextMatrix(0, 3) = "Log"
     .TextMatrix(0, 4) = "Logid"
     .TextMatrix(0, 5) = "Punch"
     .ColWidth(0) = 500
     .ColWidth(1) = 1500
     .ColWidth(2) = 2000
     .ColWidth(3) = 500
     .ColWidth(4) = 0
     .ColWidth(5) = 1000
     .Redraw = True
   End With
   ''Added by Jackuline on 13 Mar 2021
   With flx_data_summary
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 16
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Work Hours"
     .TextMatrix(0, 3) = "Status"
     .TextMatrix(0, 4) = "In Time 1"
     .TextMatrix(0, 5) = "Out Time 1"
     .TextMatrix(0, 6) = "In Time 2"
     .TextMatrix(0, 7) = "Out Time 2"
     .TextMatrix(0, 8) = "In Time 3"
     .TextMatrix(0, 9) = "Out Time 3"
     .TextMatrix(0, 10) = "In Time 4"
     .TextMatrix(0, 11) = "Out Time 4"
     .TextMatrix(0, 12) = "In Time 5"
     .TextMatrix(0, 13) = "Out Time 5"
     .TextMatrix(0, 14) = "In Time 6"
     .TextMatrix(0, 15) = "Out Time 6"
     .ColWidth(0) = 500
     .ColWidth(1) = 1300
     .ColWidth(2) = 1600
     .ColWidth(3) = 1000
     .ColWidth(4) = 3000
     .ColWidth(5) = 3000
     .ColWidth(6) = 3000
     .ColWidth(7) = 3000
     .ColWidth(8) = 3000
     .ColWidth(9) = 3000
     .ColWidth(10) = 3000
     .ColWidth(11) = 3000
     .ColWidth(12) = 3000
     .ColWidth(13) = 3000
     .ColWidth(14) = 3000
     .ColWidth(15) = 3000

     .Redraw = True
   End With
End Sub

Private Sub opt_all_Click()
    lst_dept_Click
End Sub

Private Sub opt_double_Click()
    lbl_in.Visible = True
    lbl_out.Visible = True
    dt_date_in.Visible = True
    dt_date_out.Visible = True
    
End Sub



Private Sub opt_in_punch_Click()
    lbl_in.Visible = True
    lbl_out.Visible = False
    dt_date_in.Visible = True
    dt_date_out.Visible = False
End Sub

Private Sub opt_out_punch_Click()
    lbl_in.Visible = False
    lbl_out.Visible = True
    dt_date_in.Visible = False
    dt_date_out.Visible = True
End Sub

Private Sub opt_staff_Click()
    lst_dept_Click
End Sub

Private Sub opt_worker_Click()
   lst_dept_Click
End Sub

Private Sub Refresh_Click()
    i = 0
    del_inout = 0
''    If adminpw = 0 Then
''       cmd_Assign.Enabled = False
''       cmd_delete.Enabled = False
''       cmd_modify.Enabled = False
''    Else
''        cmd_Assign.Enabled = True
''        cmd_modify.Enabled = False
''
''    End If
    
    txt_fpcode.Text = ""
    txt_empname2.Text = ""
    txt_dept.Text = ""
    fillgrid
''    flx_data.Enabled = False
End Sub
Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
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


Private Sub Text1_Change()

End Sub

