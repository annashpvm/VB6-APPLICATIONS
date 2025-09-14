VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_Attn_Corrections 
   Caption         =   "ATTENDANCE CORRECTION"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15435
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_holiday_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   65
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txt_holiday 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   64
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txt_absent_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   59
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txt_absentdays 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   58
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txt_wodays 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   57
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txt_totdays 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   56
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txt_eldays 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   55
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txt_salarydays 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17160
      TabIndex        =   54
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txt_salarydays_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   52
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txt_eldays_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   50
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txt_totdays_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   48
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txt_wodays_act 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   46
      Top             =   5880
      Width           =   1095
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
      Left            =   16560
      TabIndex        =   45
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox lst_employee 
         Height          =   2595
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   6120
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   3240
      TabIndex        =   23
      Top             =   1320
      Width           =   7935
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
         TabIndex        =   27
         Top             =   480
         Width           =   1815
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
         TabIndex        =   26
         Top             =   480
         Width           =   1215
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
         TabIndex        =   25
         Top             =   480
         Width           =   3135
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
         Left            =   6960
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   5295
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9340
         _Version        =   393216
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2880
      TabIndex        =   20
      Top             =   7680
      Width           =   1815
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_Attn_Corrections.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_Attn_Corrections.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   480
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   3015
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   9000
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   129957889
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   129957889
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   1935
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
      Left            =   6840
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   885
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
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
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
         TabIndex        =   8
         Top             =   240
         Width           =   1065
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
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   11400
      TabIndex        =   4
      Top             =   1440
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid flx_data_summary 
         Height          =   2655
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4683
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
      Left            =   12960
      TabIndex        =   1
      Top             =   8760
      Width           =   2895
      Begin VB.TextBox txt_pw 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "#"
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
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
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   5880
      TabIndex        =   43
      Top             =   9240
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Holiday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   63
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "CHANGED AS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   16920
      TabIndex        =   62
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Caption         =   "PRESENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   15360
      TabIndex        =   61
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Absent Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   60
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Salary Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   53
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Eligible Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   51
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Total Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   49
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "WO Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   47
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ATTENDANCE CORRECTIONS FOR WEEKLY OFF"
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
      TabIndex        =   44
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frm_Attn_Corrections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim del_inout As Integer
Dim mmon As Integer
Dim mdays As Integer
Private Sub cmd_clear_Click()
    Dim i As Long
    For i = 0 To lst_employee.ListCount - 1
        lst_employee.Selected(i) = False
    Next
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
   
''    pst_qry = "select * from bio_device_shiftlogs where ds_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = '" & txt_empcode.Text & "' order by ds_date"
''
''    pst_qry = "select * from bio_attendlogs where a_year = " & cmb_year.Text & " and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = '" & txt_empcode.Text & "'"
''
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''
''        For mon = 1 To 31
''        dayfind = "a_day" & mon
''        dayfind2 = "a_in_day" & mon
''
''        If (payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "H") Then
''
''
''        MsgBox (payrs.Fields(dayfind2))
''                MsgBox (payrs.Fields("a_in_day2"))
''            flx_data_summary.TextMatrix(i, 0) = i
''            flx_data_summary.TextMatrix(i, 1) = Format(payrs.Fields(dayfind2), "dd/MM/yyyy")
''            flx_data_summary.TextMatrix(i, 2) = payrs.Fields(dayfind)
''            flx_data_summary.Rows = flx_data_summary.Rows + 1
''            i = i + 1
''        End If
''        Next
''        payrs.MoveNext
''
''    Wend
''    payrs.Close
''

''    pst_qry = "select * from bio_device_shiftlogs where ds_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = '" & txt_empcode.Text & "' and ds_status in ('WO','H') order by ds_date"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''        flx_data_summary.TextMatrix(i, 0) = i
''        flx_data_summary.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
''        flx_data_summary.TextMatrix(i, 2) = payrs!ds_status
''        flx_data_summary.TextMatrix(i, 3) = payrs!ds_status
''        flx_data_summary.Rows = flx_data_summary.Rows + 1
''
''
''        payrs.MoveNext
''        i = i + 1
''    Wend
''    payrs.Close
''    pst_qry = "select * from bio_device_shiftlogs where ds_date  between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ds_fpcode = '" & txt_empcode.Text & "' and ds_status in ('WO','H') order by ds_date"
''    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
''    i = 1
''    While Not payrs.EOF
''        flx_data_summary.TextMatrix(i, 0) = i
''        flx_data_summary.TextMatrix(i, 1) = Format(payrs!ds_date, "dd/MM/yyyy")
''        flx_data_summary.TextMatrix(i, 2) = payrs!ds_status
''        flx_data_summary.TextMatrix(i, 3) = payrs!ds_status
''        flx_data_summary.Rows = flx_data_summary.Rows + 1
''
''
''        payrs.MoveNext
''        i = i + 1
''    Wend
''    payrs.Close
    
    
    
    Dim attndate As Date
''    pst_qry = "select * from bio_attendlogs where a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_year = " & cmb_year.Text & " and a_fpcode = " & txt_empcode.Text & ""
    pst_qry = "select * from bio_attendlogs where a_month = " & mmon & " and a_year = " & cmb_year.Text & " and a_fpcode = " & txt_empcode.Text & ""
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        For k = 1 To mdays
            dayfind = "a_day" & k
            If payrs.Fields(dayfind) = "A" Or payrs.Fields(dayfind) = "H" Or payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "" Or payrs.Fields(dayfind) = "WOH" Or payrs.Fields(dayfind) = "WO½P" Then
''                attndate = DateValue(Str(Month(st_date)) + "/" + Str(k) + "/" + cmb_year.Text)
                attdate = Str(k) + "/" + LTrim(RTrim(Str(mmon))) + "/" + cmb_year.Text
                attndate = DateValue(Str(k) + "/" + LTrim(RTrim(Str(mmon))) + "/" + cmb_year.Text)
                

                flx_data_summary.TextMatrix(i, 0) = i
                flx_data_summary.TextMatrix(i, 1) = Format(attndate, "dd/MM/yyyy")
                flx_data_summary.TextMatrix(i, 2) = payrs.Fields(dayfind)
                flx_data_summary.TextMatrix(i, 3) = payrs.Fields(dayfind)
                flx_data_summary.Rows = flx_data_summary.Rows + 1
                i = i + 1
            End If
        Next
        txt_holiday_act.Text = payrs!a_holiday
        txt_wodays_act.Text = payrs!a_wo
        txt_eldays_act.Text = payrs!a_eligible_days
        txt_totdays_act.Text = payrs!a_total_days
        txt_salarydays_act.Text = payrs!a_salary_days
        txt_absentdays.Text = payrs!a_absent
        txt_wodays.Text = payrs!a_wo
        txt_absent_act.Text = payrs!a_absent
        txt_eldays.Text = payrs!a_eligible_days
        txt_totdays.Text = payrs!a_total_days
        txt_salarydays.Text = payrs!a_salary_days
        payrs.MoveNext
    Wend
    payrs.Close
    

End Sub

Private Sub cmd_modify_Click()
Dim i As Integer
paydb.BeginTrans
On Error GoTo err_handler
    Dim dt As Date
    Dim dday As Integer
    For i = 1 To flx_data_summary.Rows - 1
        If flx_data_summary.TextMatrix(i, 2) <> flx_data_summary.TextMatrix(i, 3) Then
        
          dt = Format(flx_data_summary.TextMatrix(i, 1), "dd/MM/yyyy")
          aday = Trim(Str(Day(dt)))
          pst_qry = "update bio_attendlogs set a_day" & aday & " = '" & Trim(flx_data_summary.TextMatrix(i, 3)) & "' where a_fpcode = " & txt_empcode.Text & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & mmon
          paydb.Execute pst_qry
         
        End If
    Next
    
    
        
    pst_qry = "update bio_attendlogs set a_wo = '" & txt_wodays.Text & "' , a_absent = '" & txt_absentdays.Text & "' , a_eligible_days = '" & txt_eldays.Text & "' , a_total_days = '" & txt_totdays.Text & "'  , a_salary_days = '" & txt_salarydays.Text & "'  Where a_fpcode = " & txt_empcode.Text & " And a_year = " & Val(cmb_year.Text) & " And a_month = " & mmon
    paydb.Execute pst_qry
          
    paydb.CommitTrans
    
    MsgBox "Record Modified", vbOKOnly + vbInformation, "Information"
    fillgrid
        txt_wodays_act.Text = ""
        txt_eldays_act.Text = ""
        txt_totdays_act.Text = ""
        txt_salarydays_act.Text = ""
        txt_absentdays.Text = ""
        txt_wodays.Text = ""
        txt_absent_act.Text = ""
        txt_eldays.Text = ""
        txt_totdays.Text = ""
        txt_salarydays.Text = ""
''    Refresh_Click
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

    sql = "select * from mas_users  where usr_code =66"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF() Then
           adminuser = payrs.Fields("usr_pwd")
    End If
    payrs.Close
    
    If Trim(adminuser) = Trim(txt_pw.Text) Then

       cmd_modify.Enabled = True
    Else

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

Private Sub flx_data_summary_Click()
    fin_selrow = flx_data_summary.Row
    findatacol = flx_data_summary.Col

    Select Case findatacol
       Case 3
          If flx_data_summary.TextMatrix(flx_data_summary.Row, 2) = "WO" Or flx_data_summary.TextMatrix(flx_data_summary.Row, 2) = "A" Or flx_data_summary.TextMatrix(flx_data_summary.Row, 2) = "" Or flx_data_summary.TextMatrix(flx_data_summary.Row, 2) = "WOH" Then
             If flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "WO" Or flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "WOH" Or flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "WO½P" Then
                flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "A"
             Else
                flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "WO"
             End If
          End If
          


           If flx_data_summary.TextMatrix(flx_data_summary.Row, 2) = "H" Then
             If flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "H" Then
                flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "A"
             Else
                flx_data_summary.TextMatrix(flx_data_summary.Row, 3) = "H"
             End If
          End If
          Dim cnt, diffdays, abscnt, h As Integer
          cnt = 0
          abscnt = 0
          h = 0
          For i = 1 To flx_data_summary.Rows - 1
              If flx_data_summary.TextMatrix(i, 3) = "WO" Then
                 cnt = cnt + 1
              End If
              If flx_data_summary.TextMatrix(i, 3) = "WO½P" Then
                 cnt = cnt + 0.5
              End If
              If flx_data_summary.TextMatrix(i, 3) = "A" Then
                 abscnt = abscnt + 1
              End If
              If flx_data_summary.TextMatrix(i, 3) = "H" Or flx_data_summary.TextMatrix(i, 3) = "HP" Then
                 h = h + 1
              End If
          Next
          txt_wodays.Text = cnt
          txt_holiday.Text = h
          diffdays = Val(txt_wodays_act.Text) - Val(txt_wodays.Text) + Val(txt_holiday_act.Text) - Val(txt_holiday.Text)
          txt_absentdays.Text = abscnt
          
          txt_eldays.Text = Val(txt_eldays_act.Text) - diffdays
          txt_totdays.Text = Val(txt_eldays_act.Text) - diffdays
          txt_salarydays.Text = Val(txt_eldays_act.Text) - diffdays
          
    End Select
    Exit Sub

End Sub

Private Sub Form_Load()

    
    st_date.Value = Now
    end_date.Value = Now
    
    Dim monvalue As Integer
    
    
    With cmb_month
        If Month(Now) = 1 Then
            .AddItem "December"
            .AddItem "January"
        End If
        If Month(Now) = 2 Then
            .AddItem "January"
            .AddItem "February"
        End If
        If Month(Now) = 3 Then
            .AddItem "February"
            .AddItem "March"
        End If
        If Month(Now) = 4 Then
            .AddItem "March"
            .AddItem "April"
        End If
        If Month(Now) = 5 Then
            .AddItem "April"
            .AddItem "May"
        End If
        If Month(Now) = 6 Then
            .AddItem "May"
            .AddItem "June"
        End If
        
        If Month(Now) = 7 Then
            .AddItem "June"
            .AddItem "July"
        End If
        If Month(Now) = 8 Then
            .AddItem "July"
            .AddItem "August"
        End If
        If Month(Now) = 9 Then
            .AddItem "August"
            .AddItem "September"
        End If
        If Month(Now) = 10 Then
            .AddItem "September"
            .AddItem "October"
        End If
        If Month(Now) = 11 Then
            .AddItem "October"
            .AddItem "November"
        End If
        If Month(Now) = 12 Then
            .AddItem "November"
            .AddItem "December"
        End If
        
        
    End With
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
''    cmb_month.ListIndex = Month(Date) - 1
    
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
     .Cols = 4
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Date"
     .TextMatrix(0, 2) = "Status"
     .TextMatrix(0, 3) = "Change"
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     
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
   '' mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If cmb_month.Text = "April" Then mmon = 4
    If cmb_month.Text = "May" Then mmon = 5
    If cmb_month.Text = "June" Then mmon = 6
    If cmb_month.Text = "July" Then mmon = 7
    If cmb_month.Text = "August" Then mmon = 8
    If cmb_month.Text = "September" Then mmon = 9
    If cmb_month.Text = "October" Then mmon = 10
    If cmb_month.Text = "November" Then mmon = 11
    If cmb_month.Text = "December" Then mmon = 12
    If cmb_month.Text = "January" Then mmon = 1
    If cmb_month.Text = "February" Then mmon = 2
    If cmb_month.Text = "March" Then mmon = 3
    
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


