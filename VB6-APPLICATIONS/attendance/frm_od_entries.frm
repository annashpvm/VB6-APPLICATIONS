VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_od_entries 
   Caption         =   "ON DUTY ENTRIES"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_pw 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "#"
      TabIndex        =   56
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmd_r 
      Caption         =   "R"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   5160
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   480
      TabIndex        =   19
      Top             =   600
      Width           =   3015
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmd_filter 
         Caption         =   "FILTER"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4920
         Width           =   975
      End
      Begin VB.ListBox lst_employee 
         Height          =   1425
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ListBox lst_dept 
         Height          =   1425
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   120
         TabIndex        =   20
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   3600
      TabIndex        =   8
      Top             =   600
      Width           =   10575
      Begin VB.Frame Frame6 
         Caption         =   "VIEW OD DETAILS"
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
         Height          =   2415
         Left            =   120
         TabIndex        =   45
         Top             =   3840
         Width           =   6255
         Begin VB.Frame Frame8 
            Height          =   735
            Left            =   0
            TabIndex        =   46
            Top             =   240
            Width           =   6135
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
               TabIndex        =   48
               Top             =   240
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
               TabIndex        =   47
               Top             =   240
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
               TabIndex        =   50
               Top             =   360
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
               TabIndex        =   49
               Top             =   360
               Width           =   555
            End
         End
         Begin MSFlexGridLib.MSFlexGrid flx_data 
            Height          =   1335
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2355
            _Version        =   393216
         End
      End
      Begin VB.TextBox txt_purpose 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   34
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox txt_place 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   32
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   10335
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
            Height          =   735
            Left            =   8640
            TabIndex        =   52
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
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
               Left            =   240
               TabIndex        =   57
               Top             =   720
               Width           =   1335
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
               Left            =   240
               TabIndex        =   54
               Top             =   480
               Width           =   1335
            End
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
               TabIndex        =   53
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "OD Type"
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
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   1695
            Begin VB.OptionButton opt_od_full 
               Caption         =   "Full Day"
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
               TabIndex        =   42
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton opt_od_partial 
               Caption         =   "Partial day"
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
               TabIndex        =   41
               Top             =   720
               Width           =   1335
            End
         End
         Begin MSComCtl2.DTPicker dt_from 
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd/MM/yyyy HH:MM"
            Format          =   130088961
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dt_to 
            Height          =   375
            Left            =   3000
            TabIndex        =   15
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd/MM/yyyy HH:MM"
            Format          =   130088961
            CurrentDate     =   42278
         End
         Begin MSComCtl2.DTPicker dtout 
            Height          =   375
            Left            =   6240
            TabIndex        =   36
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   130088963
            CurrentDate     =   41387.375
         End
         Begin MSComCtl2.DTPicker dtIn 
            Height          =   375
            Left            =   6240
            TabIndex        =   37
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   130088963
            CurrentDate     =   41387.75
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
            Left            =   4920
            TabIndex        =   39
            Top             =   720
            Width           =   1335
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
            Left            =   4920
            TabIndex        =   38
            Top             =   360
            Width           =   1215
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
            Left            =   1920
            TabIndex        =   17
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
            Left            =   1920
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd_Assign 
         Caption         =   "Assign OD"
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
         Left            =   7440
         TabIndex        =   12
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "Modify OD"
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
         Height          =   735
         Left            =   7440
         TabIndex        =   11
         Top             =   5520
         Width           =   2175
      End
      Begin VB.CommandButton cmd_view_leaves 
         Caption         =   "View OD details"
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
         TabIndex        =   10
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete OD"
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
         Left            =   7440
         TabIndex        =   9
         Top             =   4680
         Width           =   2175
      End
      Begin MSComctlLib.ListView lst_view 
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Index           =   5
         Left            =   6120
         TabIndex        =   35
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Place of OD"
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
         Left            =   6120
         TabIndex        =   33
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   7200
      Width           =   2175
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1080
         MaskColor       =   &H000000FF&
         Picture         =   "frm_od_entries.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_od_entries.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130088961
         CurrentDate     =   39359
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130088961
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_dataold 
      Height          =   1455
      Left            =   6960
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker Dt1 
      Height          =   375
      Left            =   13800
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "HH:MM"
      Format          =   130088962
      CurrentDate     =   41387.3333333333
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   14040
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "HH:MM"
      Format          =   130088962
      CurrentDate     =   41387.7083333333
   End
   Begin MSComCtl2.DTPicker dt_entdate 
      Height          =   375
      Left            =   14160
      TabIndex        =   58
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   130088961
      CurrentDate     =   42278
   End
   Begin VB.Label Label3 
      Caption         =   "ON DUTY Entries - for Employees"
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
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_od_entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fpcode As Integer
Dim no As Integer
Dim rdate As Date
Dim del_leave As Integer
Dim ODCHK As Integer
''Private Sub cmb_leave_KeyPress(KeyAscii As Integer)
''    KeyAscii = 0
''End Sub

Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Click()
   find_dates
End Sub

Private Sub cmd_Assign_Click()
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset

    Dim min1, min2 As Integer
    
     min1 = (DatePart("h", dtout.Value) * 60) + DatePart("n", dtout.Value)
     min2 = (DatePart("h", dtIn.Value) * 60) + DatePart("n", dtIn.Value)
     
    fdate = Day(dt_from.Value)
    sdate = Day(dt_to.Value)
     
     If min1 > min2 And fdate = sdate Then
         MsgBox ("Error in FROM and TO times . Please rectify and Continue.. ")
         Exit Sub
     End If
     

''   If cmb_leave.Text = "" Then Exit Sub
''   If opt_leave_half.Value = True Then
''      If opt_an.Value = False And opt_fn.Value = False Then
''         MsgBox ("Select Leave - FORE NOON / AFTER NOON ")
''         Exit Sub
''      End If
''   End If
''
''
''   If Format(dt_from.Value, "MM/dd/yyyy") < Format(Now - 5, "MM/dd/yyyy") Then
''      If Trim(txt_pw.Text) <> "ATTN" Then
''         MsgBox ("You can't give OD for this day / period....")
''         Exit Sub
''      End If
''   End If
    
    If txt_place.Text = "" Then
        MsgBox ("Enter Place of OD ")
        txt_place.SetFocus
        Exit Sub
    End If
        

    Dim iSelected As Integer
    Dim item As ListItem
    For i = 1 To lst_view.ListItems.Count
        If lst_view.ListItems(i).Checked = True Then
          iSelected = iSelected + 1
        End If
    Next
    If iSelected = 0 Then
       MsgBox ("Employee Not selected in the view...")
       Exit Sub
    End If


    ODCHK = 0
    Dim idate As Date
    For i = 1 To lst_view.ListItems.Count
        If lst_view.ListItems(i).Checked = True Then
           For idate = dt_from To dt_to
               pst_qry = "select * from bio_emp_oddetails  where empod_fpcode = " & lst_view.ListItems(i).Text & " and empod_date = '" & Format(idate, "MM/dd/yyyy") & "'"
               payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
               If Not payrs.EOF Then
''                     MsgBox ("Already ON DUTY assigned for " + lst_view.ListItems.item(i).SubItems(1) + " Date " + Format(idate, "dd/MM/yyyy"))
                     ODCHK = 1
''                     payrs.Close
               End If
               payrs.Close
           Next
        End If
        
        If ODCHK = 1 Then
            pst_ans = MsgBox("Already ON DUTY assigned for " + lst_view.ListItems.item(i).SubItems(1) + " Date " + Format(idate, "dd/MM/yyyy") + " Do You want Continue to Assign OD .. YES-to Continue  NO-to CANCEL", vbYesNo, "Confirmation")
''            MsgBox (pst_ans)
            If pst_ans = 7 Then
               Exit Sub
            End If
                    
        End If
        
    Next
   
paydb.BeginTrans
On Error GoTo err_handler
''    If opt_leave_half.Value = True Then
''       dt_to.Value = dt_from.Value
''    End If
    pst_qry = "select max(empod_no)+1 as endno from bio_emp_oddetails "
    
    payrs.Open pst_qry, paydb, 1, 2
    no = 1
    If Not IsNull(payrs!endno) Then
        If Not payrs.EOF Then
             no = payrs!endno
        End If
    End If
    payrs.Close
    Dim ltype, sql As String
''    If opt_leave_full.Value = True Then
''       ltype = "F"
''    Else
''       If opt_fn.Value = True Then
''          ltype = "1"
''       Else
''          ltype = "2"
''       End If
''    End If
    
    For i = 1 To lst_view.ListItems.Count
        If lst_view.ListItems(i).Checked = True Then
           For idate = dt_from To dt_to
''               sql = "insert into bio_emp_oddetails  (empod_no , empod_fpcode,  empod_date, empod_fromtime,empod_totime,empod_location,empod_purpose) values (" & no & ", " & lst_view.ListItems(i).Text & ",'" & Format(idate, "MM/dd/yyyy") & "','" & Format(dtout.Value, "HH.MM") & "','" & Format(dtIn.Value, "HH.MM") & "','" & txt_place.Text & "','" & txt_purpose.Text & "')"
               sql = "insert into bio_emp_oddetails  (empod_no ,empod_entry_date , empod_fpcode,  empod_date, empod_fromtime,empod_totime,empod_location,empod_purpose) values (" & no & " ,'" & Format(dt_entdate.Value, "MM/dd/yyyy") & "', " & lst_view.ListItems(i).Text & ",'" & Format(idate, "MM/dd/yyyy") & "','" & Format(dtout.Value, "HH.MM") & "','" & Format(dtIn.Value, "HH.MM") & "','" & txt_place.Text & "','" & txt_purpose.Text & "')"
               
               paydb.Execute sql
           Next
        End If
    Next
    paydb.CommitTrans
    ''MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    MsgBox "Record Saved in the Entry Number : " + Str(no) + " Entry Date : " + "'" & Format(dt_entdate.Value, "dd/MM/YYYY") & "'", vbOKOnly + vbInformation, "Information"""
    
    od_details
    txt_pw.Visible = False
    txt_pw.Text = ""
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
    del_leave = 1
    cmd_Assign.Enabled = False
    cmd_modify.Enabled = True
    flx_data.Enabled = True
End Sub

Private Sub cmd_filter_Click()
    Dim chk As Integer
    chk = 0
     Refresh_Click
    fillgrid
    Dim payrs As New ADODB.Recordset
    ''Dim itmX As ListItem
    Dim itmX As MSComctlLib.ListItem
    lst_view.ColumnHeaders.Clear
    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
    lst_view.ColumnHeaders.Add , , "Department ", 1500
    lst_view.View = lvwReport
    lst_view.ListItems.Clear
    
''    If lst_view.SelectedItem Is Nothing Then
''       MsgBox ("Employee code Not found Select Employee...")
''       Exit Sub
''    End If
       
    
    
    If txt_empcode.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_fpcode =  '" & txt_empcode.Text & "' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf txt_empname.Text <> "" Then
      sql = "select * from bio_empmas where bioemp_name like  '%" & txt_empname.Text & "%' and bioemp_status = 'Working' order by bioemp_dept"
    ElseIf lst_employee.Text <> "" Then
          sql = "select * from bio_empmas where bioemp_name =  '" & lst_employee.Text & "' and bioemp_status = 'Working' and bioemp_dept =  '" & lst_dept.Text & "'  order by bioemp_dept"
    Else
         MsgBox ("Employee code Not found ....")
         Exit Sub
    End If
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
            Set itmX = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
            itmX.SubItems(1) = payrs.Fields("bioemp_name")
            itmX.SubItems(2) = payrs.Fields("bioemp_dept")
            If chk = 0 Then
               itmX.Checked = True
            End If
            chk = 1
            payrs.MoveNext
    Wend
    payrs.Close
    
    If opt_vou.Value = True Then
       pst_qry = "select * from bio_emp_oddetails   a ,emp_voupay_mast b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and emp_status = 'A' order by empod_date desc"
    ElseIf opt_cs.Value = True Then
           pst_qry = "select * from bio_emp_oddetails   a ,mas_caemp b  where  a.empod_fpcode = b.ca_fpcode and b.ca_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and ca_status = 'A' order by empod_date desc"
    Else
       pst_qry = "select * from bio_emp_oddetails   a ,emp_mas b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and emp_status = 'A' order by empod_date desc"
    End If
    
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        If opt_cs.Value = True Then
           flx_data.TextMatrix(i, 1) = payrs!ca_empname
        Else
           flx_data.TextMatrix(i, 1) = payrs!emp_name
        End If
        flx_data.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 3) = payrs!empod_fromtime
        flx_data.TextMatrix(i, 4) = payrs!empod_totime
        flx_data.TextMatrix(i, 5) = payrs!empod_location
        flx_data.TextMatrix(i, 6) = payrs!empod_purpose
        flx_data.TextMatrix(i, 7) = payrs!empod_no
        flx_data.Rows = flx_data.Rows + 1
        
        flx_dataold.TextMatrix(i, 0) = i
''        flx_dataold.TextMatrix(i, 1) = payrs!emp_name
        If opt_cs.Value = True Then
           flx_dataold.TextMatrix(i, 1) = payrs!ca_empname
        Else
           flx_dataold.TextMatrix(i, 1) = payrs!emp_name
        End If
        
        flx_dataold.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 3) = payrs!empod_fromtime
        flx_dataold.TextMatrix(i, 4) = payrs!empod_totime
        flx_dataold.TextMatrix(i, 5) = payrs!empod_location
        flx_dataold.TextMatrix(i, 6) = payrs!empod_purpose
        flx_dataold.TextMatrix(i, 7) = payrs!empod_no
        payrs.MoveNext
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close

    

End Sub


Private Sub cmd_modify_Click()
   fpcode = lst_view.SelectedItem.Text
''   If cmb_leave.Text = "" Then Exit Sub
''   If opt_leave_half.Value = True Then
''      If opt_an.Value = False And opt_fn.Value = False Then
''         MsgBox ("Select Leave - FORE NOON / AFTER NOON ")
''         Exit Sub
''      End If
''   End If

 

paydb.BeginTrans
On Error GoTo err_handler
    
    Dim pst_qry As String
    Dim payrs As New ADODB.Recordset
    
    Dim rdate, sdate, edate As Date
    For i = 1 To flx_dataold.Rows - 1
        sdate = Format(flx_dataold.TextMatrix(i, 2), "dd/MM/yyyy")
        sql = "delete from bio_emp_oddetails  where empod_fpcode = " & fpcode & "  and empod_date  = '" & Format(sdate, "MM/dd/yyyy") & "'"
        paydb.Execute sql
    Next
               
               
    For i = 1 To flx_data.Rows - 1
        If Val(flx_data.TextMatrix(i, 0)) > 0 Then
            sdate = Format(flx_data.TextMatrix(i, 2), "dd/MM/yyyy")
            sql = "insert into bio_emp_oddetails  (empod_no , empod_fpcode, empod_date, empod_fromtime, empod_totime, empod_location, empod_purpose) values (" & Val(flx_data.TextMatrix(i, 7)) & ", " & fpcode & ", '" & Format(sdate, "MM/dd/yyyy") & "','" & flx_data.TextMatrix(i, 3) & "','" & flx_data.TextMatrix(i, 4) & "','" & flx_data.TextMatrix(i, 5) & "','" & flx_data.TextMatrix(i, 6) & "')"
            paydb.Execute sql
        End If
    Next
    
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
Private Sub od_details()
    fpcode = lst_view.SelectedItem.Text
    fillgrid
    
''    Dim pin_row, i As Integer
''        i = 0
''        If lst_view.ListCount > 0 Then
''           For pin_row = 0 To lst_view.ListCount - 1
''               If lst_emp.Selected(pin_row) = True Then
''                  If i = 0 Then
''                     emp = " and ( {emp_mas.emp_fpcode} = '" & lst_view.List(pin_row) & "'"
''                     i = i + 1
''                  Else
''                     emp = emp + " or {emp_mas.emp_fpcode} = '" & lst_view.List(pin_row) & "'"
''                  End If
''               End If
''           Next pin_row
''        End If
    Dim payrs As New ADODB.Recordset
    ''pst_qry = "select * from bio_emp_oddetails   a ,emp_mas b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode =  '" & lst_view.SelectedItem.Text & "' and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"''
    
    pst_qry = "select * from bio_emp_oddetails   a ,emp_mas b  where  a.empod_fpcode = b.emp_fpcode and b.emp_fpcode in  ('" & lst_view.SelectedItem.Text & "') and  empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    i = 1
    While Not payrs.EOF
        flx_data.TextMatrix(i, 0) = i
        flx_data.TextMatrix(i, 1) = payrs!emp_name
        flx_data.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
        flx_data.TextMatrix(i, 3) = payrs!empod_fromtime
        flx_data.TextMatrix(i, 4) = payrs!empod_totime
        flx_data.TextMatrix(i, 5) = payrs!empod_location
        flx_data.TextMatrix(i, 6) = payrs!empod_purpose
        flx_data.TextMatrix(i, 7) = payrs!empod_no
        flx_data.Rows = flx_data.Rows + 1
        
        flx_dataold.TextMatrix(i, 0) = i
        flx_dataold.TextMatrix(i, 1) = payrs!emp_name
        flx_dataold.TextMatrix(i, 2) = Format(payrs!empod_date, "dd/MM/yyyy")
        flx_dataold.TextMatrix(i, 3) = payrs!empod_fromtime
        flx_dataold.TextMatrix(i, 4) = payrs!empod_totime
        flx_dataold.TextMatrix(i, 5) = payrs!empod_location
        flx_dataold.TextMatrix(i, 6) = payrs!empod_purpose
        flx_dataold.TextMatrix(i, 7) = payrs!empod_no
        payrs.MoveNext
        flx_dataold.Rows = flx_dataold.Rows + 1
        
        i = i + 1
    Wend
    payrs.Close
''  flx_data.Enabled = False

End Sub

Private Sub cmd_r_Click()
    txt_pw.Visible = True
    txt_pw.Text = ""
End Sub

Private Sub dt_from_Change()
     dt_to.Value = dt_from.Value
     od_details
''     dtpOut.Value = dt_from.Value
End Sub
Private Sub dt_to_Change()
'''     dtpIn.Value = dt_to.Value
If dt_from > dt_to Then
MsgBox " Todate is less than from date"
dt_to.SetFocus
Exit Sub
End If
od_details
End Sub
Private Sub exit_Click()
     Unload Me
End Sub

Private Sub flx_data_DblClick()
   If del_leave = 0 Then Exit Sub
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
                     flx_data.RemoveItem fin_selrow
                  End If
                  .Row = flx_data.Rows - 1
               End If

        End If
   End With

End Sub

Private Sub Form_Load()
    del_leave = 0
    dt_from.Value = Now
    dt_to.Value = Now
    dt_entdate.Value = Now
    opt_od_full.Value = True

   dtout.Value = "09.00 AM"
   dtIn.Value = "06.00 PM"
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

    sql = "select bioemp_dept  from bio_empmas where bioemp_status = 'Working' group by bioemp_dept order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        lst_dept.AddItem payrs("bioemp_dept")
        payrs.MoveNext
    Wend
    payrs.Close

    ''Dim itmX As ListItem
    Dim itmX As MSComctlLib.ListItem
    lst_view.ColumnHeaders.Clear
    lst_view.ColumnHeaders.Add , , "FP Code ", 1000
    lst_view.ColumnHeaders.Add , , "Emp. Name ", 2000
    lst_view.ColumnHeaders.Add , , "Department ", 1500
    lst_view.ColumnHeaders.Add , , "Type ", 1000
    lst_view.View = lvwReport
    lst_view.ListItems.Clear
    
    sql = "select * from bio_empmas where bioemp_status = 'Working' order by bioemp_dept"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
            Set itmX = lst_view.ListItems.Add(, , CStr(payrs("bioemp_fpcode")))
            itmX.SubItems(1) = payrs.Fields("bioemp_name")
            itmX.SubItems(2) = payrs.Fields("bioemp_dept")
            itmX.SubItems(3) = payrs.Fields("bioemp_team")
            payrs.MoveNext
    Wend
    
    payrs.Close
    fillgrid

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
    od_details
End Sub


Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 8
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Name"
     .TextMatrix(0, 2) = "OD Date"
     .TextMatrix(0, 3) = "Time Out"
     .TextMatrix(0, 4) = "Time In"
     .TextMatrix(0, 5) = "Location"
     .TextMatrix(0, 6) = "Purpose"
     .TextMatrix(0, 7) = "Docno"
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 0
     .ColWidth(6) = 0
     .ColWidth(7) = 0
     .Redraw = True
   End With
   With flx_dataold
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 8
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Name"
     .TextMatrix(0, 2) = "OD Date"
     .TextMatrix(0, 3) = "Time Out"
     .TextMatrix(0, 4) = "Time In"
     .TextMatrix(0, 5) = "Location"
     .TextMatrix(0, 6) = "Purpose"
     .TextMatrix(0, 7) = "Docno"
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 0
     .ColWidth(6) = 0
     .ColWidth(7) = 0
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

Private Sub lst_view_click()
od_details

End Sub

Private Sub opt_od_full_Click()
''dtpOut.Value = "08.00 AM"
''dtpIn.Value = "05.00 PM"

End Sub

Private Sub Refresh_Click()
    cmd_Assign.Enabled = True
    cmd_modify.Enabled = False
    del_leave = 0
    txt_pw.Visible = False
    txt_pw.Text = ""
End Sub

