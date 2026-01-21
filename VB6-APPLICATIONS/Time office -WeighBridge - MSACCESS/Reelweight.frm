VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form ReelWeigt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SHVPM - Reel Weight Entry"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmb_finyear 
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
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame Frame6 
      Caption         =   "Add Wegiht"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15000
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cmb_addnl_wt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmd_on 
      Caption         =   "ON"
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   10080
      Width           =   255
   End
   Begin VB.TextBox txt_getwt 
      Height          =   375
      Left            =   15000
      TabIndex        =   45
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_wt_from_serailport 
      Height          =   375
      Left            =   15000
      TabIndex        =   44
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_wt2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   15000
      TabIndex        =   43
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_getwt2 
      Height          =   375
      Left            =   15000
      TabIndex        =   42
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "REEL WT FROM SCALE"
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
      Height          =   1095
      Left            =   14520
      TabIndex        =   40
      Top             =   2520
      Width           =   3255
      Begin VB.Label lbl_getwt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   480
         TabIndex        =   41
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   9840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   1200
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   1920
      TabIndex        =   27
      Top             =   9720
      Width           =   3855
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "Reelweight.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "Reelweight.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "Reelweight.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "Reelweight.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "Reelweight.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   13815
      Begin VB.ComboBox cmb_finyear_roll 
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
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1200
         Width           =   3135
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
         Left            =   6600
         TabIndex        =   53
         Text            =   "cmb_shift"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmb_winder 
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
         Height          =   360
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtoldWT 
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
         Left            =   10680
         TabIndex        =   47
         Top             =   7200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd_Save 
         Caption         =   "&SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7200
         Picture         =   "Reelweight.frx":1DEA
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmd_getwt 
         Caption         =   "&GET WEIGHT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7080
         Width           =   1575
      End
      Begin VB.ComboBox cmb_newvariety 
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
         Left            =   10320
         TabIndex        =   36
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton cmd_vartychange 
         BackColor       =   &H00C0FFC0&
         Caption         =   "VARIETY CHANGE"
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
         Left            =   6600
         TabIndex        =   35
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtWT 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   7320
         Width           =   1695
      End
      Begin VB.TextBox txtJoints 
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
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   25
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   24
         Top             =   6120
         Width           =   6375
      End
      Begin VB.ComboBox cmbsize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   360
         Left            =   3240
         TabIndex        =   23
         Top             =   5520
         Width           =   4335
      End
      Begin VB.ComboBox cmbSONO 
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
         Left            =   3240
         TabIndex        =   22
         Top             =   4920
         Width           =   2895
      End
      Begin VB.TextBox txtGSM 
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
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   6000
         TabIndex        =   21
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtBF 
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
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   3240
         TabIndex        =   20
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cmbReelNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   420
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4320
         Width           =   3735
      End
      Begin VB.ComboBox cmb_variety 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2880
         Width           =   3135
      End
      Begin VB.ComboBox cmbRollNo 
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
         Left            =   3240
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtMonth 
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
         Height          =   375
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dt_entrydate 
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
         Format          =   121176065
         CurrentDate     =   44676
      End
      Begin MSComCtl2.DTPicker dt_proddate 
         Height          =   495
         Left            =   6600
         TabIndex        =   33
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
         Format          =   121176065
         CurrentDate     =   44676
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parent Roll Fin.Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5880
         TabIndex        =   58
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lbladdnwt 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   9840
         TabIndex        =   55
         Top             =   6120
         Width           =   3855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5880
         TabIndex        =   54
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rewinder No."
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9480
         TabIndex        =   52
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblsize 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   48
         Top             =   5520
         Width           =   4335
      End
      Begin VB.Label lblvchange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Variety Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   8640
         TabIndex        =   37
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prodn. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4920
         TabIndex        =   34
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stock Entry Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Production YYMM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Roll No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Variety"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reel Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SO Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Joints"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weight in (Kgs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   3720
         Width           =   855
      End
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fin.Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9120
      TabIndex        =   57
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REEL WEIGHT - AUTOMATION "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SRI HARI VENKATESWARA PAPER MILLS (P) LTD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "ReelWeigt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mwt As Integer
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
Dim pst_qry, sizecode, destag As String
Dim mm, yy, runyear, rno, yr   As String
Dim varietycode, custcode, codesize, pin_cnt As Integer
Dim ginfincode As Integer
    
Dim addwt As Integer
    
    Dim firstno, lastno As Double
Dim winderno As String
''Private Declare Sub GenerateBMP _
''                Lib "C:\WINDOWS\system32\quricol32.dll" _
''                Alias "GenerateBMPW" ( _
''                    ByVal FileName As Long, _
''                ByVal Text As Long, _
''                ByVal Margin As Long, _
''                ByVal Size As Long, _
''                ByVal Level As TErrorCorretion)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    Dim compcode As Integer
    Dim fincode, rollfincode As Integer
    Dim saveflag As String
    
    


Private Sub cmb_finyear_Click()
    If cmb_finyear.ListIndex = -1 Then Exit Sub
    fincode = cmb_finyear.ItemData(cmb_finyear.ListIndex)
    rollfincode = cmb_finyear.ItemData(cmb_finyear.ListIndex)
    
End Sub

Private Sub cmb_finyear_roll_Click()
    If cmb_finyear_roll.ListIndex = -1 Then Exit Sub
    rollfincode = cmb_finyear_roll.ItemData(cmb_finyear_roll.ListIndex)
End Sub

Private Sub cmb_newvariety_Change()
''    varietycode = cmb_newvariety.ItemData(cmb_newvariety.ListIndex)
''    GetVarietyDetails
End Sub

Private Sub cmb_newvariety_Click()
    txtWT.Text = ""
    lbladdnwt.Caption = ""
    addwt = 0
    If cmb_newvariety.ListIndex = -1 Then Exit Sub
    varietycode = cmb_newvariety.ItemData(cmb_newvariety.ListIndex)
    GetVarietyDetails

End Sub

Private Sub cmb_variety_Change()
    If cmb_variety.ListIndex = -1 Then Exit Sub
    varietycode = cmb_variety.ItemData(cmb_variety.ListIndex)
    GetVarietyDetails
End Sub

Private Sub cmb_variety_Click()
    txtWT.Text = ""
    lbladdnwt.Caption = ""
    addwt = 0
    If cmb_variety.ListIndex = -1 Then Exit Sub
    varietycode = cmb_variety.ItemData(cmb_variety.ListIndex)
    GetVarietyDetails
End Sub

Private Sub cmbReelNo_Click()
    If saveflag = "edit" Then
        cmbsize.Clear

        yy = "20" + Trim(Left(txtMonth.Text, 2))
        mm = Right(txtMonth.Text, 2)

    destag = ""
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
        Dim pin_cnt As Integer

    ''    pst_qry = "select var_code, concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end),space(2) ,(case when var_shade = 'N' then 'NAT' when var_shade = 'G' then 'GYT' when var_shade = 'D' then 'DP' when var_shade = 'Y' then 'SHYS' when var_shade = 'B' then 'GB'   else 'OTH' end) ) as sizecode from massal_variety ,trnsal_order_trailer , masprd_variety where  var_grpcode = var_groupcode  " _
    ''              & " and var_grpcode = " & varietycode & " and ordt_var_code = var_code and  ordt_comp_code = 90 and ordt_sono = " & cmbSONO.Text & " order by sizecode"
    ''
''        pst_qry = "select  cust_ref,cust_code,var_code, concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end),space(2) ,(case when var_shade = 'N' then 'NAT' when var_shade = 'G' then 'GYT' when var_shade = 'D' then 'DP' when var_shade = 'Y' then 'SHYS' when var_shade = 'B' then 'GB'   else 'OTH' end) ) as sizecode  from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety,massal_customer " _
                   & " where ordh_party = cust_code and  ordh_comp_code = ordt_comp_code and ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and ordt_var_code = var_code and var_grpcode = var_groupcode and ordh_comp_code = " & compcode & "  and ordh_fincode = " & fincode & "  and var_grpcode = " & varietycode & "  and ordh_sono = " & cmbSONO.Text & "  order by sizecode"


        pst_qry = "select * , concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end) ,space(2) ,(case when var_shade = 'NS' then 'NAT' when var_shade = 'GY' then 'GYT' when var_shade = 'DP' then 'DP' when var_shade = 'SY' then 'SHYS' when var_shade = 'GB' then 'GB' when var_shade = 'BB' then 'BB' when var_shade = 'VV' then 'SHVV+'   else 'OTH' end)  ) as sizecode from trnsal_finish_stock , massal_variety ,trnsal_order_header,massal_customer where  stk_comp_code = " & compcode & "  and stk_finyear  = " & fincode & " and  stk_comp_code = ordh_comp_code and stk_finyear  >=  ordh_fincode  and  stk_sono = ordh_sono and  ordh_party = cust_code and stk_var_code  = var_code and stk_sr_no = " & Val(cmbReelNo.Text)
        
        adocmd_mysql.ActiveConnection = gen_connection_mysql
        cmbSONO.Clear
        cmbsize.Clear
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
            txtCustomer.Text = adors("cust_ref")
            custcode = adors("cust_code")
            cmbSONO.AddItem (adors("stk_sono"))
            cmbsize.AddItem (adors("sizecode"))
            cmbsize.ItemData(cmbsize.NewIndex) = adors("var_code")
            cmbSONO.Text = adors("stk_sono")
            cmbsize.Text = adors("sizecode")

            txtWT.Text = adors("stk_wt")
            txtoldWT.Text = adors("stk_wt")
            destag = adors("stk_destag")
            If adors("stk_destag") <> "" Then
               MsgBox ("Already Packslip Raised.. You can't Modify..")
            End If
        End If
        adors.Close

        End If
End Sub

Private Sub cmbRollNo_Change()
    GetVariety
    
End Sub

Private Sub cmbRollNo_Click()
    GetVariety
End Sub

Private Sub cmbsize_Click()

    
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer
    pst_qry = "select * from massal_variety where var_code = " & cmbsize.ItemData(cmbsize.ListIndex)


    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then

        codesize = adors("var_code")
        lblsize.Caption = adors("var_name")
    End If

    adors.Close
End Sub

Private Sub cmbSONO_Click()
    cmbsize.Clear
        
    addwt = 0
    yy = "20" + Trim(Left(txtMonth.Text, 2))
    mm = Right(txtMonth.Text, 2)
    lbladdnwt.Caption = ""
    
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer
              

''    pst_qry = "select  cust_addnlwt,cust_ref,cust_code,var_code, concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end),space(2) ,(case when var_shade = 'NS' then 'NAT' when var_shade = 'GY' then 'GYT' when var_shade = 'DP' then 'DP' when var_shade = 'SY' then 'SHYS' when var_shade = 'GB' then 'GB' when var_shade = 'VV' then 'SHVV+' when var_shade = 'BB' then 'BB'   else 'OTH' end) ) as sizecode  from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety,massal_customer " _
''               & " where ordh_party = cust_code and  ordh_comp_code = ordt_comp_code and ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and ordt_var_code = var_code and var_grpcode = var_groupcode and ordh_comp_code = " & compcode & "  and ordh_fincode <= " & fincode & "  and var_grpcode = " & varietycode & "  and ordh_sono = " & cmbSONO.Text & "  order by sizecode"

    pst_qry = "select  cust_addnlwt,cust_ref,cust_code,var_code, concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end),space(2) ,shade_shortname ) as sizecode  from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety,massal_customer,massal_shade   " _
               & " where var_shade = shade_shortcode and  ordh_party = cust_code and  ordh_comp_code = ordt_comp_code and ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and ordt_var_code = var_code and var_grpcode = var_groupcode and ordh_comp_code = " & compcode & "  and ordh_fincode <= " & fincode & "  and var_grpcode = " & varietycode & "  and ordh_sono = " & cmbSONO.Text & "  order by sizecode"
        
        
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        txtCustomer.Text = adors("cust_ref")
        custcode = adors("cust_code")
        If adors("cust_addnlwt") = "Y" Then
             addwt = 1
             lbladdnwt.Caption = "Additional Wt - 1 Kg Allowed"
        End If
        For pin_cnt = 1 To adors.RecordCount
                 cmbsize.AddItem (adors("sizecode"))
                 cmbsize.ItemData(cmbsize.NewIndex) = adors("var_code")

                 adors.MoveNext
        Next

        
    End If

    If cmbsize.ListCount > 1 Then
       cmbsize.ListIndex = 0
    End If
    adors.Close

End Sub

Private Sub cmd_getwt_Click()
    txtWT.Text = ""
    If cmbReelNo.Text = "" Then
        MsgBox ("Select Reel Number ...")
        Exit Sub
    End If
    
    txtWT.Text = Val(lbl_getwt.Caption) + Val(cmb_addnl_wt.Text) + addwt

End Sub

Private Sub cmd_on_Click()
    mwt = 1
End Sub

Private Sub cmd_Save_Click()
    

    If Val(txtWT.Text) = "0" Then
       MsgBox ("Weight is Empty Can't Save")
    Exit Sub
    End If
     
    If cmb_shift.Text = "" Then
       MsgBox ("Select Shift ")
    Exit Sub
    End If
        
    If destag <> "" Then
       MsgBox ("Already Packslip Raised.. You can't Modify..")
       Exit Sub
    End If
    datasave

  ''  Data_Refresh
End Sub

Private Sub cmd_vartychange_Click()
    lblvchange.Visible = True
    cmb_newvariety.Visible = True
End Sub

Private Sub Command1_Click()
''     MsgBox (Round(Val(txt_wt2.Text), 0))
''     MsgBox (CInt(Format(Val(txt_wt2.Text), "#0")))
End Sub

Private Sub edit_Click()
''   ''If Val(cmbRollNo.Text) > 0 Then Exit Sub
''   saveflag = "edit"
''    cmbReelNo.Clear
''
''    Dim firstno, lastno As Double
''    yy = "20" + Trim(Left(txtMonth.Text, 2))
''    mm = Right(txtMonth.Text, 2)
''    yr = Trim(Left(txtMonth.Text, 2))
''
''    Dim adocmd_mysql As New ADODB.Command
''    Dim adors As New ADODB.Recordset
''    Dim pin_cnt As Integer
''        adocmd_mysql.ActiveConnection = gen_connection_mysql
''    pst_qry = "select stk_sr_no  from trnsal_finish_stock WHERE  stk_comp_code = " & compcode & "  and stk_finyear  = " & fincode & " and length(stk_sr_no) = 10 and  SUBSTR(stk_sr_no,6,3)  = " & cmbRollNo.Text & "  and  SUBSTR(stk_sr_no,3,2)  = " & mm & "  and  SUBSTR(stk_sr_no,1,2)  = " & yr
''
''    adocmd_mysql.CommandText = pst_qry
''    Set adors = adocmd_mysql.Execute
''    If adors.RecordCount > 0 Then
''        For pin_cnt = 1 To adors.RecordCount
''            cmbReelNo.AddItem (adors("stk_sr_no"))
''            adors.MoveNext
''        Next
''    End If
''        adors.Close
End Sub

Private Sub exit_Click()
     Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err_handler:
      
    lbladdnwt.Caption = ""
    addwt = 0
    
''    On Error GoTo err_handler
''    Dim AdoCmd_getdate As New ADODB.Command
''    Dim rs_getdate As New ADODB.Recordset
''    Dim strcnncolormdups, strcnn_mysql As String
''    Dim pst_conn As String, pst_ret As Integer
''    strcnn_mysql = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.0.0.251; PORT = 3306; DATABASE=shvpm; USER=root; PASSWORD=P@ssw0rD; OPTION=3; CHARSET = UTF8; SOCKET = MYSQL"
''
''    Set gen_connection_mysql = Nothing
''    Set gen_connection_mysql = New ADODB.Connection
''    gen_connection_mysql.CursorLocation = adUseClient
''    gen_connection_mysql.Open strcnn_mysql
''
    
    
    
    
    dt_entrydate.Value = Now
    dt_proddate.Value = Now
   
    mwt = 0
    
    cmb_addnl_wt.AddItem "1"
''    cmb_addnl_wt.AddItem "2"
''    cmb_addnl_wt.AddItem "3"
    
    cmb_addnl_wt.AddItem "0.5"
    cmb_addnl_wt.AddItem "0"
    cmb_addnl_wt.Text = "0"
    
    cmb_winder.AddItem "1"
    cmb_winder.AddItem "2"
    cmb_winder.AddItem "0"
    cmb_winder.Text = "0"
    
    cmb_shift.AddItem "A"
    cmb_shift.AddItem "B"
    cmb_shift.AddItem "C"
    cmb_shift.Text = ""
    
    
    Call gen_dbconnection
  
    compcode = 1
    '' fincode = 23
    
    Dim pin_cnt As Integer
    pst_qry = "select fin_code,fin_year from mas_finyear where fin_code >=22 order by fin_code "
    adocmd_mysql.ActiveConnection = gen_connection_mysql
cmb_finyear.Clear
cmb_finyear_roll.Clear

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_finyear.AddItem (adors("fin_year"))
                 cmb_finyear.ItemData(cmb_finyear.NewIndex) = adors("fin_code")
                 cmb_finyear_roll.AddItem (adors("fin_year"))
                 cmb_finyear_roll.ItemData(cmb_finyear.NewIndex) = adors("fin_code")
                 cmb_finyear.Text = adors("fin_year")
                 cmb_finyear_roll.Text = adors("fin_year")
                 fincode = adors("fin_code")
                 rollfincode = adors("fin_code")
                 
                 adors.MoveNext
        Next
    End If
    adors.Close

          
   '' cmb_finyear.ListIndex = find_index_item_data(cmb_finyear, Val(ginfincode))
    
    Refresh_Click
    saveflag = "new"
    
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume

End Sub




Sub GetRollNumbers()
    
    yy = "20" + Trim(Left(txtMonth.Text, 2))
    mm = Right(txtMonth.Text, 2)

    cmbRollNo.Clear
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer
    pst_qry = "select prd_rollno  from trn_dayprod_roll_details where prd_compcode = " & compcode & "  and prd_fincode = " & fincode & "  and month(prd_date) = " & mm & "  and year(prd_date)= " & Int(yy) & " and prd_roll_status = 'A'  group by prd_rollno  order by prd_rollno desc"
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmbRollNo.AddItem (adors("prd_rollno"))
                 adors.MoveNext
        Next
    End If
    adors.Close
    pst_qry = "select var_desc,var_groupcode from masprd_variety order by var_desc,var_groupcode"
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_newvariety.AddItem (adors("var_desc"))
                 cmb_newvariety.ItemData(cmb_newvariety.NewIndex) = adors("var_groupcode")
                 adors.MoveNext
        Next
    End If
    adors.Close


End Sub


Sub GetVariety()
    
    If Val(cmbRollNo.Text) = 0 Then
       MsgBox ("Error in Roll No.  Please correct")
       cmbReelNo.Clear
       Exit Sub
    End If
    cmb_variety.Clear
    cmbReelNo.Clear

    yy = "20" + Trim(Left(txtMonth.Text, 2))
    mm = Right(txtMonth.Text, 2)


Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer
''    pst_qry = "select prd_date ,prd_shift,prd_seqno,var_desc,var_groupcode,prd_rollwt,prd_finprod,prd_rollno from trn_dayprod_roll_details, masprd_variety where prd_compcode = 90 and prd_fincode = 22 and " _
''              & " prd_rollno = " & Val(cmbRollNo.Text) & " and prd_variety = var_groupcode  and month(prd_date) =  " & mm & " and year(prd_date)= " & yy & " group by prd_date,prd_shift,prd_seqno,var_desc,var_groupcode,prd_rollwt,prd_finprod,prd_rollno order by var_desc,var_groupcode"
''
    pst_qry = "select prd_date ,var_desc,var_groupcode from trn_dayprod_roll_details, masprd_variety where prd_compcode =  " & compcode & " and prd_fincode = " & rollfincode & "  and " _
              & " prd_rollno = " & Val(cmbRollNo.Text) & " and prd_variety = var_groupcode  and month(prd_date) =  " & mm & " and year(prd_date)= " & yy & " group by prd_date,var_desc,var_groupcode order by var_desc,var_groupcode"

    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        dt_proddate.Value = adors("prd_date")
        For pin_cnt = 1 To adors.RecordCount
                 cmb_variety.AddItem (adors("var_desc"))
                 cmb_variety.ItemData(cmb_variety.NewIndex) = adors("var_groupcode")
                 adors.MoveNext
        Next
    End If
    adors.Close

     If cmb_variety.ListCount > 0 Then
       cmb_variety.ListIndex = 0
    End If

    rno = Right("00" + Trim(Str(cmbRollNo.Text)), 3)

    winderno = Left(cmb_winder.Text, 1)
    firstno = Int(Trim(Left(txtMonth.Text, 2)) + mm + winderno + rno + "01")
    lastno = firstno + 98

    Dim sno As Double
    For sno = firstno To lastno
         pst_qry = "select * from trnsal_finish_stock where stk_comp_code = " & compcode & "  and stk_finyear  = " & fincode & "  and stk_sr_no = " & sno
         adocmd_mysql.CommandText = pst_qry
         Set adors = adocmd_mysql.Execute
         If adors.RecordCount = 0 Then
            cmbReelNo.AddItem (sno)
         End If
         adors.Close
    Next

    If cmbReelNo.ListCount > 1 Then
       cmbReelNo.ListIndex = 0
    End If
    
End Sub

Sub GetVarietyDetails()


    yy = "20" + Trim(Left(txtMonth.Text, 2))
    mm = Right(txtMonth.Text, 2)

    cmbSONO.Clear
    cmbsize.Clear

Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer

    pst_qry = "select var_bf,var_gsm,prd_deckle,prd_breaks,prd_roll_dia,prd_rollwt,prd_set , prd_seqno from trn_dayprod_roll_details , masprd_variety where   prd_compcode = 90 and prd_fincode = 22 and prd_rollno = " & cmbRollNo.Text & "  and month(prd_date) =  " & mm & "  and year (prd_date)= " & yy & "  and prd_variety = var_groupcode  and prd_variety = " & varietycode
    pst_qry = "select var_bf ,var_gsm  from masprd_variety where  var_groupcode = " & varietycode
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 txtBF.Text = adors("var_bf")
                 txtGSM.Text = adors("var_gsm")
                 adors.MoveNext
        Next
    End If
    adors.Close

    pst_qry = "select ordh_sono from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety " _
              & " where ordh_comp_code = ordt_comp_code and  ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and  ordt_var_code = var_code and var_grpcode = var_groupcode and    ordh_comp_code = " & compcode & "  and  ordh_fincode <= " & fincode & "   and " _
              & " var_grpcode = " & varietycode & " group by ordh_sono order by ordh_sono desc"

    pst_qry = "select ordh_sono from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety " _
              & " where ordh_comp_code = ordt_comp_code and  ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and  ordt_var_code = var_code and var_grpcode = var_groupcode and    ordh_comp_code = " & compcode & "  and  ordh_fincode <= " & fincode & "   and " _
              & " var_grpcode = " & varietycode & " group by ordh_sono"

'' ordt_clo_stat = '' and
    pst_qry = "select  ordh_fincode,ordh_sono from trnsal_order_header, trnsal_order_trailer,massal_variety,masprd_variety " _
              & " where  ordh_comp_code = ordt_comp_code and  ordh_fincode  = ordt_fincode and ordh_sono = ordt_sono and  ordt_var_code = var_code and var_grpcode = var_groupcode and    ordh_comp_code = " & compcode & "  and  ordh_fincode <= " & fincode & "   and " _
              & " var_grpcode = " & varietycode & " group by  ordh_fincode,ordh_sono order by ordh_fincode desc, ordh_sono desc"


    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmbSONO.AddItem (adors("ordh_sono"))

                 adors.MoveNext
        Next
    End If

    If cmbSONO.ListCount > 1 Then
       cmbSONO.ListIndex = 0
    End If

    adors.Close
    

End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next:
  If mwt = 0 Then
    If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
    End If
  End If
End Sub

Private Sub Delay(ByVal mS As Long)
If mwt = 1 Then Exit Sub
Dim i As Integer
    For i = 1 To mS
        Sleep 1
        DoEvents
    Next i
End Sub

Private Sub NEW_Click()
    saveflag = "new"
End Sub
Sub Data_Refresh()
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    cmbRollNo.Clear
    ''runyear = "20" + yy
    pst_qry = "select prd_rollno  from trn_dayprod_roll_details where prd_compcode = " & compcode & "  and prd_fincode = " & fincode & "   and month(prd_date) = " & mm & "  and year(prd_date)= " & yy & " and prd_roll_status = 'A'  group by prd_rollno  order by prd_rollno desc"
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmbRollNo.AddItem (adors("prd_rollno"))
                 adors.MoveNext
        Next
    End If
    adors.Close
    If cmbRollNo.ListCount > 0 Then
       cmbRollNo.ListIndex = 0
    End If
    
    txtWT.Text = ""

End Sub


Private Sub Refresh_Click()
        
 
On Error GoTo err_handler:
   
   dt_entrydate.Value = Now
   dt_proddate.Value = Now
     
    cmbsize.Clear
    cmbSONO.Clear
    txtWT.Text = ""
    txtoldWT.Text = ""
    txtCustomer.Text = ""
    cmbReelNo.Clear
    destag = ""
    

    
    Dim pst_qry As String
    Dim pdb_seqno As Long
    Dim pdb_seqno_mysql As Long
    Dim pin_cnt As Long
    Dim pin_cnt2 As Double
    

    
    
    yy = Right(Str(Year(Now)), 2)
    mm = Right("0" + LTrim(RTrim(Str(Month(Now)))), 2)
    
    txtMonth.Text = yy + mm
    
    runyear = "20" + yy
    
    Data_Refresh

    Dim fy, mill, itemcode As Integer
    
    GetRollNumbers  '' For getting roll Numbers
    
    If mwt = 0 Then
        With MSComm1
                .Settings = "9600, N, 8, 1"
                .RTSEnable = True
                .DTREnable = True
                .RThreshold = 1
                .CommPort = 1
                 If MSComm1.PortOpen = False Then
                  .PortOpen = True
                 End If
        End With
    End If
    
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
    
End Sub

Private Sub save_click()

    If Val(txtWT.Text) = "0" Then
       MsgBox ("Weight is Empty Can't Save")
    Exit Sub
    End If
     
    If cmb_shift.Text = "" Then
       MsgBox ("Select Shift ")
    Exit Sub
    End If
        
    If destag <> "" Then
       MsgBox ("Already Packslip Raised.. You can't Modify..")
       Exit Sub
    End If
    datasave
End Sub
Function datasave()
    If codesize = 0 Then
       MsgBox ("Select Size..")
       Exit Function
    End If
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    Dim pst_qry As String
    
        
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    
    If saveflag = "new" Then
        pst_qry = "select * from trnsal_finish_stock where stk_comp_code = " & compcode & " and stk_finyear = " & fincode & " and stk_sr_no = " & Val(cmbReelNo.Text)
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
           MsgBox ("Reel Number Already Saved...")
           Exit Function
        End If
        
        pst_qry = "insert into trnsal_finish_stock  (stk_comp_code,stk_finyear,stk_ent_no,stk_ent_date,stk_var_code,stk_sr_no,stk_wt,stk_sono,stk_joints,stk_yymm, stk_rollno, stk_shift) VALUES ( " & compcode & ", " & fincode & "," & 100 & ",'" & Format(dt_entrydate, "yyyy-MM-dd") & "'," & codesize & " ," & cmbReelNo.Text & ", " & Val(txtWT.Text) & "," & cmbSONO.Text & "," & Val(txtJoints.Text) & " , " & Val(txtMonth.Text) & " , " & Val(cmbRollNo.Text) & " , '" & Left(cmb_shift.Text, 1) & "'  )"
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry


        pst_qry = "update trnsal_order_trailer set ordt_fin_wt =  ordt_fin_wt +  " & Val(txtWT.Text) / 1000 & "  where ordt_comp_code = " & compcode & "  and ordt_fincode <= " & fincode & "  and ordt_sono = " & cmbSONO.Text & "   and ordt_var_code = " & codesize & " "
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry
        
        pst_qry = "update trn_dayprod_roll_details set prd_finprod = prd_finprod +  " & Val(txtWT.Text) / 1000 & "  where prd_rollno = " & Val(cmbRollNo.Text) & "  and prd_date =  '" & Format(dt_entrydate, "yyyy-MM-dd") & "'  and prd_variety = " & varietycode & " "
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry

        
        MsgBox ("Reel Number Inserted in the Finished Stock..")
        
        
    Else
    
        pst_qry = "update trnsal_finish_stock set stk_sono = " & cmbSONO.Text & " , stk_var_code = " & codesize & " ,stk_wt = " & Val(txtWT.Text) & " ,stk_joints = " & Val(txtJoints.Text) & " where stk_sr_no =  '" & cmbReelNo.Text & "' and stk_comp_code = " & compcode & " and stk_finyear = " & fincode & ""
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry
        
        pst_qry = "update trnsal_order_trailer set ordt_fin_wt =  ordt_fin_wt +  " & Val(txtWT.Text) / 1000 & " -  " & Val(txtoldWT.Text) / 1000 & "  where ordt_comp_code = " & compcode & "  and ordt_fincode <= " & fincode & "  and ordt_sono = " & cmbSONO.Text & "   and ordt_var_code = " & codesize
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry
        MsgBox ("Reel Details Modified..")
    End If
    winderno = Left(cmb_winder.Text, 1)
    firstno = Int(Trim(Left(txtMonth.Text, 2)) + mm + winderno + rno + "01")
    lastno = firstno + 98
    
    cmbReelNo.Clear
    Dim sno As Double
    For sno = firstno To lastno
         pst_qry = "select * from trnsal_finish_stock where stk_comp_code = " & compcode & "  and stk_finyear  = " & fincode & "  and stk_sr_no = " & sno
         adocmd_mysql.CommandText = pst_qry
         Set adors = adocmd_mysql.Execute
         If adors.RecordCount = 0 Then
            cmbReelNo.AddItem (sno)
         End If
         adors.Close
    Next
    txtWT.Text = ""
''    txtCustomer.Text = ""

End Function


Private Sub Timer1_Timer()
     If mwt = 1 Then Exit Sub
   REFRESH_WEIGHT
End Sub




Private Sub MSComm1_OnComm()
 On Error GoTo err_handler
       If mwt = 0 Then
        Dim Buffer As String
        Buffer = MSComm1.Input
        txt_wt_from_serailport.SelText = Buffer
''        txt_getwt.Text = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, vbCrLf) + 1, 6)
''        txt_getwt.Text = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, " ") + 1, 6)
        txt_getwt = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, " ") + 2, 6)
        txt_wt2.Text = CInt(Format(Val(txt_getwt.Text), "#0"))
''        Round(Val(txt_getwt2.Text), 0)
'' lbl_getwt.Caption = Round(Val(txt_getwt.Text), 0)
        lbl_getwt.Caption = CInt(Format(Val(txt_getwt.Text), "#0"))
       End If
    Exit Sub
err_handler:
    MsgBox ("Port Not opened..")
    Exit Sub
    
End Sub


Sub REFRESH_WEIGHT()
  If mwt = 1 Then Exit Sub
 On Error GoTo err_handler
    txt_wt_from_serailport.Text = ""
    txt_getwt.Text = ""
    lbl_getwt.Caption = ""
   
        With MSComm1
            If .PortOpen = False Then
                 .CommPort = 1
                .PortOpen = True
            End If
            If .CDHolding = True Then
                Dim i As Integer
                For i = 1 To 3
                  .Output = "+"
                  Delay (50)
                Next i
                .Output = "AT+CSQ" & vbCr
                Delay (1000)
                .Output = "ATO & vbCr   'return to connected mode"
            End If
        End With
   
    Exit Sub
err_handler:
    MsgBox ("Port Not opened..")
    Exit Sub

End Sub

Private Sub txtMonth_Change()
    cmbsize.Clear

    yy = "20" + Trim(Left(txtMonth.Text, 2))
    mm = Right(txtMonth.Text, 2)


Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
    Dim pin_cnt As Integer

    pst_qry = "select var_code, concat(cast(var_size2 as CHAR) ,space(2) ,(case when var_inchcm = 'I' then 'Inch' else 'CM' end),space(2) ,(case when var_shade = 'N' then 'NAT' when var_shade = 'G' then 'GYT' when var_shade = 'D' then 'DP' when var_shade = 'Y' then 'SHYS' when var_shade = 'B' then 'GB'   else 'OTH' end) ) as sizecode from massal_variety ,trnsal_order_trailer , masprd_variety where  var_grpcode = var_groupcode  " _
              & " and var_grpcode = " & varietycode & " and ordt_var_code = var_code and  ordt_comp_code = 1 and ordt_sono = " & cmbSONO.Text & " order by sizecode"

    pst_qry = "select prd_rollno  from trn_dayprod_roll_details where prd_compcode = " & compcode & "  and prd_fincode = " & rollfincode & "  and month(prd_date) = " & mm & "  and year(prd_date)= " & yy & "  and prd_roll_status = 'A'  group by prd_rollno  order by prd_rollno desc"


    adocmd_mysql.ActiveConnection = gen_connection_mysql

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    cmbRollNo.Clear
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmbRollNo.AddItem (adors("prd_rollno"))

                 adors.MoveNext
        Next
    End If
    adors.Close
    If cmbRollNo.ListCount > 1 Then
       cmbRollNo.ListIndex = 0
    End If
    
End Sub


