VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_mas_vou_payment 
   Caption         =   "VOUCHER PAYMENT MASTERS"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_upd2 
      Caption         =   "UPDATE"
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
      Left            =   11520
      TabIndex        =   167
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txt_newname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   165
      Top             =   2040
      Visible         =   0   'False
      Width           =   6075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2700
      Left            =   480
      TabIndex        =   132
      Top             =   0
      Width           =   10425
      Begin VB.TextBox FPCODE 
         Height          =   495
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   158
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txt_aadhaar 
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
         Left            =   3000
         MaxLength       =   12
         TabIndex        =   157
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox emp_name 
         Height          =   435
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   142
         Top             =   960
         Width           =   6075
      End
      Begin VB.TextBox emp_idcode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   141
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox empedit_cmb 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   3000
         TabIndex        =   137
         Top             =   960
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Frame EDIT_FRAME 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Staff / Worker "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   675
         Left            =   6405
         TabIndex        =   134
         Top             =   120
         Visible         =   0   'False
         Width           =   3780
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00C0E0FF&
            Caption         =   "STAFF"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   360
            TabIndex        =   136
            Top             =   315
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton opt_worker 
            BackColor       =   &H00C0E0FF&
            Caption         =   "WORKER"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2130
            TabIndex        =   135
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.TextBox txt_empcode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         TabIndex        =   133
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0E0FF&
         Caption         =   "FINGURE PASS CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   160
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label76 
         BackColor       =   &H00C0E0FF&
         Caption         =   "AADHAAR Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   360
         TabIndex        =   159
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label empcode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   360
         TabIndex        =   139
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label empname 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   360
         TabIndex        =   138
         Top             =   960
         Width           =   2085
      End
   End
   Begin VB.TextBox NET_PAY 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   8280
      TabIndex        =   5
      Top             =   8400
      Width           =   1965
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   8280
      Width           =   3615
      Begin VB.CommandButton NEW 
         Caption         =   "&New"
         Height          =   735
         Left            =   120
         Picture         =   "frm_mas_vou_payment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_edit 
         Caption         =   "&Edit"
         Height          =   735
         Left            =   960
         Picture         =   "frm_mas_vou_payment.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_save 
         Caption         =   "&Save "
         Height          =   735
         Left            =   1800
         Picture         =   "frm_mas_vou_payment.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   2640
         Picture         =   "frm_mas_vou_payment.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   794
      BackColor       =   12640511
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PERSONAL DETAILS"
      TabPicture(0)   =   "frm_mas_vou_payment.frx":12A0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(5)=   "cmb_blood"
      Tab(0).Control(6)=   "SEXFRAME"
      Tab(0).Control(7)=   "Frame2"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "PF and DEPT. DETAILS"
      TabPicture(1)   =   "frm_mas_vou_payment.frx":12BC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label53"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label50"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "week_off"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label44"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DESI"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label78"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label77"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmb_classification"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmb_da_eligible"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "weekly_off_lst"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "qualify_cmb"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "desi_cmb"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "emptype_cmb"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "work_cmb"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "dept_cmb"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmb_group"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmb_cost"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "EARNINGS"
      TabPicture(2)   =   "frm_mas_vou_payment.frx":12D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Basic"
      Tab(2).Control(1)=   "ser_wt"
      Tab(2).Control(2)=   "spl_pay"
      Tab(2).Control(3)=   "fda"
      Tab(2).Control(4)=   "hra"
      Tab(2).Control(5)=   "ca"
      Tab(2).Control(6)=   "splall"
      Tab(2).Control(7)=   "teaall"
      Tab(2).Control(8)=   "medall"
      Tab(2).Control(9)=   "attall"
      Tab(2).Control(10)=   "mazall"
      Tab(2).Control(11)=   "fuelall"
      Tab(2).Control(12)=   "profall"
      Tab(2).Control(13)=   "phoneall"
      Tab(2).Control(14)=   "cityall"
      Tab(2).Control(15)=   "othall"
      Tab(2).Control(16)=   "Gross"
      Tab(2).Control(17)=   "lta"
      Tab(2).Control(18)=   "washall"
      Tab(2).Control(19)=   "vda"
      Tab(2).Control(20)=   "mealsall"
      Tab(2).Control(21)=   "eduall"
      Tab(2).Control(22)=   "healthall"
      Tab(2).Control(23)=   "Label15"
      Tab(2).Control(24)=   "Label12"
      Tab(2).Control(25)=   "Label13"
      Tab(2).Control(26)=   "Label14"
      Tab(2).Control(27)=   "Label16"
      Tab(2).Control(28)=   "Label17"
      Tab(2).Control(29)=   "Label18"
      Tab(2).Control(30)=   "Label19"
      Tab(2).Control(31)=   "Label20"
      Tab(2).Control(32)=   "Label21"
      Tab(2).Control(33)=   "Label22"
      Tab(2).Control(34)=   "Label23"
      Tab(2).Control(35)=   "Label24"
      Tab(2).Control(36)=   "Label25"
      Tab(2).Control(37)=   "Label37"
      Tab(2).Control(38)=   "Label38"
      Tab(2).Control(39)=   "Label39"
      Tab(2).Control(40)=   "Label40"
      Tab(2).Control(41)=   "Label41"
      Tab(2).Control(42)=   "Label42"
      Tab(2).Control(43)=   "Label43"
      Tab(2).Control(44)=   "Label32"
      Tab(2).Control(45)=   "Label33"
      Tab(2).Control(46)=   "Label45"
      Tab(2).ControlCount=   47
      TabCaption(3)   =   "STANDARD DEDUCTIONS"
      TabPicture(3)   =   "frm_mas_vou_payment.frx":12F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt_tds_per"
      Tab(3).Control(1)=   "txt_tds_amount"
      Tab(3).Control(2)=   "pfamt"
      Tab(3).Control(3)=   "lic"
      Tab(3).Control(4)=   "rd"
      Tab(3).Control(5)=   "houserent"
      Tab(3).Control(6)=   "pfdeduction"
      Tab(3).Control(7)=   "bankdeduction"
      Tab(3).Control(8)=   "txt_wfund"
      Tab(3).Control(9)=   "txt_teadeduction"
      Tab(3).Control(10)=   "Label26"
      Tab(3).Control(11)=   "Label27"
      Tab(3).Control(12)=   "Label28"
      Tab(3).Control(13)=   "Label29"
      Tab(3).Control(14)=   "Label31"
      Tab(3).Control(15)=   "Label47"
      Tab(3).Control(16)=   "Label48"
      Tab(3).Control(17)=   "pfpercentage"
      Tab(3).Control(18)=   "Label51"
      Tab(3).Control(19)=   "Label52"
      Tab(3).ControlCount=   20
      TabCaption(4)   =   "BANK ACCOUNT DETAILS"
      TabPicture(4)   =   "frm_mas_vou_payment.frx":1310
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_bank_ifsc"
      Tab(4).Control(1)=   "cmd_ref_bank"
      Tab(4).Control(2)=   "cmb_bank"
      Tab(4).Control(3)=   "txt_bank_acno"
      Tab(4).Control(4)=   "txt_email"
      Tab(4).Control(5)=   "lbl_bank_ifsc"
      Tab(4).Control(6)=   "Label54"
      Tab(4).Control(7)=   "Label55"
      Tab(4).Control(8)=   "Label56"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "EMPLOYEE STATUS"
      TabPicture(5)   =   "frm_mas_vou_payment.frx":132C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "frame_resigned"
      Tab(5).ControlCount=   1
      Begin VB.ComboBox cmb_cost 
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
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   162
         Top             =   4080
         Width           =   3090
      End
      Begin VB.ComboBox cmb_group 
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
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   161
         Top             =   4680
         Width           =   3090
      End
      Begin VB.TextBox txt_bank_ifsc 
         Height          =   450
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   155
         Top             =   2040
         Width           =   4605
      End
      Begin VB.CommandButton cmd_ref_bank 
         Caption         =   "R"
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
         Left            =   -65760
         TabIndex        =   154
         Top             =   840
         Width           =   255
      End
      Begin VB.Frame frame_resigned 
         Caption         =   "RESIGNED DETAILS"
         Height          =   1815
         Left            =   -74520
         TabIndex        =   149
         Top             =   1320
         Width           =   9015
         Begin VB.ComboBox cmb_empstatus 
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
            ItemData        =   "frm_mas_vou_payment.frx":1348
            Left            =   3600
            List            =   "frm_mas_vou_payment.frx":134A
            TabIndex        =   152
            Text            =   "cmb_empstatus"
            Top             =   600
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker dt_resigned 
            Height          =   315
            Left            =   3600
            TabIndex        =   150
            Top             =   1200
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
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
            Format          =   152633345
            CurrentDate     =   37491
         End
         Begin VB.Label Label49 
            Caption         =   "EMPLOYEE WORK STATUS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1080
            TabIndex        =   153
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label57 
            Caption         =   "RESIGNED DATE"
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
            Height          =   375
            Left            =   1080
            TabIndex        =   151
            Top             =   1200
            Width           =   2175
         End
      End
      Begin VB.TextBox txt_tds_per 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72960
         TabIndex        =   145
         Top             =   4080
         Width           =   570
      End
      Begin VB.TextBox txt_tds_amount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72000
         TabIndex        =   143
         Top             =   4080
         Width           =   1650
      End
      Begin VB.Frame Frame2 
         Caption         =   "Date of "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   780
         Left            =   -74505
         TabIndex        =   81
         Top             =   1980
         Width           =   5205
         Begin MSComCtl2.DTPicker doj 
            Height          =   300
            Left            =   3300
            TabIndex        =   82
            Top             =   345
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   152633345
            CurrentDate     =   37491
         End
         Begin MSComCtl2.DTPicker dob 
            Height          =   315
            Left            =   855
            TabIndex        =   83
            Top             =   300
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Format          =   152633345
            CurrentDate     =   37491
         End
         Begin VB.Label Label1 
            Caption         =   "Joining"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   2400
            TabIndex        =   85
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Birth"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   105
            TabIndex        =   84
            Top             =   375
            Width           =   990
         End
      End
      Begin VB.Frame SEXFRAME 
         Caption         =   "SEX"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   -66120
         TabIndex        =   78
         Top             =   480
         Width           =   1425
         Begin VB.OptionButton MALE 
            Caption         =   "MALE"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton FEMALE 
            Caption         =   "FEMALE"
            ForeColor       =   &H00800000&
            Height          =   480
            Left            =   120
            TabIndex        =   79
            Top             =   480
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmb_blood 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65910
         TabIndex        =   64
         Top             =   2385
         Width           =   1155
      End
      Begin VB.ComboBox dept_cmb 
         Height          =   315
         Left            =   3510
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1140
         Width           =   4080
      End
      Begin VB.ComboBox work_cmb 
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1860
         Width           =   4065
      End
      Begin VB.ComboBox emptype_cmb 
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2220
         Width           =   4035
      End
      Begin VB.TextBox Basic 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72975
         TabIndex        =   60
         Top             =   720
         Width           =   1545
      End
      Begin VB.TextBox ser_wt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   59
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox spl_pay 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   58
         Top             =   1920
         Width           =   1545
      End
      Begin VB.TextBox fda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   57
         Top             =   2520
         Width           =   1545
      End
      Begin VB.TextBox hra 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72975
         TabIndex        =   56
         Top             =   3720
         Width           =   1545
      End
      Begin VB.TextBox ca 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   55
         Top             =   735
         Width           =   1545
      End
      Begin VB.TextBox splall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   54
         Top             =   1260
         Width           =   1545
      End
      Begin VB.TextBox teaall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   53
         Top             =   1740
         Width           =   1545
      End
      Begin VB.TextBox medall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   52
         Top             =   2220
         Width           =   1545
      End
      Begin VB.TextBox attall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   51
         Top             =   4440
         Width           =   1545
      End
      Begin VB.TextBox pfamt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72000
         TabIndex        =   50
         Top             =   1620
         Width           =   1650
      End
      Begin VB.TextBox lic 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72000
         TabIndex        =   49
         Top             =   2220
         Width           =   1650
      End
      Begin VB.TextBox rd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72000
         TabIndex        =   48
         Top             =   2820
         Width           =   1650
      End
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   -74535
         TabIndex        =   41
         Top             =   1410
         Width           =   9825
         Begin VB.ComboBox Religion_cmb 
            Height          =   315
            Left            =   1275
            TabIndex        =   44
            Top             =   180
            Width           =   2805
         End
         Begin VB.ComboBox Community_cmb 
            Height          =   315
            Left            =   5130
            TabIndex        =   43
            Top             =   195
            Width           =   1470
         End
         Begin VB.ComboBox caste_cmb 
            Height          =   315
            Left            =   7305
            TabIndex        =   42
            Top             =   210
            Width           =   2415
         End
         Begin VB.Label Label34 
            Caption         =   "Community"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   4170
            TabIndex        =   47
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label35 
            Caption         =   "Religion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   120
            TabIndex        =   46
            Top             =   225
            Width           =   705
         End
         Begin VB.Label Label36 
            Caption         =   "Caste"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6675
            TabIndex        =   45
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.TextBox mazall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   40
         Top             =   3960
         Width           =   1545
      End
      Begin VB.TextBox fuelall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   39
         Top             =   4620
         Width           =   1545
      End
      Begin VB.TextBox profall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66240
         TabIndex        =   38
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox phoneall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -66240
         TabIndex        =   37
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox cityall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66240
         TabIndex        =   36
         Top             =   1980
         Width           =   1515
      End
      Begin VB.TextBox othall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66225
         TabIndex        =   35
         Top             =   4185
         Width           =   1515
      End
      Begin VB.TextBox Gross 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66225
         TabIndex        =   34
         Top             =   4845
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         Caption         =   "Marital Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   -68880
         TabIndex        =   31
         Top             =   2115
         Width           =   2760
         Begin VB.OptionButton M_YES 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   165
            TabIndex        =   33
            Top             =   255
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.OptionButton M_NO 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1350
            TabIndex        =   32
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.ComboBox desi_cmb 
         Height          =   315
         Left            =   3510
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1500
         Width           =   4065
      End
      Begin VB.TextBox lta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   29
         Top             =   3360
         Width           =   1545
      End
      Begin VB.TextBox washall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   28
         Top             =   2700
         Width           =   1545
      End
      Begin VB.TextBox vda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72975
         TabIndex        =   27
         Top             =   3120
         Width           =   1545
      End
      Begin VB.ComboBox qualify_cmb 
         Height          =   315
         Left            =   3510
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   780
         Width           =   4050
      End
      Begin VB.TextBox houserent 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -72000
         TabIndex        =   25
         Top             =   3420
         Width           =   1650
      End
      Begin VB.TextBox mealsall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66225
         TabIndex        =   24
         Top             =   2580
         Width           =   1515
      End
      Begin VB.TextBox eduall 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66240
         TabIndex        =   23
         Top             =   3060
         Width           =   1515
      End
      Begin VB.ListBox weekly_off_lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3510
         TabIndex        =   22
         Top             =   3060
         Width           =   3420
      End
      Begin VB.TextBox healthall 
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
         Left            =   -66240
         TabIndex        =   21
         Top             =   3660
         Width           =   1575
      End
      Begin VB.TextBox pfdeduction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67200
         TabIndex        =   20
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox bankdeduction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67200
         TabIndex        =   19
         Top             =   2220
         Width           =   1695
      End
      Begin VB.ComboBox cmb_da_eligible 
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3480
         Width           =   1905
      End
      Begin VB.TextBox txt_wfund 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67200
         TabIndex        =   17
         Top             =   3420
         Width           =   1695
      End
      Begin VB.TextBox txt_teadeduction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67200
         TabIndex        =   16
         Top             =   2940
         Width           =   1695
      End
      Begin VB.ComboBox cmb_classification 
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2580
         Width           =   4035
      End
      Begin VB.ComboBox cmb_bank 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71640
         TabIndex        =   14
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txt_bank_acno 
         Height          =   450
         Left            =   -71640
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1440
         Width           =   4605
      End
      Begin VB.TextBox txt_email 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -71640
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2640
         Width           =   4605
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relationship"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   -74520
         TabIndex        =   7
         Top             =   480
         Width           =   8145
         Begin VB.OptionButton opt_relationship_2 
            Caption         =   "HUSBAND"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   480
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1395
         End
         Begin VB.OptionButton opt_relationship_1 
            Caption         =   "FATHER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox fathername 
            Height          =   375
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   8
            Top             =   480
            Width           =   6180
         End
         Begin VB.Label father 
            Caption         =   "Father's / Husband's Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   6195
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   -73965
         TabIndex        =   65
         Top             =   2955
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   4154
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12640511
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PRESENT ADDRESS"
         TabPicture(0)   =   "frm_mas_vou_payment.frx":134C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label63"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "c_pin"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "c_add3"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "c_add2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "c_add1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txt_phoneno"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "PERMENANT ADDRESS"
         TabPicture(1)   =   "frm_mas_vou_payment.frx":1368
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chk"
         Tab(1).Control(1)=   "p_add1"
         Tab(1).Control(2)=   "p_add3"
         Tab(1).Control(3)=   "p_add2"
         Tab(1).Control(4)=   "p_pin"
         Tab(1).Control(5)=   "Label6"
         Tab(1).Control(6)=   "Label7"
         Tab(1).ControlCount=   7
         Begin VB.CheckBox chk 
            Caption         =   "PICKUP FROM PRESENT ADDRESS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -74280
            TabIndex        =   148
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox txt_phoneno 
            Height          =   375
            Left            =   5640
            MaxLength       =   25
            TabIndex        =   146
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox c_add1 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   73
            Top             =   585
            Width           =   5895
         End
         Begin VB.TextBox c_add2 
            Height          =   375
            Left            =   2310
            MaxLength       =   50
            TabIndex        =   72
            Top             =   990
            Width           =   5895
         End
         Begin VB.TextBox c_add3 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   71
            Top             =   1380
            Width           =   5895
         End
         Begin VB.TextBox c_pin 
            Height          =   375
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   70
            Top             =   1830
            Width           =   1815
         End
         Begin VB.TextBox p_add1 
            Height          =   345
            Left            =   -72660
            MaxLength       =   50
            TabIndex        =   69
            Top             =   735
            Width           =   5895
         End
         Begin VB.TextBox p_add3 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   68
            Top             =   1440
            Width           =   5895
         End
         Begin VB.TextBox p_add2 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   67
            Top             =   1080
            Width           =   5895
         End
         Begin VB.TextBox p_pin 
            Height          =   345
            Left            =   -72660
            MaxLength       =   7
            TabIndex        =   66
            Top             =   1890
            Width           =   2175
         End
         Begin VB.Label Label63 
            Caption         =   "Contact No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   4320
            TabIndex        =   147
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label Label3 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   465
            TabIndex        =   77
            Top             =   735
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Pin code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   480
            TabIndex        =   76
            Top             =   1830
            Width           =   1560
         End
         Begin VB.Label Label6 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   -74235
            TabIndex        =   75
            Top             =   855
            Width           =   1125
         End
         Begin VB.Label Label7 
            Caption         =   "Pin code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74220
            TabIndex        =   74
            Top             =   1965
            Width           =   1020
         End
      End
      Begin VB.Label Label77 
         Caption         =   "COST TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   164
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label78 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   163
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label lbl_bank_ifsc 
         Caption         =   "Bank IFSC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   156
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "TDS %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -73680
         TabIndex        =   144
         Top             =   4200
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "BLOOD GROUP"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -65940
         TabIndex        =   131
         Top             =   2115
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "DEPARTMENT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   130
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "WORK PLACE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   129
         Top             =   1860
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "EMPLOYEE TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   128
         Top             =   2220
         Width           =   2175
      End
      Begin VB.Label DESI 
         Caption         =   "DESIGNATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   127
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label Label15 
         Height          =   435
         Left            =   -74415
         TabIndex        =   126
         Top             =   3135
         Width           =   1515
      End
      Begin VB.Label Label12 
         Caption         =   "Basic Pay"
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
         Height          =   480
         Left            =   -74715
         TabIndex        =   125
         Top             =   795
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Service Weightage"
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
         Left            =   -74730
         TabIndex        =   124
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label14 
         Caption         =   "Special Pay"
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
         Height          =   390
         Left            =   -74760
         TabIndex        =   123
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Fixed DA"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   122
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label17 
         Caption         =   "Variable DA"
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
         Height          =   225
         Left            =   -74730
         TabIndex        =   121
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label Label18 
         Caption         =   "House Rent Allow"
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
         Height          =   465
         Left            =   -74760
         TabIndex        =   120
         Top             =   3840
         Width           =   1710
      End
      Begin VB.Label Label19 
         Caption         =   "Conv. Allow"
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
         Height          =   330
         Left            =   -71280
         TabIndex        =   119
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label Label20 
         Caption         =   "Spl. Allow"
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
         Height          =   225
         Left            =   -71280
         TabIndex        =   118
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label Label21 
         Caption         =   "Tea Allow"
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
         Height          =   315
         Left            =   -71280
         TabIndex        =   117
         Top             =   1905
         Width           =   1230
      End
      Begin VB.Label Label22 
         Caption         =   "Medical Allow"
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
         Height          =   225
         Left            =   -71280
         TabIndex        =   116
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "Attn. Allow"
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
         Height          =   465
         Left            =   -74760
         TabIndex        =   115
         Top             =   4440
         Width           =   1560
      End
      Begin VB.Label Label24 
         Caption         =   "Washing Allow"
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
         Height          =   330
         Left            =   -71280
         TabIndex        =   114
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label Label25 
         Caption         =   " L.T.A"
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
         Height          =   225
         Left            =   -71280
         TabIndex        =   113
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label27 
         Caption         =   "PF AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -73680
         TabIndex        =   112
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label28 
         Caption         =   "L.I.C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   -73680
         TabIndex        =   111
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label29 
         Caption         =   "R.D."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -73680
         TabIndex        =   110
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label37 
         Caption         =   "MAZ.ALL"
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
         Height          =   345
         Left            =   -71280
         TabIndex        =   109
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label Label38 
         Caption         =   "Fuel Allow"
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
         Height          =   225
         Left            =   -71280
         TabIndex        =   108
         Top             =   4680
         Width           =   1395
      End
      Begin VB.Label Label39 
         Caption         =   "Prof.Dev."
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
         Height          =   330
         Left            =   -67920
         TabIndex        =   107
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Label40 
         Caption         =   "Phone Allow"
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
         Height          =   330
         Left            =   -67920
         TabIndex        =   106
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label41 
         Caption         =   "City Allow"
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
         Height          =   300
         Left            =   -67920
         TabIndex        =   105
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label Label42 
         Caption         =   "Other"
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
         Height          =   465
         Left            =   -67920
         TabIndex        =   104
         Top             =   4260
         Width           =   1395
      End
      Begin VB.Label Label43 
         Caption         =   "Gross Pay (Rs.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   -68040
         TabIndex        =   103
         Top             =   4950
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "QUALIFICATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   102
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label31 
         Caption         =   "HOUSE RENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -73680
         TabIndex        =   101
         Top             =   3540
         Width           =   1395
      End
      Begin VB.Label Label32 
         Caption         =   "Meals Allow"
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
         Height          =   300
         Left            =   -67920
         TabIndex        =   100
         Top             =   2580
         Width           =   1395
      End
      Begin VB.Label Label33 
         Caption         =   "Edu. Allow"
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
         Height          =   300
         Left            =   -67920
         TabIndex        =   99
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label44 
         Caption         =   "Label44"
         Height          =   30
         Left            =   1680
         TabIndex        =   98
         Top             =   2565
         Width           =   30
      End
      Begin VB.Label week_off 
         Caption         =   "WEEK HOLIDAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   97
         Top             =   3060
         Width           =   2175
      End
      Begin VB.Label Label45 
         Caption         =   "Health Allow"
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
         Left            =   -67920
         TabIndex        =   96
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label47 
         Caption         =   "PF DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   -69720
         TabIndex        =   95
         Top             =   1740
         Width           =   2055
      End
      Begin VB.Label Label48 
         Caption         =   "BANK LOAN DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -69720
         TabIndex        =   94
         Top             =   2340
         Width           =   2415
      End
      Begin VB.Label pfpercentage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   -71760
         TabIndex        =   93
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label50 
         Caption         =   "DA ELIGIBLE (Y/N)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   92
         Top             =   3540
         Width           =   2175
      End
      Begin VB.Label Label51 
         Caption         =   "WELFARE FUND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -69720
         TabIndex        =   91
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Label Label52 
         Caption         =   "TEA DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -69720
         TabIndex        =   90
         Top             =   2940
         Width           =   2055
      End
      Begin VB.Label Label53 
         Caption         =   "EMP. CLASSIFICATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   89
         Top             =   2580
         Width           =   2175
      End
      Begin VB.Label Label54 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   88
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label55 
         Caption         =   "Bank A/C No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   87
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label56 
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   86
         Top             =   2880
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   645
      Left            =   0
      Top             =   6420
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1138
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label79 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CHANGE Employee Name Name As"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11280
      TabIndex        =   166
      Top             =   1680
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.Label netpay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "NET PAY (Rs.)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6600
      TabIndex        =   140
      Top             =   8520
      Width           =   1455
   End
End
Attribute VB_Name = "frm_mas_vou_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paydb As New ADODB.Connection
Dim payrs As New ADODB.Recordset
Private Sub attall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub bankdeduction_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub Basic_Change()
  find_Grosspay
  find_netpay
End Sub
Private Sub ca_Change()
  find_Grosspay
  find_netpay
End Sub
Private Sub ca_KeyPress(KeyAscii As Integer)
  find_Grosspay
  find_netpay
  On Error GoTo err_handler
    chk_keyascii splall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub attall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay

 On Error GoTo err_handler
    chk_keyascii attall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub Basic_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
find_Grosspay
find_netpay
    chk_keyascii Basic, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub cityall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub cityall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii cityall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub cmb_empstatus_Click()
''       If cmb_empstatus.Text = "RESIGNED" Then
''          frame_resigned.Visible = True
''       Else
''          frame_resigned.Visible = False
''       End If
End Sub

Private Sub cmd_getpf_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select max(emp_pfno)+1 as eno from emp_voupay_mast  where emp_company = '" & company_code & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
        payrs.MoveFirst
        empedit_cmb.Clear
        While Not payrs.EOF
            pfno.Text = payrs("eno")
            payrs.MoveNext
        Wend
    End If
''    Set paydb = vbNullString
''    Set payrs = vbNullString
    
End Sub

Private Sub cmd_ref_bank_Click()
    cmb_bank.Clear
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from payroll_bank order by bank_name"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_bank.AddItem payrs("bank_name")
        cmb_bank.ItemData(cmb_bank.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub cmd_upd2_Click()
   If emp_idcode.Text = "" Then
      MsgBox ("Employee Not Selected...")
      Exit Sub
   End If
   If txt_newname.Text = "" Then
      MsgBox ("New Employee Name is Missing")
      Exit Sub
   End If



    sql = "update emp_voupay_mast set EMP_COSTTYPE = '" & cmb_cost.Text & "' ,EMP_WORKTYPE = '" & cmb_group.Text & "' , EMP_AADHAAR = '" & txt_aadhaar.Text & "' , EMP_NAME = '" & txt_newname.Text & "'     where emp_company = '" & company_code & "' and emp_name = '" & empedit_cmb.Text & "' and  emp_cat = 'R' and emp_code = " & Val(emp_idcode.Text)
    sql = "update emp_voupay_mast set emp_holiday = '" & weekly_off_lst.Text & "' ,emp_bank = " & cmb_bank.ItemData(cmb_bank.ListIndex) & " , emp_bank_acno = '" & txt_bank_acno.Text & "' , emp_bank_ifsc = '" & txt_bank_ifsc.Text & "' , emp_email = '" & txt_email.Text & "', EMP_COSTTYPE = '" & cmb_cost.Text & "' ,EMP_WORKTYPE = '" & cmb_group.Text & "' , EMP_AADHAAR = '" & txt_aadhaar.Text & "'  where emp_company = '" & company_code & "' and emp_name = '" & empedit_cmb.Text & "' and  emp_cat = 'R' and emp_code = " & Val(emp_idcode.Text)
    paydb.Execute sql
   
    MsgBox ("Updated...")
    
   txt_newname.Text = ""
End Sub

Private Sub eduall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub emp_idcode_KeyPress(KeyAscii As Integer)
''  On Error GoTo err_handler
''    If KeyAscii <> 8 Then chk_keyascii fda, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
End Sub

Private Sub emp_edit_Click()
    emp_idcode.Enabled = False
''    EDIT_FRAME.Visible = True
    opt_staff.Enabled = True
   '' opt_staff.SetFocus
    savechk = 1
    opt_staff_Click
End Sub


Private Sub emp_name_Change()
    If emp_idcode.Text = "" Then
       MsgBox ("Enter Employee code & Continue ...")
       emp_name.Text = ""
       Exit Sub
    End If
End Sub

Private Sub healthall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub opt_staff_Click()
    If savechk = 1 Then
        emp_name.Visible = False
        ''emp_idcode.Enabled = False
        empedit_cmb.Visible = True
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        If data_source = "A" Then
           sql = "select * from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat in ('S','M')  order by emp_name"
        ElseIf data_source = "H" Then
           sql = "select * from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat in ('S','M') and emp_classification  = 'A' order by emp_name"
        Else
           sql = "select * from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat in ('S','M') and emp_classification  = 'B' order by emp_name"
        End If
        
        
        sql = "select * from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat in ('R')  order by emp_name"
        
        
        paydb.Open pay
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF Then
            payrs.MoveFirst
            empedit_cmb.Clear
            While Not payrs.EOF
                empedit_cmb.AddItem payrs("emp_name")
                payrs.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub opt_worker_Click()
    If savechk = 1 Then
        emp_name.Visible = False
        empedit_cmb.Visible = True
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        sql = "select * from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat = 'W' order by emp_name"
        paydb.Open pay
        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
        If Not payrs.EOF Then
            payrs.MoveFirst
            empedit_cmb.Clear
            While Not payrs.EOF
                empedit_cmb.AddItem payrs("emp_name")
              ''  empedit_cmb.ItemData(empedit_cmb.NewIndex) = payrs(0)
                payrs.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub emp_save_Click()
   emp_idcode.Text = UCase(emp_idcode.Text)
   txt_empcode.Text = emp_idcode.Text
   If Trim(emp_idcode) = "" Then
      MsgBox ("Employee ID code is blank ")
      emp_idcode.SetFocus
      Exit Sub
   End If
   If Trim(txt_empcode) = "" Then
      MsgBox ("Employee code is blank ")
      txt_empcode.SetFocus
      Exit Sub
   End If
   If Trim(emp_name) = "" Then
      MsgBox ("Employee Name is blank - correct it ")
      emp_name.SetFocus
      Exit Sub
   End If
   If Trim(cmb_blood.Text) = "" Then
      MsgBox ("Blood type is blank ")
      cmb_blood.SetFocus
      Exit Sub
   End If
   
   If Trim(cmb_cost) = "" Then
      MsgBox ("Employee Cost type is blank - correct it ")
      cmb_cost.SetFocus
      Exit Sub
   End If
   
   If Trim(cmb_group) = "" Then
      MsgBox ("Employee Group type is blank - correct it ")
      cmb_group.SetFocus
      Exit Sub
   End If
   If Trim(txt_aadhaar.Text) = "" Then
      MsgBox ("Employee Aadhaar Number is blank ")
      txt_aadhaar.SetFocus
      Exit Sub
   End If
   
   If Trim(qualify_cmb) = "" Then
      MsgBox ("Employee Qualification is blank - correct it ")
      qualify_cmb.SetFocus
      Exit Sub
   End If
   
   If Trim(fathername) = "" Then
      MsgBox ("Employee father name Name is blank - correct it ")
      fathername.SetFocus
      Exit Sub
   End If
   If Trim(Religion_cmb) = "" Then
      MsgBox ("Employee Religion name is blank - correct it ")
      Religion_cmb.SetFocus
      Exit Sub
   End If
   If Trim(Community_cmb) = "" Then
      MsgBox ("Employee community is blank - correct it ")
      Community_cmb.SetFocus
      Exit Sub
   End If
   If Trim(caste_cmb) = "" Then
      MsgBox ("Employee caste is blank - correct it ")
      caste_cmb.SetFocus
      Exit Sub
   End If
   
   If M_YES.Value = False And M_NO.Value = False Then
      MsgBox ("Please select marital status")
      Exit Sub
   End If
   If Trim(dept_cmb) = "" Then
      MsgBox ("Department name is blank - Select department ")
      dept_cmb.SetFocus
      Exit Sub
   End If
   If Trim(desi_cmb) = "" Then
      MsgBox ("Designation name is blank - Select Designation")
      desi_cmb.SetFocus
      Exit Sub
   End If
   If Trim(work_cmb) = "" Then
      MsgBox ("working place in is blank - Enter data")
      work_cmb.SetFocus
      Exit Sub
   End If
   If Trim(emptype_cmb) = "" Then
      MsgBox ("Employee type is blank - Select Employee type")
      emptype_cmb.SetFocus
      Exit Sub
   End If
   If Val(Basic) = 0 Then
       MsgBox ("Enter BASIC Amount...")
       Basic.SetFocus
       Exit Sub
   End If
   If Trim(cmb_bank.Text) = "" Then
      MsgBox ("Select Bank ")
      cmb_bank.SetFocus
      Exit Sub
   End If
''''   If PF_ELIGIBLE.Value = True And Val(PF.Text) = 0 Then
''''      MsgBox ("PF % is Nil... check it..")
''''      PF.SetFocus
''''      Exit Sub
''''   End If
''   If PF_ELIGIBLE.Value = True And Val(pfno.Text) = 0 Then
''      MsgBox ("PF Number is Nil... check it..")
''      PF.SetFocus
''      Exit Sub
''   End If
   Set paydb = New ADODB.Connection
   Set payrs = New ADODB.Recordset
   find_Grosspay
   paydb.Open pay
   If savechk = 0 Then
      sql = "select * from emp_voupay_mast where emp_code = '" & emp_idcode.Text & "' and emp_company = '" & company_code & "'"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If Not payrs.EOF Then
         MsgBox ("Employee code Already Entered for ... " + payrs("emp_name"))
         payrs.Close
         paydb.Close
         Exit Sub
      End If
      payrs.Close
      sql = "Select * from emp_voupay_mast"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      payrs.AddNew
      payrs.Fields("emp_name") = UCase(emp_name.Text)
   Else
      sql = ("select * from emp_voupay_mast where emp_code = '" & emp_idcode & "' and emp_company = '" & company_code & "'")
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      payrs.Fields("emp_name") = UCase(empedit_cmb.Text)
   End If
   payrs.Fields("emp_company") = company_code
   payrs.Fields("emp_code") = emp_idcode.Text
   payrs.Fields("emp_fname") = UCase(Trim(fathername))
   If opt_relationship_1.Value = True Then
      payrs.Fields("emp_relation") = "F"
   Else
      payrs.Fields("emp_relation") = "H"
   End If
   If MALE.Value = True Then
      payrs.Fields("emp_sex") = "M"
   Else
      payrs.Fields("emp_sex") = "F"
   End If
   If savechk = 0 Then
      payrs.Fields("emp_religion") = Religion_cmb.ItemData(Religion_cmb.ListIndex)
      payrs.Fields("emp_community") = Community_cmb.ItemData(Community_cmb.ListIndex)
      payrs.Fields("emp_caste") = caste_cmb.ItemData(caste_cmb.ListIndex)
   Else
      find_religioncode (Religion_cmb.Text)
      payrs.Fields("emp_religion") = dcode
      find_communitycode (Community_cmb.Text)
      payrs.Fields("emp_community") = dcode
      find_castecode (caste_cmb.Text)
      payrs.Fields("emp_caste") = dcode
   End If
   payrs.Fields("emp_dob") = dob
   payrs.Fields("emp_doj") = doj
   If M_YES.Value = True Then
      payrs.Fields("emp_marital") = "Y"
   Else
      payrs.Fields("emp_marital") = "N"
   End If
   payrs.Fields("emp_blood") = cmb_blood.Text
   payrs.Fields("emp_cadd1") = c_add1
   payrs.Fields("emp_cadd2") = c_add2
   payrs.Fields("emp_cadd3") = c_add3
   payrs.Fields("emp_cpin") = c_pin
   payrs.Fields("emp_padd1") = p_add1
   payrs.Fields("emp_padd2") = p_add2
   payrs.Fields("emp_padd3") = p_add3
   payrs.Fields("emp_ppin") = p_pin
   If savechk = 0 Then
      payrs.Fields("emp_dept") = dept_cmb.ItemData(dept_cmb.ListIndex)
      payrs.Fields("emp_design") = desi_cmb.ItemData(desi_cmb.ListIndex)
      payrs.Fields("emp_type") = emptype_cmb.ItemData(emptype_cmb.ListIndex)
      payrs.Fields("emp_qualify") = qualify_cmb.ItemData(qualify_cmb.ListIndex)
      If emptype_cmb.ItemData(emptype_cmb.ListIndex) = 0 Or emptype_cmb.ItemData(emptype_cmb.ListIndex) = 1 Then
         payrs.Fields("emp_cat") = "S"
      ElseIf emptype_cmb.ItemData(emptype_cmb.ListIndex) = 2 Or emptype_cmb.ItemData(emptype_cmb.ListIndex) = 3 Then
         payrs.Fields("emp_cat") = "W"
      ElseIf emptype_cmb.ItemData(emptype_cmb.ListIndex) = 4 Then
         payrs.Fields("emp_cat") = "M"
''      ElseIf emptype_cmb.ItemData(emptype_cmb.ListIndex) = 7 Then
''         payrs.Fields("emp_cat") = "C"
      Else
         payrs.Fields("emp_cat") = "O"
      End If
  Else
      find_deptcode (dept_cmb.Text)
      payrs.Fields("emp_dept") = dcode
      find_designcode (desi_cmb.Text)
      payrs.Fields("emp_design") = dcode
      find_typecode (emptype_cmb.Text)
      payrs.Fields("emp_type") = dcode
      If dcode = 0 Or dcode = 1 Then
         payrs.Fields("emp_cat") = "S"
      ElseIf dcode = 2 Or dcode = 3 Then
         payrs.Fields("emp_cat") = "W"
      ElseIf dcode = 4 Then
         payrs.Fields("emp_cat") = "M"
''      ElseIf dcode = 7 Then
''         payrs.Fields("emp_cat") = "C"
      Else
         payrs.Fields("emp_cat") = "O"
      End If
      find_qualifycode (qualify_cmb.Text)
      payrs.Fields("emp_qualify") = dcode
   End If
   payrs.Fields("emp_cat") = "R"
   
''   If PF_ELIGIBLE.Value = True Then
''      payrs.Fields("emp_pfeligible") = "Y"
''   Else
''      payrs.Fields("emp_pfeligible") = "N"
''   End If
   payrs.Fields("emp_pfp") = Val(PF)
   payrs.Fields("emp_pfno") = Val(pfno)
   payrs.Fields("emp_basic") = Val(Basic)
   payrs.Fields("emp_serwt") = Val(ser_wt)
   payrs.Fields("emp_splpay") = Val(spl_pay)
   payrs.Fields("emp_fda") = Val(fda)
   payrs.Fields("emp_vda") = Val(vda)
   payrs.Fields("emp_hra") = Val(hra)
   payrs.Fields("emp_convall") = Val(ca)
   payrs.Fields("emp_medall") = Val(medall)
   payrs.Fields("emp_splall") = Val(splall)
   payrs.Fields("emp_teaall") = Val(teaall)
   payrs.Fields("emp_attall") = Val(attall)
   payrs.Fields("emp_healthall") = Val(healthall)
   payrs.Fields("emp_washall") = Val(washall)
   payrs.Fields("emp_mealsall") = Val(mealsall)
   payrs.Fields("emp_lta") = Val(lta)
   payrs.Fields("emp_eduall") = Val(eduall)
   payrs.Fields("emp_magall") = Val(mazall)
   payrs.Fields("emp_fuelall") = Val(fuelall)
   payrs.Fields("emp_profall") = Val(profall)
   payrs.Fields("emp_phoneall") = Val(phoneall)
   payrs.Fields("emp_cityall") = Val(cityall)
   payrs.Fields("emp_othall") = Val(othall)
   payrs.Fields("emp_lic") = Val(lic)
   payrs.Fields("emp_rd") = Val(rd)
   payrs.Fields("emp_pfded") = Val(pfduction)
   payrs.Fields("emp_bankded") = Val(bankdeduction)
   payrs.Fields("emp_houserent") = Val(houserent)
   payrs.Fields("emp_teaded") = Val(txt_teadeduction.Text)
   payrs.Fields("emp_wfund") = Val(txt_wfund.Text)
   payrs.Fields("emp_fpcode") = Val(FPCODE)
   payrs.Fields("emp_holiday") = weekly_off_lst.Text
   
   If cmb_empstatus.Text = "CURRENT EMPLOYEE" Then
      payrs.Fields("emp_status") = "A"
   ElseIf cmb_empstatus.Text = "RESIGNED" Then
      payrs.Fields("emp_status") = "R"
   ElseIf cmb_empstatus.Text = "WORKING AS RETAINER" Then
      payrs.Fields("emp_status") = "A"
   ElseIf cmb_empstatus.Text = "WORKING AS TEMPORARY" Then
      payrs.Fields("emp_status") = "C"
   End If

''   payrs.Fields("emp_status") = Left(cmb_empstatus, 1)
   
   payrs.Fields("emp_code") = txt_empcode.Text
   If cmb_da_eligible.Text = "YES" Then
       payrs.Fields("emp_da_eligible") = "Y"
   Else
       payrs.Fields("emp_da_eligible") = "N"
   End If
   If cmb_classification.Text = "BELOW MANAGER" Then
      payrs.Fields("emp_classification") = "B"
   ElseIf cmb_classification.Text = "MANAGEMENT" Then
      payrs.Fields("emp_classification") = "M"
   Else
      payrs.Fields("emp_classification") = "A"
   End If
   payrs.Fields("emp_bank") = cmb_bank.ItemData(cmb_bank.ListIndex)
   payrs.Fields("emp_bank_acno") = txt_bank_acno.Text
   payrs.Fields("emp_bank_ifsc") = txt_bank_ifsc.Text
   payrs.Fields("emp_email") = txt_email.Text
   payrs.Fields("emp_tds_per") = Val(txt_tds_per.Text)
   payrs.Fields("emp_tds_amount") = Val(txt_tds_amount.Text)
   payrs.Fields("emp_phoneno") = txt_phoneno.Text
   
   payrs.Fields("emp_aadhaar") = Left(Trim(txt_aadhaar.Text), 12)
   payrs.Fields("emp_costtype") = cmb_cost.Text
   payrs.Fields("emp_worktype") = cmb_group.Text
   
   If cmb_empstatus.Text = "RESIGNED" Then
       payrs.Fields("emp_resigneddate") = dt_resigned.Value
     
   Else
       payrs.Fields("emp_resigneddate") = Null
     

   End If
   
   payrs.Update
   payrs.Close
   MsgBox ("Data updated")
   refresh_data
End Sub

Private Sub empedit_cmb_Click()
    dt_resigned.Value = Now
    savechk = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_voupay_mast where emp_name = '" & Trim(empedit_cmb.Text) & "' and emp_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
    Else
       emp_idcode = payrs.Fields("emp_code")
       emp_name = empedit_cmb.Text 'empedit_cmb.ItemData(empedit_cmb.ListIndex)
       
       txt_newname.Text = empedit_cmb.Text
       
       fathername = payrs.Fields("emp_fname")
       If payrs.Fields("emp_sex") = "M" Then
           MALE.Value = True
       Else
           FEMALE.Value = True
       End If
       If payrs.Fields("emp_marital") = "Y" Then
           M_YES.Value = True
       Else
           M_NO.Value = True
       End If
       
       If payrs.Fields("emp_relation") = "F" Then
           opt_relationship_1.Value = True
       Else
           opt_relationship_2.Value = True
       End If
       
       weekly_off_lst.Text = payrs.Fields("emp_holiday")
       cmb_blood.Text = payrs.Fields("emp_blood")
       Religion_cmb.Text = payrs.Fields("emp_religion")
       Community_cmb.Text = payrs.Fields("emp_community")
       caste_cmb.Text = payrs.Fields("emp_caste")
       c_add1.Text = payrs.Fields("emp_cadd1")
       c_add2.Text = payrs.Fields("emp_cadd2")
       c_add3.Text = payrs.Fields("emp_cadd3")
       c_pin.Text = payrs.Fields("emp_cpin")
       p_add1.Text = payrs.Fields("emp_padd1")
       p_add2.Text = payrs.Fields("emp_padd2")
       p_add3.Text = payrs.Fields("emp_padd3")
       p_pin.Text = payrs.Fields("emp_ppin")
       Basic = payrs.Fields("emp_basic")
       ser_wt = payrs.Fields("emp_serwt")
       spl_pay = payrs.Fields("emp_splpay")
       fda = payrs.Fields("emp_fda")
       vda = payrs.Fields("emp_vda")
       hra = payrs.Fields("emp_hra")
       attall = payrs.Fields("emp_attall")
       ca = payrs.Fields("emp_convall")
       splall = payrs.Fields("emp_splall")
       teaall = payrs.Fields("emp_teaall")
       medall = payrs.Fields("emp_medall")
       washall = payrs.Fields("emp_washall")
       lta = payrs.Fields("emp_lta")
       mazall = payrs.Fields("emp_magall")
       fuelall = payrs.Fields("emp_fuelall")
       profall = payrs.Fields("emp_profall")
       cityall = payrs.Fields("emp_cityall")
       phoneall = payrs.Fields("emp_phoneall")
       healthall = payrs.Fields("emp_healthall")
       FPCODE = payrs.Fields("emp_fpcode")
       eduall = payrs.Fields("emp_eduall")
       mealsall = payrs.Fields("emp_mealsall")
       othall.Text = payrs.Fields("emp_othall")
       lic = payrs.Fields("emp_lic")
       houserent.Text = payrs.Fields("emp_houserent")
       rd = payrs.Fields("emp_rd")
       PF = payrs.Fields("emp_pfp")
       pfno = payrs.Fields("emp_pfno")
       pfdeduction = payrs.Fields("emp_pfded")
       bankdeduction = payrs.Fields("emp_bankded")
       txt_teadeduction.Text = payrs.Fields("emp_teaded")
       txt_wfund.Text = payrs.Fields("emp_wfund")
       find_deptname (payrs.Fields("emp_dept"))
       dept_cmb.Text = dname
       find_desiname (payrs.Fields("emp_design"))
       desi_cmb.Text = dname
       find_etypename (payrs.Fields("emp_type"))
       emptype_cmb.Text = dname
       dname = ""
       find_qualifyname (payrs.Fields("emp_qualify"))
       qualify_cmb.Text = dname
''       work_cmb.Text = payrs.Fields("emp_workplace")
       If payrs.Fields("emp_dob") <> " " Then
          dob = payrs.Fields("emp_dob")
       Else
          dob = DATE
       End If
       If payrs.Fields("emp_doj") <> " " Then
          doj = payrs.Fields("emp_doj")
       Else
          doj = DATE
       End If
       Set payrs2 = New ADODB.Recordset
       sql = ("select * from preli_mas where preli_code = " & Val(Religion_cmb.Text))
       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If Not payrs.EOF Then
          Religion_cmb.Text = payrs2.Fields("preli_name")
       End If
       Set payrs2 = New ADODB.Recordset
       sql = ("select * from pcomm_mas where pcomm_code = " & Val(Community_cmb.Text))
       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If Not payrs.EOF Then
          Community_cmb.Text = payrs2.Fields("pcomm_name")
       End If
       Set payrs2 = New ADODB.Recordset
       sql = ("select * from pcast_mas where pcast_code = " & Val(caste_cmb.Text))
       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
       If Not payrs.EOF Then
          caste_cmb.Text = payrs2.Fields("pcast_name")
       End If
''       If payrs.Fields("emp_pfeligible") = "Y" Then
''           PF_ELIGIBLE.Value = True
''       Else
''           PF_NONELIGIBLE.Value = True
''       End If
       weekly_off_lst.AddItem payrs.Fields("emp_holiday")
       weekly_off_lst.AddItem ("SUNDAY")
       weekly_off_lst.AddItem ("MONDAY")
       weekly_off_lst.AddItem ("TUESDAY")
       weekly_off_lst.AddItem ("WEDNESDAY")
       weekly_off_lst.AddItem ("THURSDAY")
       weekly_off_lst.AddItem ("FRIDAY")
       weekly_off_lst.AddItem ("SATURDAY")
       weekly_off_lst.Text = payrs.Fields("emp_holiday")
''          frame_resigned.Visible = False
       
''
''       If Left(payrs.Fields("emp_status"), 1) = "A" Then
''          cmb_empstatus.Text = "CURRENT EMPLOYEE"
''       ElseIf Left(payrs.Fields("emp_status"), 1) = "B" Then
''          cmb_empstatus.Text = "WORKING AS RETAINER"
''       ElseIf Left(payrs.Fields("emp_status"), 1) = "C" Then
''          cmb_empstatus.Text = "WORKING AS RETAINER"
''       ElseIf Left(payrs.Fields("emp_status"), 1) = "R" Then
''''          cmb_empstatus.Text = "RESIGNED"
''  ''        frame_resigned.Visible = True
''''          If IsNull(payrs.Fields("emp_resigneddate")) = False Then
''''             dt_resigned.Value = payrs.Fields("emp_resigneddate")
''''          End If
''''          cmb_reason.Text = payrs.Fields("emp_reason")
''       End If
       If payrs.Fields("emp_da_eligible") = "Y" Then
          cmb_da_eligible.Text = "YES"
       Else
          cmb_da_eligible.Text = "NO"
       End If
       txt_empcode.Text = payrs.Fields("emp_code")
       If payrs.Fields("emp_classification") = "B" Then
          cmb_classification.Text = "BELOW MANAGER"
       ElseIf payrs.Fields("emp_classification") = "M" Then
          cmb_classification.Text = "MANAGEMENT"
       Else
          cmb_classification.Text = "MANAGER & ABOVE"
       End If
       cmb_bank.ListIndex = payrs.Fields("emp_bank")
      cmb_bank.ListIndex = find_index_item_data(cmb_bank, payrs.Fields("emp_bank"))
       txt_bank_acno.Text = payrs.Fields("emp_bank_acno")
       txt_bank_ifsc.Text = payrs.Fields("emp_bank_ifsc")
       txt_email.Text = payrs.Fields("emp_email")
       
       txt_tds_per.Text = payrs.Fields("emp_tds_per")
       txt_tds_amount.Text = payrs.Fields("emp_tds_amount")
       txt_phoneno.Text = payrs.Fields("emp_phoneno")
       
       
       If Left(payrs.Fields("emp_status"), 1) = "A" Then
          cmb_empstatus.Text = "CURRENT EMPLOYEE"
       ElseIf Left(payrs.Fields("emp_status"), 1) = "B" Then
          cmb_empstatus.Text = "WORKING AS RETAINER"
       ElseIf Left(payrs.Fields("emp_status"), 1) = "C" Then
          cmb_empstatus.Text = "WORKING AS RETAINER"
       ElseIf Left(payrs.Fields("emp_status"), 1) = "R" Then
          cmb_empstatus.Text = "RESIGNED"
          frame_resigned.Visible = True
          If IsNull(payrs.Fields("emp_resigneddate")) = False Then
             dt_resigned.Value = payrs.Fields("emp_resigneddate")
          End If
       End If
       
       txt_aadhaar.Text = payrs.Fields("emp_aadhaar")
       cmb_cost.Text = payrs.Fields("emp_costtype")
       cmb_group.Text = payrs.Fields("emp_worktype")

       
       find_netpay
       
           ''Modified on 22/08/2015 As informed by Sr.Mgr HR
    ''for MASTER CONTROL
    
    emp_save.Enabled = True
    
    Set payrs2 = New ADODB.Recordset
    
    sql = "select count(*) as nos from emp_salary  where s_company = '" & company_code & "'  and s_empcode = '" & Trim(emp_idcode.Text) & "'"
    payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
      If payrs2.Fields("nos") > 0 And poweruser = 0 Then
         emp_save.Enabled = False
         MsgBox ("You can't Modify.. Only view...")
         
      End If
    End If
    payrs2.Close
    End If
''    If PF_NONELIGIBLE.Value = True Then
''       cmd_getpf.Visible = True
''    Else
''       cmd_getpf.Visible = False
''    End If
End Sub

Private Sub exit_Click()
  Unload Me
End Sub

Private Sub fda_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub fda_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
  On Error GoTo err_handler
    chk_keyascii fda, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub Form_Load()
''    cmb_reason.AddItem "CESSATION"
''    cmb_reason.AddItem "SUPERANNUATION"
''    cmb_reason.AddItem "RETIREMENT"
''    cmb_reason.AddItem "DEATH IN SERVICE"
''    cmb_reason.AddItem "PERMANENT DISABLEMENT"
''    emp_idcode.Enabled = False
    cmb_da_eligible.AddItem "YES"
    cmb_da_eligible.AddItem "NO"
    dob.Value = Now
    doj.Value = Now
    dt_resigned.Value = Now
    
    cmb_cost.AddItem "FIXED COST"
    cmb_cost.AddItem "VARIABLE COST"
    
    cmb_group.AddItem "SENIOR"
    cmb_group.AddItem "ESSENTIAL"
    cmb_group.AddItem "REGULAR"
    
''    dt_resigned.Value = Now
''    emp_voupay_mast_entry.Caption = emp_voupay_mast_entry.Caption
    savechk = 0
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    MALE.Value = True
''    PF_ELIGIBLE.Value = True
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  preli_mas order by preli_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        Religion_cmb.AddItem payrs(1)
        Religion_cmb.ItemData(Religion_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pcomm_mas order by pcomm_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        Community_cmb.AddItem payrs(1)
        Community_cmb.ItemData(Community_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pcast_mas order by pcast_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        caste_cmb.AddItem payrs(1)
        caste_cmb.ItemData(caste_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pdept_mas order by dept_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        dept_cmb.AddItem payrs(1)
        dept_cmb.ItemData(dept_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pdesi_mas order by pdesi_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        desi_cmb.AddItem payrs(1)
        desi_cmb.ItemData(desi_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pemptype_mas order by dtype_code")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        emptype_cmb.AddItem payrs(1)
        emptype_cmb.ItemData(emptype_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("Select * from  pqly_mas order by pqly_name")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
        qualify_cmb.AddItem payrs(1)
        qualify_cmb.ItemData(qualify_cmb.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select * from payroll_bank order by bank_name"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_bank.AddItem payrs("bank_name")
        cmb_bank.ItemData(cmb_bank.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    
    
'    PF_ELIGIBLE.Value = True
    cmb_blood.AddItem ("O +VE")
    cmb_blood.AddItem ("O -VE")
    cmb_blood.AddItem ("AB +VE ")
    cmb_blood.AddItem ("AB -VE")
    cmb_blood.AddItem ("A +VE")
    cmb_blood.AddItem ("A -VE")
    cmb_blood.AddItem ("A1 +VE")
    cmb_blood.AddItem ("A1 -VE")
    cmb_blood.AddItem ("A2 +VE")
    cmb_blood.AddItem ("A1 B+VE")
    cmb_blood.AddItem ("A1 B-VE")
    cmb_blood.AddItem ("A1B +VE")
    cmb_blood.AddItem ("A1B -VE")
    cmb_blood.AddItem ("A2B +VE")
    cmb_blood.AddItem ("A2B -VE")
    cmb_blood.AddItem ("B +VE")
    cmb_blood.AddItem ("B -VE")
    cmb_blood.AddItem ("B1 +VE")
    cmb_blood.AddItem ("B1 -VE")
    cmb_blood.AddItem ("OTHERS")
    weekly_off_lst.AddItem ("SUNDAY")
    weekly_off_lst.AddItem ("MONDAY")
    weekly_off_lst.AddItem ("TUESDAY")
    weekly_off_lst.AddItem ("WEDNESDAY")
    weekly_off_lst.AddItem ("THURSDAY")
    weekly_off_lst.AddItem ("FRIDAY")
    weekly_off_lst.AddItem ("SATURDAY")
    weekly_off_lst.Text = "SUNDAY"
    cmb_empstatus.Clear
    cmb_empstatus.AddItem ("CURRENT EMPLOYEE")

    cmb_empstatus.AddItem ("WORKING AS RETAINER")
    cmb_empstatus.AddItem ("RESIGNED")
    
    cmb_empstatus.Text = "WORKING AS RETAINER"
    
    cmb_classification.AddItem "BELOW MANAGER"
    cmb_classification.AddItem "MANAGER & ABOVE"
    cmb_classification.AddItem "MANAGEMENT"
    cmb_classification.Text = "BELOW MANAGER"
    
End Sub

Private Sub fuelall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub fuelall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii fuelall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub houserent_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub houserent_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler
    chk_keyascii houserent, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub hra_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub hra_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler
    chk_keyascii hra, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub lic_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub lic_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii lic, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub lta_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub lta_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii lta, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub mazall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub mazall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii mazall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub mealsall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub medall_Change()
  find_Grosspay
 find_netpay
End Sub

Private Sub medall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii medall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub NEW_Click()
    If opt_staff.Value = False And opt_worker.Value = False Then
       MsgBox ("Select Staff / Worker...")
       Exit Sub
    End If
    emp_name.Visible = True
    empedit_cmb.Visible = False
    refresh_data
    emp_idcode.Text = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select max(convert(int,emp_code))+1 as eno from emp_voupay_mast  where emp_company = '" & company_code & "' and emp_cat = 'R' "
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
            payrs.MoveFirst
            empedit_cmb.Clear
            While Not payrs.EOF
                If IsNull(payrs.Fields("eno")) = False Then
                   emp_idcode.Text = payrs.Fields("eno")
                End If
                payrs.MoveNext
            Wend
    End If
    payrs.Close
    
          
    
End Sub



Private Sub othall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub othall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii othall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub PF_Change()
   pfamt.Text = Round(((Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda)) * Val(PF) / 100), 0)
   pfpercentage.Caption = Trim(PF) + "%"
  find_Grosspay
  find_netpay
End Sub

Private Sub PF_ELIGIBLE_Click()
   PF.Enabled = True
   pfno.Enabled = True
   
End Sub

Private Sub PF_NONELIGIBLE_Click()
   PF.Enabled = False
   pfno.Enabled = False
   pfamt.Text = "0"
   PF = 0
   pfno = "  "
End Sub

Private Sub pfdeduction_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub phoneall_Change()
  find_Grosspay
  find_netpay
End Sub
Private Sub phoneall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii phoneall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub profall_Change()
  find_Grosspay
  
End Sub

Private Sub profall_KeyPress(KeyAscii As Integer)
find_Grosspay

On Error GoTo err_handler
    chk_keyascii profall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub rd_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub rd_KeyPress(KeyAscii As Integer)

 On Error GoTo err_handler
    chk_keyascii rd, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub ser_wt_Change()
 find_Grosspay
 find_netpay
End Sub

Private Sub ser_wt_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii ser_wt, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub spl_pay_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub spl_pay_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii spl_pay, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub splall_Change()
  find_Grosspay
 find_netpay
End Sub




Private Sub teaall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub teaall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
  On Error GoTo err_handler
    chk_keyascii teaall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txt_tds_per_Change()
     txt_tds_amount.Text = Round(Val(Gross.Text) * Val(txt_tds_per.Text) / 100, 0)
     find_Grosspay
     find_netpay
End Sub

Private Sub txt_tds_per_KeyPress(KeyAscii As Integer)

  On Error GoTo err_handler
    chk_keyascii txt_tds_per, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_teadeduction_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub txt_wfund_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub vda_Change()
    find_Grosspay
    find_netpay
End Sub

Private Sub vda_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii vda, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub washall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub washall_KeyPress(KeyAscii As Integer)
find_Grosspay
find_netpay
On Error GoTo err_handler
    chk_keyascii washall, "N", 5, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub
Private Sub find_Grosspay()
   Gross.Text = Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda) + Val(hra) + Val(attall) + Val(ca) + Val(splall) + Val(teaall) + _
                Val(medall) + Val(washall) + Val(lta) + Val(mazall) + Val(fuelall) + Val(profall) + Val(phoneall) + Val(cityall) + Val(othall) + _
                Val(mealsall) + Val(eduall) + Val(healthall)
                
   Gross.Text = Format$(Val(Gross), "0.00")
   txt_tds_amount.Text = Round(Val(Gross.Text) * Val(txt_tds_per.Text) / 100, 0)
End Sub

Private Sub find_netpay()
''   If PF_ELIGIBLE.Value = True Then
''      pfamt = Round(((Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda)) * Val(PF) / 100), 0)
''   Else
''      pfamt = "0"
''   End If
   NET_PAY.Text = Val(Gross.Text) - Val(pfamt) - Val(pfdeduction) - Val(rd) - Val(lic) - Val(houserent) - Val(bankdeduction) - Val(txt_wfund.Text) - Val(txt_teadeduction.Text) - Val(txt_tds_amount.Text)
End Sub


Public Sub refresh_data()
   txt_aadhaar.Text = ""
   fathername = " "
   Religion_cmb = " "
   Community_cmb = " "
   caste_cmb = " "
   c_add1 = " "
   c_add2 = ""
   c_add3 = " "
   p_pin = " "
   p_add1 = " "
   p_add2 = ""
   p_add3 = " "
   p_pin = " "
   ''dept_cmb = " "
   ''desi_cmb = " "
   ''emptype_cmb = ""
   ''qualify_cmb = ""
   pfno = " "
   Basic = ""
   ser_wt = ""
   spl_pay = ""
   fda = ""
   vda = ""
   hra = ""
   attall = ""
   ca = ""
   splall = ""
   teaall = ""
   medall = ""
   washall = ""
   lta = ""
   mazall = ""
   fuelall = ""
   profall = ""
   phoneall = ""
   ca = ""
   othall = ""
   lic = ""
   rd = ""
   houserent = ""
   savechk = 0
   txt_empcode.Text = ""
   txt_teadeduction.Text = ""
   txt_wfund.Text = ""
   txt_bank_acno.Text = ""
   emp_name = ""
   emp_idcode = ""
   txt_email.Text = ""
   txt_tds_amount.Text = 0
   txt_bank_ifsc.Text = ""
   txt_phoneno.Text = ""
''   cmd_getpf.Visible = False
End Sub




Private Sub chk_Click()
    If chk.Value = 1 Then
       p_add1.Text = c_add1.Text
       p_add2.Text = c_add2.Text
       p_add3.Text = c_add3.Text
       p_pin.Text = c_pin.Text
    Else
       p_add1.Text = ""
       p_add2.Text = ""
       p_add3.Text = ""
       p_pin.Text = ""
    End If
End Sub
