VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form emp_mas_modify_new 
   Caption         =   "EMPLYOEE MODIFICATION"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   16890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1500
      Left            =   450
      TabIndex        =   172
      Top             =   360
      Width           =   10425
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
         Left            =   3960
         TabIndex        =   181
         Top             =   960
         Visible         =   0   'False
         Width           =   6150
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
         Left            =   5040
         TabIndex        =   178
         Top             =   120
         Width           =   2340
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00C0E0FF&
            Caption         =   "STAFF"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton opt_worker 
            BackColor       =   &H00C0E0FF&
            Caption         =   "WORKER"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1200
            TabIndex        =   179
            Top             =   240
            Width           =   1095
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
         Left            =   3720
         TabIndex        =   177
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Type"
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
         Left            =   7440
         TabIndex        =   173
         Top             =   120
         Width           =   2820
         Begin VB.OptionButton opt_All 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ALL"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   960
            TabIndex        =   176
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton opt_Active 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Active"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   120
            TabIndex        =   175
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton opt_resigned 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Resigned"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1680
            TabIndex        =   174
            Top             =   240
            Width           =   975
         End
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
         Left            =   480
         TabIndex        =   183
         Top             =   375
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
         Left            =   435
         TabIndex        =   182
         Top             =   960
         Width           =   3165
      End
   End
   Begin VB.TextBox emp_name 
      Height          =   435
      Left            =   4470
      MaxLength       =   50
      TabIndex        =   171
      Top             =   1260
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
      Left            =   2490
      MaxLength       =   10
      TabIndex        =   170
      Top             =   660
      Width           =   1575
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
      Left            =   8490
      TabIndex        =   6
      Top             =   7380
      Width           =   1965
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   570
      TabIndex        =   1
      Top             =   7380
      Width           =   3615
      Begin VB.CommandButton NEW 
         Caption         =   "&New"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "emp_mas_modify_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_edit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   735
         Left            =   960
         Picture         =   "emp_mas_modify_new.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_save 
         Caption         =   "&Save "
         Height          =   735
         Left            =   1800
         Picture         =   "emp_mas_modify_new.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   2640
         Picture         =   "emp_mas_modify_new.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox ctc 
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
      Left            =   8490
      TabIndex        =   0
      Top             =   7980
      Width           =   1965
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Left            =   450
      TabIndex        =   7
      Top             =   1860
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
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
      TabPicture(0)   =   "emp_mas_modify_new.frx":12A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmb_blood"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SEXFRAME"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "PF and DEPT. DETAILS"
      TabPicture(1)   =   "emp_mas_modify_new.frx":12BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label60"
      Tab(1).Control(1)=   "Label53"
      Tab(1).Control(2)=   "Label50"
      Tab(1).Control(3)=   "Label46"
      Tab(1).Control(4)=   "week_off"
      Tab(1).Control(5)=   "Label44"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "DESI"
      Tab(1).Control(8)=   "Label11"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label9"
      Tab(1).Control(11)=   "cmb_mc"
      Tab(1).Control(12)=   "cmb_classification"
      Tab(1).Control(13)=   "cmb_da_eligible"
      Tab(1).Control(14)=   "FPCODE"
      Tab(1).Control(15)=   "weekly_off_lst"
      Tab(1).Control(16)=   "qualify_cmb"
      Tab(1).Control(17)=   "desi_cmb"
      Tab(1).Control(18)=   "Frame3"
      Tab(1).Control(19)=   "emptype_cmb"
      Tab(1).Control(20)=   "work_cmb"
      Tab(1).Control(21)=   "dept_cmb"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "EARNINGS"
      TabPicture(2)   =   "emp_mas_modify_new.frx":12D8
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
      TabPicture(3)   =   "emp_mas_modify_new.frx":12F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pfamt"
      Tab(3).Control(1)=   "lic"
      Tab(3).Control(2)=   "rd"
      Tab(3).Control(3)=   "houserent"
      Tab(3).Control(4)=   "pfdeduction"
      Tab(3).Control(5)=   "bankdeduction"
      Tab(3).Control(6)=   "txt_wfund"
      Tab(3).Control(7)=   "txt_teadeduction"
      Tab(3).Control(8)=   "Label27"
      Tab(3).Control(9)=   "Label28"
      Tab(3).Control(10)=   "Label29"
      Tab(3).Control(11)=   "Label31"
      Tab(3).Control(12)=   "Label47"
      Tab(3).Control(13)=   "Label48"
      Tab(3).Control(14)=   "pfpercentage"
      Tab(3).Control(15)=   "Label51"
      Tab(3).Control(16)=   "Label52"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "BANK ACCOUNT DETAILS"
      TabPicture(4)   =   "emp_mas_modify_new.frx":1310
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_bank_ifsc"
      Tab(4).Control(1)=   "cmb_bank"
      Tab(4).Control(2)=   "txt_bank_acno"
      Tab(4).Control(3)=   "txt_email"
      Tab(4).Control(4)=   "lbl_bank_ifss"
      Tab(4).Control(5)=   "Label54"
      Tab(4).Control(6)=   "Label55"
      Tab(4).Control(7)=   "Label56"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "EMPLOYEE STATUS"
      TabPicture(5)   =   "emp_mas_modify_new.frx":132C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmb_empstatus"
      Tab(5).Control(1)=   "frame_resigned"
      Tab(5).Control(2)=   "Label49"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "OTHERS"
      TabPicture(6)   =   "emp_mas_modify_new.frx":1348
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txt_appointedby"
      Tab(6).Control(1)=   "txt_preinterviewby"
      Tab(6).Control(2)=   "txt_interviewername"
      Tab(6).Control(3)=   "txt_oe"
      Tab(6).Control(4)=   "txt_ie"
      Tab(6).Control(5)=   "Label64"
      Tab(6).Control(6)=   "Label65"
      Tab(6).Control(7)=   "Label66"
      Tab(6).Control(8)=   "Label67"
      Tab(6).Control(9)=   "Label68"
      Tab(6).Control(10)=   "Label69"
      Tab(6).ControlCount=   11
      Begin VB.TextBox txt_bank_ifsc 
         Height          =   450
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   186
         Top             =   2040
         Width           =   4605
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
         Left            =   495
         TabIndex        =   110
         Top             =   1980
         Width           =   5205
         Begin MSComCtl2.DTPicker doj 
            Height          =   300
            Left            =   3300
            TabIndex        =   111
            Top             =   345
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   153026561
            CurrentDate     =   37491
         End
         Begin MSComCtl2.DTPicker dob 
            Height          =   315
            Left            =   855
            TabIndex        =   112
            Top             =   300
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Format          =   153026561
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
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   375
            Width           =   990
         End
      End
      Begin VB.Frame SEXFRAME 
         Caption         =   "SEX"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   8880
         TabIndex        =   107
         Top             =   480
         Width           =   1425
         Begin VB.OptionButton MALE 
            Caption         =   "MALE"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton FEMALE 
            Caption         =   "FEMALE"
            ForeColor       =   &H00800000&
            Height          =   480
            Left            =   120
            TabIndex        =   108
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
         Left            =   9090
         TabIndex        =   90
         Top             =   2385
         Width           =   1155
      End
      Begin VB.ComboBox dept_cmb 
         Height          =   315
         Left            =   -71490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   1140
         Width           =   4080
      End
      Begin VB.ComboBox work_cmb 
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   1860
         Width           =   4065
      End
      Begin VB.ComboBox emptype_cmb 
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   2820
         Width           =   1650
      End
      Begin VB.Frame Frame3 
         Caption         =   "PF DETAILS "
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
         Height          =   1545
         Left            =   -74160
         TabIndex        =   62
         Top             =   3780
         Width           =   8235
         Begin VB.TextBox txt_uan 
            Height          =   330
            Left            =   840
            MaxLength       =   12
            TabIndex        =   188
            Top             =   1080
            Width           =   2085
         End
         Begin VB.TextBox PF 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   68
            Top             =   600
            Width           =   810
         End
         Begin VB.TextBox pfno 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   67
            Top             =   600
            Width           =   1485
         End
         Begin VB.OptionButton PF_ELIGIBLE 
            Caption         =   "PF ELIGIBLE"
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
            Height          =   375
            Left            =   720
            TabIndex        =   66
            Top             =   240
            Width           =   2235
         End
         Begin VB.OptionButton PF_NONELIGIBLE 
            Caption         =   "PF NON ELIGIBLE"
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
            Height          =   330
            Left            =   2280
            TabIndex        =   65
            Top             =   240
            Width           =   2160
         End
         Begin VB.CommandButton cmd_getpf 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Generate PF Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6360
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt_esino 
            Height          =   330
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   63
            Top             =   1080
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dt_pf_join 
            Height          =   300
            Left            =   6180
            TabIndex        =   69
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   153026561
            CurrentDate     =   37491
         End
         Begin VB.Label Label70 
            Caption         =   "UAN"
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
            Left            =   120
            TabIndex        =   189
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label26 
            Caption         =   "PF (%)"
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
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label30 
            Caption         =   "PF Number"
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
            Height          =   285
            Left            =   3000
            TabIndex        =   72
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Label59 
            Caption         =   "PF Joining Date"
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
            Left            =   4680
            TabIndex        =   71
            Top             =   285
            Width           =   1380
         End
         Begin VB.Label Label61 
            Caption         =   "ESI Number"
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
            Height          =   285
            Left            =   3000
            TabIndex        =   70
            Top             =   1080
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   465
         TabIndex        =   55
         Top             =   1320
         Width           =   9825
         Begin VB.ComboBox Religion_cmb 
            Height          =   315
            Left            =   1275
            TabIndex        =   58
            Top             =   180
            Width           =   2805
         End
         Begin VB.ComboBox Community_cmb 
            Height          =   315
            Left            =   5130
            TabIndex        =   57
            Top             =   195
            Width           =   1470
         End
         Begin VB.ComboBox caste_cmb 
            Height          =   315
            Left            =   7305
            TabIndex        =   56
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   59
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         Left            =   6120
         TabIndex        =   45
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.ComboBox desi_cmb 
         Height          =   315
         Left            =   -71490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   3120
         Width           =   1545
      End
      Begin VB.ComboBox qualify_cmb 
         Height          =   315
         Left            =   -71490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   3000
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
         Left            =   -71490
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   3660
         Width           =   1575
      End
      Begin VB.TextBox FPCODE 
         Height          =   495
         Left            =   -66000
         TabIndex        =   34
         Top             =   2520
         Width           =   1095
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   2220
         Width           =   1695
      End
      Begin VB.ComboBox cmb_da_eligible 
         Height          =   315
         Left            =   -71490
         Style           =   2  'Dropdown List
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   2940
         Width           =   1695
      End
      Begin VB.ComboBox cmb_classification 
         Height          =   315
         Left            =   -71490
         Style           =   2  'Dropdown List
         TabIndex        =   28
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
         ItemData        =   "emp_mas_modify_new.frx":1364
         Left            =   -71640
         List            =   "emp_mas_modify_new.frx":1366
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txt_bank_acno 
         Height          =   450
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   2640
         Width           =   4605
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relationship"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   480
         TabIndex        =   20
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox fathername 
            Height          =   375
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   21
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
            TabIndex        =   24
            Top             =   240
            Width           =   6195
         End
      End
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
         ItemData        =   "emp_mas_modify_new.frx":1368
         Left            =   -71880
         List            =   "emp_mas_modify_new.frx":136A
         TabIndex        =   19
         Text            =   "cmb_empstatus"
         Top             =   960
         Width           =   3255
      End
      Begin VB.Frame frame_resigned 
         Height          =   1695
         Left            =   -74520
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   9135
         Begin VB.ComboBox cmb_reason 
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
            Left            =   2640
            TabIndex        =   15
            Top             =   840
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker dt_resigned 
            Height          =   315
            Left            =   2640
            TabIndex        =   16
            Top             =   360
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
            Format          =   153026561
            CurrentDate     =   37491
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
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label58 
            Caption         =   "Reason for leaving"
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
            TabIndex        =   17
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmb_mc 
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
         Left            =   -66360
         TabIndex        =   13
         Text            =   "cmb_mc"
         Top             =   1920
         Width           =   1545
      End
      Begin VB.TextBox txt_appointedby 
         Height          =   495
         Left            =   -71640
         TabIndex        =   12
         Text            =   " "
         Top             =   2640
         Width           =   5415
      End
      Begin VB.TextBox txt_preinterviewby 
         Height          =   495
         Left            =   -71640
         TabIndex        =   11
         Text            =   " "
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txt_interviewername 
         Height          =   495
         Left            =   -71640
         TabIndex        =   10
         Text            =   " "
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox txt_oe 
         Height          =   615
         Left            =   -71160
         TabIndex        =   9
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txt_ie 
         Height          =   615
         Left            =   -68760
         TabIndex        =   8
         Top             =   3480
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   960
         TabIndex        =   91
         Top             =   2880
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
         TabPicture(0)   =   "emp_mas_modify_new.frx":136C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label63"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txt_phoneno"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "c_pin"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "c_add3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "c_add2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "c_add1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "PERMENANT ADDRESS"
         TabPicture(1)   =   "emp_mas_modify_new.frx":1388
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label7"
         Tab(1).Control(1)=   "Label6"
         Tab(1).Control(2)=   "chk"
         Tab(1).Control(3)=   "p_pin"
         Tab(1).Control(4)=   "p_add2"
         Tab(1).Control(5)=   "p_add3"
         Tab(1).Control(6)=   "p_add1"
         Tab(1).ControlCount=   7
         Begin VB.TextBox c_add1 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   101
            Top             =   585
            Width           =   5895
         End
         Begin VB.TextBox c_add2 
            Height          =   375
            Left            =   2310
            MaxLength       =   50
            TabIndex        =   100
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox c_add3 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   99
            Top             =   1380
            Width           =   5895
         End
         Begin VB.TextBox c_pin 
            Height          =   375
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   98
            Top             =   1830
            Width           =   1815
         End
         Begin VB.TextBox p_add1 
            Height          =   345
            Left            =   -72660
            MaxLength       =   50
            TabIndex        =   97
            Top             =   855
            Width           =   5895
         End
         Begin VB.TextBox p_add3 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   96
            Top             =   1560
            Width           =   5895
         End
         Begin VB.TextBox p_add2 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   95
            Top             =   1200
            Width           =   5895
         End
         Begin VB.TextBox p_pin 
            Height          =   345
            Left            =   -72660
            MaxLength       =   7
            TabIndex        =   94
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txt_phoneno 
            Height          =   375
            Left            =   5760
            MaxLength       =   25
            TabIndex        =   93
            Top             =   1800
            Width           =   2415
         End
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
            Left            =   -73800
            TabIndex        =   92
            Top             =   480
            Width           =   4575
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
            TabIndex        =   106
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
            TabIndex        =   105
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
            Left            =   -73800
            TabIndex        =   104
            Top             =   840
            Width           =   885
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
            TabIndex        =   103
            Top             =   1920
            Width           =   1020
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
            Left            =   4440
            TabIndex        =   102
            Top             =   1920
            Width           =   1200
         End
      End
      Begin VB.Label lbl_bank_ifss 
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
         TabIndex        =   187
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "BLOOD GROUP"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9060
         TabIndex        =   169
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
         Left            =   -73920
         TabIndex        =   168
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
         Left            =   -73920
         TabIndex        =   167
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
         Left            =   -73920
         TabIndex        =   166
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
         Left            =   -73920
         TabIndex        =   165
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label Label15 
         Height          =   435
         Left            =   -74415
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
         Top             =   1740
         Width           =   2235
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         Left            =   -73920
         TabIndex        =   140
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
         TabIndex        =   139
         Top             =   3540
         Width           =   2475
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
         TabIndex        =   138
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
         TabIndex        =   137
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label44 
         Caption         =   "Label44"
         Height          =   30
         Left            =   -73320
         TabIndex        =   136
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
         Left            =   -73920
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
         Top             =   2340
         Width           =   2415
      End
      Begin VB.Label Label46 
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
         Height          =   495
         Left            =   -67200
         TabIndex        =   131
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   130
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
         Left            =   -73920
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         Left            =   -73920
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
         Top             =   2760
         Width           =   1575
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
         Left            =   -74400
         TabIndex        =   122
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label60 
         Caption         =   "MACHINE"
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
         Left            =   -67320
         TabIndex        =   121
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label64 
         Caption         =   "Appointed by"
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
         Left            =   -74520
         TabIndex        =   120
         Top             =   2640
         Width           =   2685
      End
      Begin VB.Label Label65 
         Caption         =   "Preliminary Interview by"
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
         Left            =   -74520
         TabIndex        =   119
         Top             =   1920
         Width           =   2685
      End
      Begin VB.Label Label66 
         Caption         =   "Interviewed by"
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
         Left            =   -74520
         TabIndex        =   118
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label67 
         Caption         =   "Previous Experience(In Years)"
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
         Height          =   525
         Left            =   -74520
         TabIndex        =   117
         Top             =   3480
         Width           =   2565
      End
      Begin VB.Label Label68 
         Caption         =   "OE"
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
         Left            =   -71760
         TabIndex        =   116
         Top             =   3600
         Width           =   525
      End
      Begin VB.Label Label69 
         Caption         =   "IE"
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
         Left            =   -69480
         TabIndex        =   115
         Top             =   3720
         Width           =   525
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   645
      Left            =   0
      Top             =   6780
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
      Left            =   6810
      TabIndex        =   185
      Top             =   7380
      Width           =   1455
   End
   Begin VB.Label Label62 
      BackColor       =   &H00C0E0FF&
      Caption         =   "         CTC (Rs.)"
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
      Left            =   6690
      TabIndex        =   184
      Top             =   8100
      Width           =   1695
   End
End
Attribute VB_Name = "emp_mas_modify_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_move_Click()

End Sub

Private Sub cmd_update_address_Click()

End Sub



Private Sub cmb_empstatus_Click()
      If cmb_empstatus.Text = "RESIGNED" Then
          frame_resigned.Visible = True
       Else
          frame_resigned.Visible = False
       End If
End Sub

Private Sub emp_edit_Click()
savechk = 1
End Sub

Private Sub emp_save_Click()
   Dim ecat As String
   emp_idcode.Text = txt_empcode.Text
   txt_empcode.Text = txt_empcode.Text
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
   If Val(FPCODE.Text) = 0 And work_cmb.Text <> "COIMBATORE" Then
      MsgBox ("FP code is missing ")
      FPCODE.SetFocus
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
   If Trim(txt_phoneno.Text) = "" Then
      MsgBox ("Enter contact Number")
      txt_phoneno.SetFocus
      Exit Sub
   End If
      
   If Trim(c_add1.Text) = "" Then
      MsgBox ("Address should not be empty..")
      c_add1.SetFocus
      Exit Sub
   End If
   If Trim(c_add2.Text) = "" Then
      MsgBox ("Address should not be empty..")
      c_add2.SetFocus
      Exit Sub
   End If
   If Trim(cmb_mc.Text) = "" And data_source <> "H" Then
      MsgBox ("Select Machine.. ")
      cmb_mc.SetFocus
      Exit Sub
   End If
   
   If PF_ELIGIBLE.Value = True And Val(PF.Text) = 0 Then
      MsgBox ("PF % is Nil... check it..")
      PF.SetFocus
      Exit Sub
   End If
   If PF_ELIGIBLE.Value = True And Val(pfno.Text) = 0 Then
      MsgBox ("PF Number is Nil... check it..")
      PF.SetFocus
      Exit Sub
   End If
   If txt_interviewername.Text = "" Then
      MsgBox "Enter Interviewer name"
      txt_interviewername.SetFocus
      Exit Sub
   End If
   
   If txt_preinterviewby.Text = "" Then
      MsgBox "Enter Preliminary Interviewer name"
      txt_preinterviewby.SetFocus
      Exit Sub
   End If
   
   If txt_appointedby.Text = "" Then
      MsgBox "Enter Appointed by name"
      txt_appointedby.SetFocus
      Exit Sub
   End If
   
   
   Set paydb = New ADODB.Connection
   Set payrs = New ADODB.Recordset
   find_Grosspay
   paydb.Open pay
   sql = ("select * from emp_mas where emp_code = '" & emp_idcode & "' and emp_company = '" & company_code & "'")
   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   
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
      
   find_religioncode (Religion_cmb.Text)
   payrs.Fields("emp_religion") = dcode
   find_communitycode (Community_cmb.Text)
   payrs.Fields("emp_community") = dcode
   find_castecode (caste_cmb.Text)
   payrs.Fields("emp_caste") = dcode
   
   If M_YES.Value = True Then
      payrs.Fields("emp_marital") = "Y"
   Else
      payrs.Fields("emp_marital") = "N"
   End If
   payrs.Fields("emp_blood") = cmb_blood.Text
   payrs.Fields("emp_cadd1") = UCase(Trim(c_add1))
   payrs.Fields("emp_cadd2") = UCase(Trim(c_add2))
   payrs.Fields("emp_cadd3") = UCase(Trim(c_add3))
   payrs.Fields("emp_cpin") = c_pin
   payrs.Fields("emp_contactno") = txt_phoneno.Text
   payrs.Fields("emp_padd1") = UCase(Trim(p_add1))
   payrs.Fields("emp_padd2") = UCase(Trim(p_add2))
   payrs.Fields("emp_padd3") = UCase(Trim(p_add3))
   payrs.Fields("emp_ppin") = p_pin
      
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
   Else
      payrs.Fields("emp_cat") = "O"
   End If
   find_qualifycode (qualify_cmb.Text)
   payrs.Fields("emp_qualify") = dcode
   
   
      wplace = ""
   If PF_ELIGIBLE.Value = True Then
      payrs.Fields("emp_pfeligible") = "Y"
      payrs.Fields("emp_pfjoin_date") = dt_pf_join.Value
   Else
      payrs.Fields("emp_pfeligible") = "N"
      payrs.Fields("emp_pfjoin_date") = Null
   End If
   payrs.Fields("emp_pfp") = Val(PF)
   payrs.Fields("emp_pfno") = Val(pfno)
   payrs.Fields("emp_uan") = txt_uan.Text
   
   
   payrs.Fields("emp_holiday") = weekly_off_lst.Text
   
   If cmb_empstatus.Text = "CURRENT EMPLOYEE" Then
      payrs.Fields("emp_status") = "A"
   ElseIf cmb_empstatus.Text = "RESIGNED" Then
      payrs.Fields("emp_status") = "R"
   ElseIf cmb_empstatus.Text = "WORKING AS RETAINER" Then
      payrs.Fields("emp_status") = "B"
   ElseIf cmb_empstatus.Text = "WORKING AS TEMPORARY" Then
      payrs.Fields("emp_status") = "C"
   End If

''   payrs.Fields("emp_status") = Left(cmb_empstatus, 1)
   
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
   
   payrs.Fields("emp_esino") = txt_esino.Text
   If cmb_empstatus.Text = "RESIGNED" Then
       payrs.Fields("emp_resigneddate") = dt_resigned.Value
       payrs.Fields("emp_reason") = cmb_reason.Text
   Else
       payrs.Fields("emp_resigneddate") = Null
       payrs.Fields("emp_reason") = ""
      
   End If
   payrs.Fields("emp_work_unit") = cmb_mc.Text
   payrs.Fields("emp_interview_by") = txt_interviewername.Text
   payrs.Fields("emp_final_interview_by") = txt_preinterviewby.Text
   payrs.Fields("emp_appointed_by") = txt_appointedby.Text
   
   If txt_ie.Text = "" Then
         payrs.Fields("emp_preexp_inside") = 0
   Else
        payrs.Fields("emp_preexp_inside") = txt_ie.Text
   End If
   
   If txt_oe.Text = "" Then
        payrs.Fields("emp_preexp_outside") = 0
   Else
        payrs.Fields("emp_preexp_outside") = txt_oe.Text
   End If
   
   
   payrs.Update
   payrs.Close
   MsgBox ("Data updated")
         
   refresh_data

End Sub

Private Sub empedit_cmb_Click()
    dt_resigned.Value = Format(Now, "dd/mm/yyyy")
    savechk = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    If opt_staff.Value = True Then
        sql = ("select * from emp_mas where emp_name = '" & Trim(empedit_cmb.Text) & "' and emp_company = '" & company_code & "' and emp_Cat in ('S','M')")
    Else
        sql = ("select * from emp_mas where emp_name = '" & Trim(empedit_cmb.Text) & "' and emp_company = '" & company_code & "' and emp_Cat = 'W'")
    End If
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
    Else
       swmr = payrs.Fields("emp_cat")
       emp_idcode = payrs.Fields("emp_code")
       emp_name = empedit_cmb.Text 'empedit_cmb.ItemData(empedit_cmb.ListIndex)
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
       txt_phoneno.Text = payrs.Fields("emp_contactno")
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
       txt_uan.Text = payrs.Fields("emp_uan")
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
       If payrs.Fields("emp_pfeligible") = "Y" Then
          PF_ELIGIBLE.Value = True
          If IsNull(payrs.Fields("emp_pfjoin_date")) = False Then
             dt_pf_join.Value = payrs.Fields("emp_pfjoin_date")
          End If
       Else
           PF_NONELIGIBLE.Value = True
       End If
       weekly_off_lst.AddItem payrs.Fields("emp_holiday")
       weekly_off_lst.AddItem ("SUNDAY")
       weekly_off_lst.AddItem ("MONDAY")
       weekly_off_lst.AddItem ("TUESDAY")
       weekly_off_lst.AddItem ("WEDNESDAY")
       weekly_off_lst.AddItem ("THURSDAY")
       weekly_off_lst.AddItem ("FRIDAY")
       weekly_off_lst.AddItem ("SATURDAY")
       weekly_off_lst.Text = payrs.Fields("emp_holiday")
       frame_resigned.Visible = False
       
      
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
          cmb_reason.Text = payrs.Fields("emp_reason")
       End If
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
       txt_bank_acno.Text = payrs.Fields("emp_bank_acno")
       txt_bank_ifsc.Text = payrs.Fields("emp_bank_ifsc")
       txt_email.Text = payrs.Fields("emp_email")
       txt_esino.Text = payrs.Fields("emp_esino")
       cmb_mc.Text = payrs.Fields("emp_work_unit")
   If opt_staff.Value = True Then
      pfcalc = Val(Basic) + Val(spl_pay)
      If pfcalc >= 15000 Then
         pfcalc = 15000
      End If
      pfamt.Text = Round(((pfcalc) * Val(PF) / 100), 0)
   Else
      pfamt.Text = Round(((Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda)) * Val(PF) / 100), 0)
   End If
   
   pfpercentage.Caption = Trim(PF) + "%"
       
       find_Grosspay
       find_netpay
''       If payrs.Fields("emp_ctc") <> Null Then
''            ctc = payrs.Fields("emp_ctc")
''       Else
''            ctc.Text = 0
''       End If
    End If
    If PF_NONELIGIBLE.Value = True Then
       cmd_getpf.Visible = True
    Else
       cmd_getpf.Visible = False
    End If
    
    txt_interviewername.Text = payrs.Fields("emp_interview_by")
    txt_preinterviewby.Text = payrs.Fields("emp_final_interview_by")
    txt_appointedby.Text = payrs.Fields("emp_appointed_by")
    txt_oe.Text = payrs.Fields("emp_preexp_outside")
    txt_ie.Text = payrs.Fields("emp_preexp_inside")
    
    ''Modified on 22/08/2015 As informed by Sr.Mgr HR
    ''for MASTER CONTROL
    
    emp_save.Enabled = True
    

    
    
End Sub


Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    swmr = "S"
    cmb_reason.AddItem "CESSATION"
    cmb_reason.AddItem "SUPERANNUATION"
    cmb_reason.AddItem "RETIREMENT"
    cmb_reason.AddItem "DEATH IN SERVICE"
    cmb_reason.AddItem "PERMANENT DISABLEMENT"
    emp_idcode.Enabled = False
    cmb_da_eligible.AddItem "YES"
    cmb_da_eligible.AddItem "NO"
    cmb_mc.Clear
    
    If company_code = 1 Then
       cmb_mc.AddItem "PM1"
       cmb_mc.AddItem "PM2"
       cmb_mc.AddItem "PM3"
    ElseIf company_code = 3 Then
       cmb_mc.AddItem "VJPM"
    ElseIf company_code = 5 Then
       cmb_mc.AddItem "POWER"
    Else
       cmb_mc.AddItem "OIL"
    End If
    dob.Value = Now
    doj.Value = Now
    dt_resigned.Value = Now
    dob.Value = Format(Now, "dd/mm/yyyy")
    doj.Value = Format(Now, "dd/mm/yyyy")
    dt_resigned.Value = Format(Now, "dd/mm/yyyy")
    dt_pf_join.Value = Format(Now, "dd/mm/yyyy")
    
    ''emp_mas_entry.Caption = emp_mas_entry.Caption
    
''    savechk = 0
    savechk = 1
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    MALE.Value = True
    PF_ELIGIBLE.Value = True
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
    sql = "select * from payroll_bank order by bank_code"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
        cmb_bank.AddItem payrs("bank_name")
        cmb_bank.ItemData(cmb_bank.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    
    work_cmb.AddItem "MILL"
    work_cmb.AddItem "COIMBATORE"
    work_cmb.AddItem "SIVAKASI"
    work_cmb.AddItem "CHENNAI"
    
    PF_ELIGIBLE.Value = True
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
    cmb_empstatus.AddItem ("RESIGNED")
    cmb_empstatus.AddItem ("WORKING AS RETAINER")
    cmb_empstatus.AddItem ("WORKING AS TEMPORARY")
    
    cmb_empstatus.Text = "CURRENT EMPLOYEE"
    
    cmb_classification.AddItem "BELOW MANAGER"
    cmb_classification.AddItem "MANAGER & ABOVE"
    cmb_classification.AddItem "MANAGEMENT"
    cmb_classification.Text = "BELOW MANAGER"
    If data_source = "A" Then
       loc = ""
       loc2 = ""
    ElseIf data_source = "H" Then
       loc = " and emp_workplace = 'CBE'"
       loc2 = " and s_workplace = 'CBE'"
    Else
       loc = " and emp_workplace = 'MILL'"
       loc2 = " and s_workplace = 'MILL'"
    End If
''----
    loc = ""
    loc2 = ""
'-----
    
    emp_idcode.Enabled = False
    opt_staff.Enabled = True
    savechk = 1
    If opt_staff.Value = True Then
       opt_staff_Click
    Else
       opt_worker_Click
    End If
End Sub

Private Sub opt_staff_Click()
    empedit_cmb.Clear
    If savechk = 1 Then
        emp_name.Visible = False
        ''emp_idcode.Enabled = False
        empedit_cmb.Visible = True
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        If opt_All.Value = True Then
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat in ('S','M') " & loc & "    order by emp_name"
        ElseIf opt_Active.Value = True Then
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat in ('S','M') " & loc & "  and emp_status = 'A'  order by emp_name"
        Else
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat in ('S','M') " & loc & "  and emp_status = 'R'  order by emp_name"
        End If
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
     empedit_cmb.Clear
    If savechk = 1 Then
        emp_name.Visible = False
        empedit_cmb.Visible = True
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' " & loc & "   order by emp_name"
        
        If opt_All.Value = True Then
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' " & loc & "   order by emp_name"
        ElseIf opt_Active.Value = True Then
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' " & loc & "   and emp_status = 'A'   order by emp_name"
        Else
           sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_cat = 'W' " & loc & "   and emp_status = 'R'   order by emp_name"

        End If
        
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

Private Sub find_Grosspay()
   Gross.Text = Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda) + Val(hra) + Val(attall) + Val(ca) + Val(splall) + Val(teaall) + _
                Val(medall) + Val(washall) + Val(lta) + Val(mazall) + Val(fuelall) + Val(profall) + Val(phoneall) + Val(cityall) + Val(othall) + _
                Val(mealsall) + Val(eduall) + Val(healthall)
                
   Gross.Text = Format$(Val(Gross), "0.00")
End Sub

Private Sub find_netpay()
   
   Dim pfcalc, bonus, gratuity As Double
   If PF_ELIGIBLE.Value = True Then
        If opt_staff.Value = True Then
           pfcalc = Val(Basic) + Val(spl_pay)
           If pfcalc >= 15000 Then
              pfcalc = 15000
           End If
           pfamt = Round(((pfcalc) * Val(PF) / 100), 0)
           bonus = (Val(Basic) + Val(spl_pay)) * 8.33 / 100
           gratuity = (Val(Basic) + Val(spl_pay)) * 4.8 / 100
        Else
           pfamt = Round(((Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda)) * Val(PF) / 100), 0)
        End If
   Else
      pfamt = "0"
   End If
   NET_PAY.Text = Val(Gross.Text) - Val(pfamt) - Val(pfdeduction) - Val(rd) - Val(lic) - Val(houserent) - Val(bankdeduction) - Val(txt_wfund.Text) - Val(txt_teadeduction.Text)
   ctc.Text = Round(Val(Gross.Text) + Val(pfamt) + bonus + gratuity, 0)
End Sub


Public Sub refresh_data()
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
   savechk = 1
   txt_empcode.Text = ""
   txt_teadeduction.Text = ""
   txt_wfund.Text = ""
   txt_bank_acno.Text = ""
   emp_name = ""
   emp_idcode = ""
   txt_email.Text = ""
   cmd_getpf.Visible = False
   empedit_cmb.Text = ""
   FPCODE.Text = ""
   txt_esino.Text = ""
   eduall.Text = ""
   txt_interviewername.Text = ""
   txt_preinterviewby.Text = ""
   txt_appointedby.Text = ""
   txt_oe.Text = ""
   txt_ie.Text = ""
   
End Sub


