VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox tincome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   8685
      TabIndex        =   8
      Top             =   7890
      Width           =   2445
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5220
      Left            =   735
      TabIndex        =   6
      Top             =   2475
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   9208
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12640511
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PERSONAL DETAILS"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "father"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fathername"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "age"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SEXFRAME"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SSTab2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "blood"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "marital"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "PF and DEPT. DETAILS"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "DESIGNATION"
      Tab(1).Control(2)=   "Combo1"
      Tab(1).Control(3)=   "work"
      Tab(1).Control(4)=   "dept"
      Tab(1).Control(5)=   "DESI"
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(8)=   "Label9"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "EARNINGS"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text9"
      Tab(2).Control(1)=   "othall"
      Tab(2).Control(2)=   "cityall"
      Tab(2).Control(3)=   "Text6"
      Tab(2).Control(4)=   "profall"
      Tab(2).Control(5)=   "fuelall"
      Tab(2).Control(6)=   "mazall"
      Tab(2).Control(7)=   "lta"
      Tab(2).Control(8)=   "washall"
      Tab(2).Control(9)=   "attall"
      Tab(2).Control(10)=   "medall"
      Tab(2).Control(11)=   "teaall"
      Tab(2).Control(12)=   "aplall"
      Tab(2).Control(13)=   "ca"
      Tab(2).Control(14)=   "hra"
      Tab(2).Control(15)=   "vda"
      Tab(2).Control(16)=   "fda"
      Tab(2).Control(17)=   "spl_pay"
      Tab(2).Control(18)=   "ser_wt"
      Tab(2).Control(19)=   "Basic"
      Tab(2).Control(20)=   "Label43"
      Tab(2).Control(21)=   "Label42"
      Tab(2).Control(22)=   "Label41"
      Tab(2).Control(23)=   "Label40"
      Tab(2).Control(24)=   "Label39"
      Tab(2).Control(25)=   "Label38"
      Tab(2).Control(26)=   "Label37"
      Tab(2).Control(27)=   "Label25"
      Tab(2).Control(28)=   "Label24"
      Tab(2).Control(29)=   "Label23"
      Tab(2).Control(30)=   "Label22"
      Tab(2).Control(31)=   "Label21"
      Tab(2).Control(32)=   "Label20"
      Tab(2).Control(33)=   "Label19"
      Tab(2).Control(34)=   "Label18"
      Tab(2).Control(35)=   "Label17"
      Tab(2).Control(36)=   "Label16"
      Tab(2).Control(37)=   "Label14"
      Tab(2).Control(38)=   "Label13"
      Tab(2).Control(39)=   "Label12"
      Tab(2).Control(40)=   "Label15"
      Tab(2).ControlCount=   41
      TabCaption(3)   =   "DEDUCTIONS"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text2"
      Tab(3).Control(1)=   "Text1"
      Tab(3).Control(2)=   "Text22"
      Tab(3).Control(3)=   "Text21"
      Tab(3).Control(4)=   "Text20"
      Tab(3).Control(5)=   "Label33"
      Tab(3).Control(6)=   "Label32"
      Tab(3).Control(7)=   "Label29"
      Tab(3).Control(8)=   "Label28"
      Tab(3).Control(9)=   "Label27"
      Tab(3).ControlCount=   10
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66345
         TabIndex        =   102
         Top             =   3555
         Width           =   1515
      End
      Begin VB.TextBox othall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66360
         TabIndex        =   101
         Top             =   2730
         Width           =   1515
      End
      Begin VB.TextBox cityall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66375
         TabIndex        =   100
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66345
         TabIndex        =   99
         Top             =   1575
         Width           =   1515
      End
      Begin VB.TextBox profall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66330
         TabIndex        =   98
         Top             =   1065
         Width           =   1515
      End
      Begin VB.TextBox fuelall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66345
         TabIndex        =   97
         Top             =   495
         Width           =   1515
      End
      Begin VB.TextBox mazall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   96
         Top             =   3930
         Width           =   1515
      End
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   465
         TabIndex        =   89
         Top             =   1065
         Width           =   9825
         Begin VB.TextBox caste 
            Height          =   300
            Left            =   7470
            TabIndex        =   95
            Top             =   210
            Width           =   2280
         End
         Begin VB.ComboBox Community 
            Height          =   315
            Left            =   5115
            TabIndex        =   94
            Top             =   195
            Width           =   1470
         End
         Begin VB.ComboBox Religion 
            Height          =   315
            Left            =   1275
            TabIndex        =   93
            Top             =   225
            Width           =   2805
         End
         Begin VB.Label Label36 
            Caption         =   "Caste"
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
            Height          =   255
            Left            =   6675
            TabIndex        =   92
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "Religion"
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
            Left            =   120
            TabIndex        =   91
            Top             =   225
            Width           =   705
         End
         Begin VB.Label Label34 
            Caption         =   "Community"
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
            Left            =   4170
            TabIndex        =   90
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.TextBox Text2 
         Height          =   510
         Left            =   -69900
         TabIndex        =   86
         Top             =   3915
         Width           =   1650
      End
      Begin VB.TextBox Text1 
         Height          =   510
         Left            =   -69915
         TabIndex        =   85
         Top             =   3105
         Width           =   1650
      End
      Begin VB.TextBox marital 
         Height          =   390
         Left            =   9360
         TabIndex        =   84
         Top             =   2115
         Width           =   675
      End
      Begin VB.Frame Frame3 
         Caption         =   "PF DETAILS "
         Height          =   1650
         Left            =   -73530
         TabIndex        =   76
         Top             =   3105
         Width           =   7875
         Begin VB.TextBox pfno 
            Height          =   450
            Left            =   5775
            TabIndex        =   82
            Top             =   1020
            Width           =   1845
         End
         Begin VB.TextBox PF 
            Height          =   390
            Left            =   2670
            TabIndex        =   80
            Top             =   1065
            Width           =   810
         End
         Begin VB.TextBox PFYESNO 
            Height          =   375
            Left            =   2670
            MaxLength       =   1
            TabIndex        =   78
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label30 
            Caption         =   "PF Number"
            Height          =   405
            Left            =   4200
            TabIndex        =   81
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label26 
            Caption         =   "PF (%)"
            Height          =   375
            Left            =   420
            TabIndex        =   79
            Top             =   1035
            Width           =   1740
         End
         Begin VB.Label Label8 
            Caption         =   "PF ELIGIBLE (Y/N) "
            Height          =   375
            Left            =   480
            TabIndex        =   77
            Top             =   315
            Width           =   1875
         End
      End
      Begin VB.TextBox Text22 
         Height          =   510
         Left            =   -69915
         TabIndex        =   75
         Top             =   2235
         Width           =   1650
      End
      Begin VB.TextBox Text21 
         Height          =   510
         Left            =   -69900
         TabIndex        =   74
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox Text20 
         Height          =   510
         Left            =   -69900
         TabIndex        =   71
         Top             =   750
         Width           =   1650
      End
      Begin VB.TextBox lta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   69
         Top             =   3270
         Width           =   1515
      End
      Begin VB.TextBox washall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   67
         Top             =   2685
         Width           =   1500
      End
      Begin VB.TextBox attall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73020
         TabIndex        =   66
         Top             =   4035
         Width           =   1545
      End
      Begin VB.TextBox medall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   65
         Top             =   2055
         Width           =   1455
      End
      Begin VB.TextBox teaall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69615
         TabIndex        =   64
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox aplall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69645
         TabIndex        =   63
         Top             =   1035
         Width           =   1440
      End
      Begin VB.TextBox ca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69630
         TabIndex        =   62
         Top             =   435
         Width           =   1425
      End
      Begin VB.TextBox hra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73005
         TabIndex        =   61
         Top             =   3300
         Width           =   1530
      End
      Begin VB.TextBox vda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73005
         TabIndex        =   60
         Top             =   2670
         Width           =   1530
      End
      Begin VB.TextBox fda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73035
         TabIndex        =   59
         Top             =   2040
         Width           =   1560
      End
      Begin VB.TextBox spl_pay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73020
         TabIndex        =   58
         Top             =   1515
         Width           =   1515
      End
      Begin VB.TextBox ser_wt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73005
         TabIndex        =   57
         Top             =   1020
         Width           =   1500
      End
      Begin VB.TextBox Basic 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73005
         TabIndex        =   56
         Top             =   405
         Width           =   1530
      End
      Begin VB.TextBox DESIGNATION 
         Height          =   465
         Left            =   -71445
         TabIndex        =   42
         Top             =   1170
         Width           =   4290
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -71445
         TabIndex        =   39
         Top             =   2400
         Width           =   3555
      End
      Begin VB.ComboBox work 
         Height          =   315
         Left            =   -71445
         TabIndex        =   38
         Top             =   1890
         Width           =   2745
      End
      Begin VB.ComboBox dept 
         Height          =   315
         Left            =   -71445
         TabIndex        =   37
         Top             =   615
         Width           =   4080
      End
      Begin VB.ComboBox blood 
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
         Left            =   7800
         TabIndex        =   27
         Top             =   2130
         Width           =   1155
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   1020
         TabIndex        =   20
         Top             =   2655
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PRESENT ADDRESS"
         TabPicture(0)   =   "Form1.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "c_add1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "c_add2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "c_add3"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "c_pin"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "PERMENANT ADDRESS"
         TabPicture(1)   =   "Form1.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "Label7"
         Tab(1).Control(2)=   "p_add1"
         Tab(1).Control(3)=   "p_add3"
         Tab(1).Control(4)=   "p_add2"
         Tab(1).Control(5)=   "p_pin"
         Tab(1).ControlCount=   6
         Begin VB.TextBox p_pin 
            Height          =   345
            Left            =   -72645
            TabIndex        =   32
            Top             =   1650
            Width           =   2175
         End
         Begin VB.TextBox p_add2 
            Height          =   345
            Left            =   -72690
            TabIndex        =   31
            Top             =   855
            Width           =   5895
         End
         Begin VB.TextBox p_add3 
            Height          =   345
            Left            =   -72675
            TabIndex        =   30
            Top             =   1200
            Width           =   5895
         End
         Begin VB.TextBox p_add1 
            Height          =   345
            Left            =   -72675
            TabIndex        =   29
            Top             =   510
            Width           =   5895
         End
         Begin VB.TextBox c_pin 
            Height          =   375
            Left            =   2295
            TabIndex        =   25
            Top             =   1830
            Width           =   1815
         End
         Begin VB.TextBox c_add3 
            Height          =   375
            Left            =   2295
            TabIndex        =   24
            Top             =   1380
            Width           =   5895
         End
         Begin VB.TextBox c_add2 
            Height          =   375
            Left            =   2295
            TabIndex        =   23
            Top             =   990
            Width           =   5895
         End
         Begin VB.TextBox c_add1 
            Height          =   375
            Left            =   2295
            TabIndex        =   22
            Top             =   615
            Width           =   5895
         End
         Begin VB.Label Label7 
            Caption         =   "Pin code"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74220
            TabIndex        =   34
            Top             =   1725
            Width           =   1020
         End
         Begin VB.Label Label6 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   33
            Top             =   615
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "Pin code"
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   480
            TabIndex        =   26
            Top             =   1830
            Width           =   1560
         End
         Begin VB.Label Label3 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   21
            Top             =   735
            Width           =   1695
         End
      End
      Begin VB.Frame SEXFRAME 
         Caption         =   "SEX"
         ForeColor       =   &H00C00000&
         Height          =   540
         Left            =   7815
         TabIndex        =   17
         Top             =   360
         Width           =   2505
         Begin VB.OptionButton FEMALE 
            Caption         =   "FEMALE"
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1215
            TabIndex        =   19
            Top             =   165
            Width           =   1155
         End
         Begin VB.OptionButton MALE 
            Caption         =   "MALE"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   165
            TabIndex        =   18
            Top             =   165
            Width           =   750
         End
      End
      Begin VB.TextBox age 
         Height          =   420
         Left            =   615
         TabIndex        =   11
         Top             =   1950
         Width           =   900
      End
      Begin VB.TextBox fathername 
         Height          =   375
         Left            =   2610
         TabIndex        =   10
         Top             =   525
         Width           =   4380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Age  &&  Date of "
         ForeColor       =   &H00C00000&
         Height          =   780
         Left            =   495
         TabIndex        =   12
         Top             =   1725
         Width           =   6870
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   5235
            TabIndex        =   16
            Top             =   330
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   37491
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2310
            TabIndex        =   15
            Top             =   315
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   37491
         End
         Begin VB.Label Label2 
            Caption         =   "Birth"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1335
            TabIndex        =   14
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Joining"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   4170
            TabIndex        =   13
            Top             =   375
            Width           =   780
         End
      End
      Begin VB.Label Label43 
         Caption         =   "total income"
         Height          =   465
         Left            =   -68040
         TabIndex        =   109
         Top             =   3660
         Width           =   1395
      End
      Begin VB.Label Label42 
         Caption         =   "Other"
         Height          =   465
         Left            =   -67890
         TabIndex        =   108
         Top             =   2715
         Width           =   1395
      End
      Begin VB.Label Label41 
         Caption         =   "City Allow"
         Height          =   300
         Left            =   -67935
         TabIndex        =   107
         Top             =   2310
         Width           =   1395
      End
      Begin VB.Label Label40 
         Caption         =   "Phone Allow"
         Height          =   465
         Left            =   -67905
         TabIndex        =   106
         Top             =   1605
         Width           =   1395
      End
      Begin VB.Label Label39 
         Caption         =   "Prof.Dev."
         Height          =   465
         Left            =   -67920
         TabIndex        =   105
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label38 
         Caption         =   "Fuel Allow"
         Height          =   465
         Left            =   -67905
         TabIndex        =   104
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label37 
         Caption         =   "MAZ.ALL"
         Height          =   465
         Left            =   -71085
         TabIndex        =   103
         Top             =   3915
         Width           =   1395
      End
      Begin VB.Label Label33 
         Caption         =   "CLOTH - 2"
         ForeColor       =   &H00C00000&
         Height          =   540
         Left            =   -73485
         TabIndex        =   88
         Top             =   3930
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "CLOTH - 1"
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   -73455
         TabIndex        =   87
         Top             =   3150
         Width           =   2295
      End
      Begin VB.Label Label31 
         Caption         =   "MARITAL"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9315
         TabIndex        =   83
         Top             =   1740
         Width           =   780
      End
      Begin VB.Label Label29 
         Caption         =   "R.D."
         ForeColor       =   &H00C00000&
         Height          =   585
         Left            =   -73470
         TabIndex        =   73
         Top             =   2265
         Width           =   2640
      End
      Begin VB.Label Label28 
         Caption         =   "L.I.C"
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   -73455
         TabIndex        =   72
         Top             =   1560
         Width           =   2205
      End
      Begin VB.Label Label27 
         Caption         =   "PF AMOUNT"
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   -73455
         TabIndex        =   70
         Top             =   780
         Width           =   2235
      End
      Begin VB.Label Label25 
         Caption         =   " L.T.A"
         Height          =   465
         Left            =   -71055
         TabIndex        =   68
         Top             =   3315
         Width           =   1395
      End
      Begin VB.Label Label24 
         Caption         =   "Washing Allow"
         Height          =   465
         Left            =   -71040
         TabIndex        =   55
         Top             =   2700
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "Attn. Allow"
         Height          =   465
         Left            =   -74685
         TabIndex        =   54
         Top             =   4080
         Width           =   1875
      End
      Begin VB.Label Label22 
         Caption         =   "Medical Allow"
         Height          =   465
         Left            =   -71025
         TabIndex        =   53
         Top             =   2070
         Width           =   1875
      End
      Begin VB.Label Label21 
         Caption         =   "Tea Allow"
         Height          =   465
         Left            =   -71025
         TabIndex        =   52
         Top             =   1605
         Width           =   1875
      End
      Begin VB.Label Label20 
         Caption         =   "Spl. Allow"
         Height          =   465
         Left            =   -71010
         TabIndex        =   51
         Top             =   1065
         Width           =   1875
      End
      Begin VB.Label Label19 
         Caption         =   "Conv. Allow"
         Height          =   465
         Left            =   -71010
         TabIndex        =   50
         Top             =   510
         Width           =   1440
      End
      Begin VB.Label Label18 
         Caption         =   "House Rent Allow"
         Height          =   465
         Left            =   -74670
         TabIndex        =   49
         Top             =   3345
         Width           =   1875
      End
      Begin VB.Label Label17 
         Caption         =   "Variable DA"
         Height          =   465
         Left            =   -74625
         TabIndex        =   48
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Label Label16 
         Caption         =   "Fixed DA"
         Height          =   465
         Left            =   -74610
         TabIndex        =   47
         Top             =   2130
         Width           =   1875
      End
      Begin VB.Label Label14 
         Caption         =   "Special Pay"
         Height          =   510
         Left            =   -74610
         TabIndex        =   46
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Service Weightage"
         Height          =   390
         Left            =   -74610
         TabIndex        =   45
         Top             =   1020
         Width           =   2340
      End
      Begin VB.Label Label12 
         Caption         =   "Basic Pay"
         Height          =   480
         Left            =   -74610
         TabIndex        =   44
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label15 
         Height          =   435
         Left            =   -74415
         TabIndex        =   43
         Top             =   2835
         Width           =   1515
      End
      Begin VB.Label DESI 
         Caption         =   "DESIGNATION"
         Height          =   255
         Left            =   -73770
         TabIndex        =   41
         Top             =   1275
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "STAFF / WORKER"
         Height          =   375
         Left            =   -73785
         TabIndex        =   40
         Top             =   2385
         Width           =   1845
      End
      Begin VB.Label Label10 
         Caption         =   "WORKING IN"
         Height          =   450
         Left            =   -73695
         TabIndex        =   36
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "DEPARTMENT "
         Height          =   480
         Left            =   -73710
         TabIndex        =   35
         Top             =   645
         Width           =   2085
      End
      Begin VB.Label Label5 
         Caption         =   "BLOOD GROUP"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   7770
         TabIndex        =   28
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label father 
         Caption         =   "Father's Name"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   9
         Top             =   570
         Width           =   1995
      End
   End
   Begin VB.TextBox emp_code 
      Height          =   420
      Left            =   4770
      TabIndex        =   0
      Top             =   825
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save "
      Height          =   645
      Left            =   4425
      TabIndex        =   2
      Top             =   7755
      Width           =   1260
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   645
      Left            =   1410
      Top             =   7875
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox emp_name 
      Height          =   435
      Left            =   4770
      TabIndex        =   1
      Top             =   1710
      Width           =   5235
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1995
      Left            =   720
      TabIndex        =   3
      Top             =   420
      Width           =   10425
      Begin VB.Label empname 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   435
         TabIndex        =   5
         Top             =   1305
         Width           =   3165
      End
      Begin VB.Label empcode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   435
         TabIndex        =   4
         Top             =   420
         Width           =   2805
      End
   End
   Begin VB.Label netpay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "NET PAY (Rs.)"
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
      Height          =   465
      Left            =   6660
      TabIndex        =   7
      Top             =   7935
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim paydb As New ADODB.Connection
'Dim payrs As New ADODB.Recordset
Private Sub Command1_Click()
'   Set paydb = New ADODB.Connection
'   Set payrs = New ADODB.Recordset
   paydb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
   sql = "Select * from emp"
   payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   payrs.AddNew
   payrs.Fields("Eno") = Text1.Text
   payrs.Fields("Ename") = Text2.Text
   payrs.Fields("e_add1") = Text3.Text
   payrs.Fields("e_pin") = Text6.Text
   payrs.Fields("basic") = Val(Text7.Text)
   payrs.Fields("sw") = Val(Text8.Text)
   payrs.Fields("spay") = Val(Text9.Text)
   payrs.Fields("da") = Val(Text10.Text)
   payrs.Fields("hra") = Val(Text11.Text)
   payrs.Fields("pf") = Val(Text13.Text)
   payrs.Fields("lic") = Val(Text14.Text)
   payrs.Fields("house") = Val(Text15.Text)
   payrs.Update
   payrs.Close
   Unload Me
End Sub

Private Sub Form_Load()
MALE.Value = True
'   Set paydb = New ADODB.Connection
'   Set payrs = New ADODB.Recordset
'   paydb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=anna_test;Data Source=servalldata"
'   sql = "Select * from emp"
'   payrs.Open sql, paydb, adOpenDynamic, adLockBatchOptimistic
End Sub


Private Sub Text7_Change()
 L.T.A
End Sub
