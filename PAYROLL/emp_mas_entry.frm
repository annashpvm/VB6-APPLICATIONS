VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form emp_mas_entry 
   BackColor       =   &H00C0E0FF&
   Caption         =   "EMPLOYEE DETAILS ENTRY"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   18810
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   2175
      Left            =   18480
      TabIndex        =   144
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
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
         Left            =   1800
         TabIndex        =   164
         Top             =   0
         Width           =   1965
      End
      Begin VB.ComboBox cmb_decholiday_eligbile_yn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3000
         TabIndex        =   159
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txt_mess_subsidy 
         Height          =   495
         Left            =   3000
         TabIndex        =   158
         Top             =   2160
         Width           =   1335
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
         Left            =   240
         TabIndex        =   150
         Top             =   960
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
         Left            =   1800
         TabIndex        =   149
         Top             =   960
         Width           =   2160
      End
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   148
         Top             =   240
         Visible         =   0   'False
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
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   145
         Top             =   480
         Visible         =   0   'False
         Width           =   1650
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
         Left            =   960
         TabIndex        =   165
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label74 
         Caption         =   "DH Add.Wages Eligible"
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
         Height          =   285
         Left            =   240
         TabIndex        =   161
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label Label75 
         Caption         =   "Mess Subsidy"
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
         Left            =   240
         TabIndex        =   160
         Top             =   2280
         Width           =   2445
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
         Left            =   120
         TabIndex        =   147
         Top             =   240
         Visible         =   0   'False
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
         Left            =   0
         TabIndex        =   146
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
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
      Left            =   16440
      TabIndex        =   139
      Top             =   360
      Visible         =   0   'False
      Width           =   2340
      Begin VB.OptionButton opt_staff 
         BackColor       =   &H00C0E0FF&
         Caption         =   "STAFF"
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton opt_worker 
         BackColor       =   &H00C0E0FF&
         Caption         =   "WORKER"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1080
         TabIndex        =   140
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   "&Move"
      Height          =   495
      Left            =   5760
      Picture         =   "emp_mas_entry.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   960
      TabIndex        =   81
      Top             =   8880
      Width           =   3615
      Begin VB.CommandButton Exit 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   2640
         Picture         =   "emp_mas_entry.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_save 
         Caption         =   "&Save "
         Height          =   735
         Left            =   1800
         Picture         =   "emp_mas_entry.frx":05CC
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton emp_edit 
         Caption         =   "&Edit"
         Height          =   735
         Left            =   960
         Picture         =   "emp_mas_entry.frx":0A0E
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton NEW 
         Caption         =   "&New"
         Height          =   735
         Left            =   120
         Picture         =   "emp_mas_entry.frx":1078
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         Width           =   855
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
      Left            =   8880
      TabIndex        =   35
      Top             =   9000
      Width           =   1965
   End
   Begin TabDlg.SSTab z 
      Height          =   5835
      Left            =   360
      TabIndex        =   39
      Top             =   2760
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   8
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
      TabPicture(0)   =   "emp_mas_entry.frx":16E2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "SEXFRAME"
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(4)=   "cmb_blood"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(6)=   "Frame5"
      Tab(0).Control(7)=   "Frame7"
      Tab(0).Control(8)=   "cmdr"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "PF and DEPT. DETAILS"
      TabPicture(1)   =   "emp_mas_entry.frx":16FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_Year_Passed"
      Tab(1).Control(1)=   "weekly_off_lst"
      Tab(1).Control(2)=   "qualify_cmb"
      Tab(1).Control(3)=   "desi_cmb"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "emptype_cmb"
      Tab(1).Control(6)=   "work_cmb"
      Tab(1).Control(7)=   "dept_cmb"
      Tab(1).Control(8)=   "Label94"
      Tab(1).Control(9)=   "week_off"
      Tab(1).Control(10)=   "Label44"
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(12)=   "DESI"
      Tab(1).Control(13)=   "Label11"
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(15)=   "Label9"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "EARNINGS"
      TabPicture(2)   =   "emp_mas_entry.frx":171A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label18"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label25"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label42"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label43"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label82"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label83"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label85"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Basic"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "fda"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "hra"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "medall"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "othall"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Gross"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "lta"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txt_grosspay"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txt_pfsalary"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txt_esisalary"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Frame11"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "STANDARD DEDUCTIONS"
      TabPicture(3)   =   "emp_mas_entry.frx":1736
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label27"
      Tab(3).Control(1)=   "Label28"
      Tab(3).Control(2)=   "Label29"
      Tab(3).Control(3)=   "pfpercentage"
      Tab(3).Control(4)=   "Label80"
      Tab(3).Control(5)=   "Label84"
      Tab(3).Control(6)=   "pfamt"
      Tab(3).Control(7)=   "lic"
      Tab(3).Control(8)=   "rd"
      Tab(3).Control(9)=   "txt_itded"
      Tab(3).Control(10)=   "txt_esiamt"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "BANK ACCOUNT DETAILS"
      TabPicture(4)   =   "emp_mas_entry.frx":1752
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label54"
      Tab(4).Control(1)=   "Label55"
      Tab(4).Control(2)=   "Label56"
      Tab(4).Control(3)=   "lbl_bank_ifsc"
      Tab(4).Control(4)=   "cmb_bank"
      Tab(4).Control(5)=   "txt_bank_acno"
      Tab(4).Control(6)=   "txt_email"
      Tab(4).Control(7)=   "txt_bank_ifsc"
      Tab(4).Control(8)=   "cmd_ref_bank"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "EMPLOYEE STATUS"
      TabPicture(5)   =   "emp_mas_entry.frx":176E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label49"
      Tab(5).Control(1)=   "cmb_empstatus"
      Tab(5).Control(2)=   "frame_resigned"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "OTHERS"
      TabPicture(6)   =   "emp_mas_entry.frx":178A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label64"
      Tab(6).Control(1)=   "Label65"
      Tab(6).Control(2)=   "Label66"
      Tab(6).Control(3)=   "Label67"
      Tab(6).Control(4)=   "Label68"
      Tab(6).Control(5)=   "Label69"
      Tab(6).Control(6)=   "Label72"
      Tab(6).Control(7)=   "Label73"
      Tab(6).Control(8)=   "txt_appointedby"
      Tab(6).Control(9)=   "txt_preinterviewby"
      Tab(6).Control(10)=   "txt_interviewername"
      Tab(6).Control(11)=   "txt_oe"
      Tab(6).Control(12)=   "txt_ie"
      Tab(6).Control(13)=   "cmb_pi_eligbile_yn"
      Tab(6).Control(14)=   "cmb_work_hrs"
      Tab(6).ControlCount=   15
      TabCaption(7)   =   "INCREMENT DETAILS"
      TabPicture(7)   =   "emp_mas_entry.frx":17A6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txtIncrement"
      Tab(7).Control(1)=   "flx_increment"
      Tab(7).Control(2)=   "cmd_Add"
      Tab(7).Control(3)=   "cmbYear"
      Tab(7).Control(4)=   "cmbMonth"
      Tab(7).Control(5)=   "txt_Start_Salary"
      Tab(7).Control(6)=   "cmd_Save"
      Tab(7).Control(7)=   "Label93"
      Tab(7).Control(8)=   "Label92"
      Tab(7).Control(9)=   "Label91"
      Tab(7).Control(10)=   "Label90"
      Tab(7).ControlCount=   11
      Begin VB.Frame Frame11 
         Caption         =   "Other Benifits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   7800
         TabIndex        =   191
         Top             =   1080
         Width           =   3495
         Begin VB.CommandButton cmd_Save2 
            Caption         =   "SAVE"
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
            TabIndex        =   198
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txt_other_benefit 
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
            Left            =   2040
            TabIndex        =   196
            Top             =   1440
            Width           =   1185
         End
         Begin VB.TextBox txt_Petrol 
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
            Left            =   2040
            TabIndex        =   194
            Top             =   960
            Width           =   1185
         End
         Begin VB.TextBox txt_houseRent 
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
            Left            =   2040
            TabIndex        =   192
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label txtAny_Others 
            Caption         =   "Any Ohers"
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
            Height          =   360
            Left            =   240
            TabIndex        =   197
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label96 
            Caption         =   "Petrol"
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
            Height          =   360
            Left            =   240
            TabIndex        =   195
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label95 
            Caption         =   "House Rent"
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
            Height          =   360
            Left            =   240
            TabIndex        =   193
            Top             =   600
            Width           =   1245
         End
      End
      Begin VB.TextBox txt_Year_Passed 
         Height          =   285
         Left            =   -72480
         TabIndex        =   190
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtIncrement 
         Height          =   375
         Left            =   -65880
         TabIndex        =   186
         Top             =   1560
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid flx_increment 
         Height          =   2775
         Left            =   -74160
         TabIndex        =   185
         Top             =   2160
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4895
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Add 
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
         Left            =   -64320
         TabIndex        =   184
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   -69120
         TabIndex        =   183
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   -72600
         TabIndex        =   181
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txt_Start_Salary 
         Height          =   495
         Left            =   -72600
         TabIndex        =   178
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdr 
         Caption         =   "R"
         Height          =   195
         Left            =   -64680
         TabIndex        =   175
         Top             =   2730
         Width           =   255
      End
      Begin VB.ComboBox weekly_off_lst 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   172
         Top             =   4170
         Width           =   4035
      End
      Begin VB.TextBox txt_esisalary 
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
         Left            =   1680
         TabIndex        =   168
         Top             =   2370
         Width           =   1545
      End
      Begin VB.TextBox txt_esiamt 
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
         TabIndex        =   166
         Top             =   2010
         Width           =   1650
      End
      Begin VB.TextBox txt_pfsalary 
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
         Left            =   1680
         TabIndex        =   162
         Top             =   1770
         Width           =   1545
      End
      Begin VB.TextBox txt_grosspay 
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
         Left            =   1680
         TabIndex        =   156
         Top             =   1170
         Width           =   1545
      End
      Begin VB.TextBox txt_itded 
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
         Left            =   -72000
         TabIndex        =   142
         Top             =   2610
         Width           =   1695
      End
      Begin VB.ComboBox cmb_work_hrs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67560
         TabIndex        =   135
         Top             =   4290
         Width           =   1215
      End
      Begin VB.ComboBox cmb_pi_eligbile_yn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -71760
         TabIndex        =   133
         Top             =   4170
         Width           =   1215
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
         Left            =   -65640
         TabIndex        =   129
         Top             =   1290
         Width           =   255
      End
      Begin VB.TextBox txt_bank_ifsc 
         Height          =   450
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   127
         Top             =   2490
         Width           =   4605
      End
      Begin VB.TextBox txt_ie 
         Height          =   615
         Left            =   -69480
         TabIndex        =   124
         Top             =   3210
         Width           =   1455
      End
      Begin VB.TextBox txt_oe 
         Height          =   615
         Left            =   -71640
         TabIndex        =   123
         Top             =   3210
         Width           =   1455
      End
      Begin VB.TextBox txt_interviewername 
         Height          =   495
         Left            =   -71640
         TabIndex        =   118
         Text            =   " "
         Top             =   1170
         Width           =   5415
      End
      Begin VB.TextBox txt_preinterviewby 
         Height          =   495
         Left            =   -71640
         TabIndex        =   117
         Text            =   " "
         Top             =   1890
         Width           =   5415
      End
      Begin VB.TextBox txt_appointedby 
         Height          =   495
         Left            =   -71640
         TabIndex        =   116
         Text            =   " "
         Top             =   2610
         Width           =   5415
      End
      Begin VB.Frame frame_resigned 
         Height          =   2175
         Left            =   -74520
         TabIndex        =   99
         Top             =   1890
         Visible         =   0   'False
         Width           =   9135
         Begin VB.TextBox txt_reason 
            Height          =   285
            Left            =   2640
            TabIndex        =   177
            Top             =   1560
            Width           =   6015
         End
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
            TabIndex        =   103
            Top             =   840
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker dt_resigned 
            Height          =   315
            Left            =   2640
            TabIndex        =   100
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
            Format          =   130023425
            CurrentDate     =   37491
         End
         Begin VB.Label Label89 
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
            TabIndex        =   176
            Top             =   1560
            Width           =   1935
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
            TabIndex        =   102
            Top             =   960
            Width           =   2415
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
            TabIndex        =   101
            Top             =   360
            Width           =   2175
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
         ItemData        =   "emp_mas_entry.frx":17C2
         Left            =   -71880
         List            =   "emp_mas_entry.frx":17C4
         TabIndex        =   97
         Text            =   "cmb_empstatus"
         Top             =   1410
         Width           =   3255
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relationship"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   -74520
         TabIndex        =   92
         Top             =   930
         Width           =   8145
         Begin VB.TextBox fathername 
            Height          =   375
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   95
            Top             =   480
            Width           =   6180
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
            TabIndex        =   94
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
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
            TabIndex        =   93
            Top             =   480
            Width           =   1395
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
            TabIndex        =   96
            Top             =   240
            Width           =   6195
         End
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
         TabIndex        =   90
         Top             =   3090
         Width           =   4605
      End
      Begin VB.TextBox txt_bank_acno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -71640
         MaxLength       =   20
         TabIndex        =   89
         Top             =   1890
         Width           =   4605
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
         ItemData        =   "emp_mas_entry.frx":17C6
         Left            =   -71640
         List            =   "emp_mas_entry.frx":17C8
         TabIndex        =   86
         Top             =   1290
         Width           =   5895
      End
      Begin VB.ComboBox qualify_cmb 
         Height          =   315
         Left            =   -72480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   840
         Width           =   4050
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
         Left            =   5880
         TabIndex        =   29
         Top             =   3210
         Width           =   1545
      End
      Begin VB.ComboBox desi_cmb 
         Height          =   315
         Left            =   -72480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2250
         Width           =   4065
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
         Left            =   -69000
         TabIndex        =   74
         Top             =   2490
         Width           =   2760
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
            TabIndex        =   9
            Top             =   225
            Width           =   1110
         End
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
            TabIndex        =   8
            Top             =   255
            Value           =   -1  'True
            Width           =   720
         End
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
         Left            =   9720
         TabIndex        =   31
         Top             =   4920
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
         Height          =   465
         Left            =   5880
         TabIndex        =   30
         Top             =   4650
         Width           =   1515
      End
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   -74535
         TabIndex        =   68
         Top             =   1770
         Width           =   9825
         Begin VB.ComboBox caste_cmb 
            Height          =   315
            Left            =   7305
            TabIndex        =   5
            Top             =   210
            Width           =   2415
         End
         Begin VB.ComboBox Community_cmb 
            Height          =   315
            Left            =   5130
            TabIndex        =   4
            Top             =   195
            Width           =   1470
         End
         Begin VB.ComboBox Religion_cmb 
            Height          =   315
            Left            =   1275
            TabIndex        =   3
            Top             =   180
            Width           =   2805
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
            TabIndex        =   71
            Top             =   270
            Width           =   615
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
            TabIndex        =   70
            Top             =   225
            Width           =   705
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
            TabIndex        =   69
            Top             =   270
            Width           =   915
         End
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
         Height          =   4395
         Left            =   -68280
         TabIndex        =   65
         Top             =   1170
         Width           =   4635
         Begin VB.ComboBox cmb_pf_eligible 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   600
            Width           =   1905
         End
         Begin VB.ComboBox cmb_esi_eligible 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   1080
            Width           =   1905
         End
         Begin VB.TextBox txt_uan 
            Height          =   330
            Left            =   2160
            MaxLength       =   12
            TabIndex        =   131
            Top             =   3120
            Width           =   2085
         End
         Begin VB.TextBox txt_esino 
            Height          =   330
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   106
            Top             =   3600
            Width           =   2085
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
            Height          =   195
            Left            =   3840
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   4140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.TextBox pfno 
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   24
            Top             =   2640
            Width           =   2085
         End
         Begin VB.TextBox PF 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   23
            Top             =   1680
            Width           =   810
         End
         Begin MSComCtl2.DTPicker dt_pf_join 
            Height          =   300
            Left            =   2160
            TabIndex        =   105
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   130023425
            CurrentDate     =   37491
         End
         Begin VB.Label Label50 
            Caption         =   "ESI ELIGIBLE (Y/N)"
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
            Left            =   240
            TabIndex        =   155
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label81 
            Caption         =   "PF ELIGIBLE (Y/N)"
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
            Left            =   240
            TabIndex        =   154
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label59 
            Caption         =   "PF Join Date"
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
            Height          =   405
            Left            =   240
            TabIndex        =   151
            Top             =   2160
            Width           =   1305
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
            Height          =   285
            Left            =   240
            TabIndex        =   130
            Top             =   3120
            Width           =   585
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
            Left            =   240
            TabIndex        =   107
            Top             =   3600
            Width           =   1305
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
            Left            =   240
            TabIndex        =   67
            Top             =   2760
            Width           =   1305
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
            Left            =   240
            TabIndex        =   66
            Top             =   1680
            Width           =   780
         End
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
         TabIndex        =   34
         Top             =   3990
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
         TabIndex        =   33
         Top             =   3330
         Width           =   1650
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
         TabIndex        =   32
         Top             =   1350
         Width           =   1650
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
         Left            =   5880
         TabIndex        =   28
         Top             =   3930
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
         Left            =   5880
         TabIndex        =   27
         Top             =   2610
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
         Left            =   5880
         TabIndex        =   26
         Top             =   1890
         Width           =   1545
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
         Left            =   5880
         TabIndex        =   25
         Top             =   1170
         Width           =   1545
      End
      Begin VB.ComboBox emptype_cmb 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3450
         Width           =   4035
      End
      Begin VB.ComboBox work_cmb 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2850
         Width           =   4065
      End
      Begin VB.ComboBox dept_cmb 
         Height          =   315
         Left            =   -72480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1770
         Width           =   4080
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
         TabIndex        =   10
         Top             =   2730
         Width           =   1155
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   -73920
         TabIndex        =   46
         Top             =   3330
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
         TabPicture(0)   =   "emp_mas_entry.frx":17CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label63"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "c_add1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "c_add2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "c_add3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "c_pin"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txt_phoneno"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "PERMENANT ADDRESS"
         TabPicture(1)   =   "emp_mas_entry.frx":17E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "Label7"
         Tab(1).Control(2)=   "p_add1"
         Tab(1).Control(3)=   "p_add3"
         Tab(1).Control(4)=   "p_add2"
         Tab(1).Control(5)=   "p_pin"
         Tab(1).Control(6)=   "chk"
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
            Left            =   -73800
            TabIndex        =   115
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txt_phoneno 
            Height          =   375
            Left            =   5760
            MaxLength       =   25
            TabIndex        =   113
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox p_pin 
            Height          =   345
            Left            =   -72660
            MaxLength       =   7
            TabIndex        =   18
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox p_add2 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1200
            Width           =   5895
         End
         Begin VB.TextBox p_add3 
            Height          =   345
            Left            =   -72675
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1560
            Width           =   5895
         End
         Begin VB.TextBox p_add1 
            Height          =   345
            Left            =   -72660
            MaxLength       =   50
            TabIndex        =   15
            Top             =   855
            Width           =   5895
         End
         Begin VB.TextBox c_pin 
            Height          =   375
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   14
            Top             =   1830
            Width           =   1815
         End
         Begin VB.TextBox c_add3 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1380
            Width           =   5895
         End
         Begin VB.TextBox c_add2 
            Height          =   375
            Left            =   2310
            MaxLength       =   50
            TabIndex        =   12
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox c_add1 
            Height          =   375
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   11
            Top             =   585
            Width           =   5895
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
            TabIndex        =   114
            Top             =   1920
            Width           =   1200
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
            TabIndex        =   51
            Top             =   1920
            Width           =   1020
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
            TabIndex        =   50
            Top             =   840
            Width           =   885
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
            TabIndex        =   48
            Top             =   1830
            Width           =   1560
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
            TabIndex        =   47
            Top             =   735
            Width           =   1695
         End
      End
      Begin VB.Frame SEXFRAME 
         Caption         =   "GENDER"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   -66120
         TabIndex        =   44
         Top             =   930
         Width           =   1425
         Begin VB.OptionButton FEMALE 
            Caption         =   "FEMALE"
            ForeColor       =   &H00800000&
            Height          =   480
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1155
         End
         Begin VB.OptionButton MALE 
            Caption         =   "MALE"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   750
         End
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
         Left            =   -74520
         TabIndex        =   41
         Top             =   2490
         Width           =   5205
         Begin MSComCtl2.DTPicker doj 
            Height          =   300
            Left            =   3300
            TabIndex        =   7
            Top             =   345
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   130023425
            CurrentDate     =   37491
         End
         Begin MSComCtl2.DTPicker dob 
            Height          =   315
            Left            =   855
            TabIndex        =   6
            Top             =   300
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Format          =   130023425
            CurrentDate     =   37491
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
            TabIndex        =   43
            Top             =   375
            Width           =   990
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
            TabIndex        =   42
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.CommandButton cmd_Save 
         Caption         =   "SAVE"
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
         Left            =   -64320
         TabIndex        =   188
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label94 
         Caption         =   "YEAR OF PASSED"
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
         Left            =   -74760
         TabIndex        =   189
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label93 
         Caption         =   "Increment"
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
         Left            =   -66960
         TabIndex        =   187
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label92 
         Caption         =   "Year of Revision"
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
         Left            =   -70680
         TabIndex        =   182
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label91 
         Caption         =   "Month of Revision"
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
         Left            =   -74520
         TabIndex        =   180
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label90 
         Caption         =   "Starting Salary"
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
         Left            =   -74520
         TabIndex        =   179
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label85 
         Caption         =   "ESI Salary"
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
         Height          =   360
         Left            =   360
         TabIndex        =   169
         Top             =   2490
         Width           =   1245
      End
      Begin VB.Label Label84 
         Caption         =   "ESI AMOUNT"
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
         Left            =   -74520
         TabIndex        =   167
         Top             =   2070
         Width           =   2235
      End
      Begin VB.Label Label83 
         Caption         =   "PF Salary"
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
         Height          =   360
         Left            =   360
         TabIndex        =   163
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label82 
         Caption         =   "Gross Pay"
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
         Height          =   360
         Left            =   360
         TabIndex        =   157
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label Label80 
         Caption         =   "IT  DEDUCTION"
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
         Left            =   -74520
         TabIndex        =   143
         Top             =   2610
         Width           =   2055
      End
      Begin VB.Label Label73 
         Caption         =   "Working Hrs (8 to 16)"
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
         Height          =   285
         Left            =   -69720
         TabIndex        =   134
         Top             =   4290
         Width           =   2205
      End
      Begin VB.Label Label72 
         Caption         =   "OT hours Eligible"
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
         Height          =   285
         Left            =   -74520
         TabIndex        =   132
         Top             =   4290
         Width           =   1845
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
         TabIndex        =   128
         Top             =   2610
         Width           =   1575
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
         Left            =   -69960
         TabIndex        =   126
         Top             =   3330
         Width           =   525
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
         Left            =   -72120
         TabIndex        =   125
         Top             =   3330
         Width           =   525
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
         TabIndex        =   122
         Top             =   3330
         Width           =   2565
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
         TabIndex        =   121
         Top             =   1170
         Width           =   1605
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
         TabIndex        =   120
         Top             =   1890
         Width           =   2685
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
         TabIndex        =   119
         Top             =   2730
         Width           =   2685
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
         TabIndex        =   98
         Top             =   1410
         Width           =   2535
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
         TabIndex        =   91
         Top             =   3210
         Width           =   1575
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
         TabIndex        =   88
         Top             =   2010
         Width           =   1575
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
         TabIndex        =   87
         Top             =   1290
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
         TabIndex        =   80
         Top             =   1470
         Width           =   975
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
         Left            =   -74760
         TabIndex        =   79
         Top             =   4230
         Width           =   2175
      End
      Begin VB.Label Label44 
         Caption         =   "Label44"
         Height          =   30
         Left            =   -73320
         TabIndex        =   78
         Top             =   3015
         Width           =   30
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
         Left            =   -74760
         TabIndex        =   77
         Top             =   840
         Width           =   2175
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
         Left            =   7680
         TabIndex        =   73
         Top             =   5040
         Width           =   1995
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
         Left            =   4320
         TabIndex        =   72
         Top             =   4650
         Width           =   1395
      End
      Begin VB.Label Label29 
         Caption         =   "RD"
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
         Left            =   -74520
         TabIndex        =   64
         Top             =   4050
         Width           =   1395
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
         Left            =   -74520
         TabIndex        =   63
         Top             =   3330
         Width           =   1515
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
         Left            =   -74520
         TabIndex        =   62
         Top             =   1410
         Width           =   2235
      End
      Begin VB.Label Label25 
         Caption         =   "T.A"
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
         Left            =   4200
         TabIndex        =   61
         Top             =   3330
         Width           =   1395
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
         Left            =   4200
         TabIndex        =   60
         Top             =   4050
         Width           =   1395
      End
      Begin VB.Label Label18 
         Caption         =   "HRA"
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
         Left            =   4200
         TabIndex        =   59
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "DA"
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
         Left            =   4200
         TabIndex        =   58
         Top             =   2010
         Width           =   1365
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
         Height          =   360
         Left            =   4200
         TabIndex        =   57
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label Label15 
         Height          =   435
         Left            =   585
         TabIndex        =   56
         Top             =   3570
         Width           =   1515
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
         Left            =   -74760
         TabIndex        =   55
         Top             =   2370
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
         Left            =   -74760
         TabIndex        =   54
         Top             =   3450
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
         Left            =   -74760
         TabIndex        =   53
         Top             =   2850
         Width           =   2175
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
         Left            =   -74760
         TabIndex        =   52
         Top             =   1830
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "BLOOD GROUP"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -65940
         TabIndex        =   49
         Top             =   2490
         Width           =   1275
      End
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
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   645
      Left            =   270
      Top             =   6480
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
   Begin VB.TextBox emp_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
      Width           =   6075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2700
      Left            =   360
      TabIndex        =   36
      Top             =   0
      Width           =   11505
      Begin VB.TextBox txt_location 
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
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   173
         Top             =   2160
         Width           =   3495
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
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   171
         Top             =   360
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
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   138
         Top             =   1560
         Width           =   5415
      End
      Begin VB.ComboBox CMB_EMPCODE 
         Height          =   315
         Left            =   9120
         TabIndex        =   136
         Text            =   "Combo1"
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   8400
         TabIndex        =   108
         Top             =   120
         Width           =   2820
         Begin VB.OptionButton opt_resigned 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Resigned"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1680
            TabIndex        =   111
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_Active 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Active"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton opt_All 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ALL"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   960
            TabIndex        =   109
            Top             =   240
            Width           =   615
         End
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
         Left            =   2520
         TabIndex        =   75
         Top             =   960
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.Label Label87 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Location"
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
         Left            =   240
         TabIndex        =   174
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label86 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search Code"
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
         Left            =   4440
         TabIndex        =   170
         Top             =   360
         Width           =   1245
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
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Top             =   1680
         Width           =   2055
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
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1965
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
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1605
      End
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
      Left            =   7440
      TabIndex        =   40
      Top             =   9120
      Width           =   1455
   End
End
Attribute VB_Name = "emp_mas_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim paydb As New ADODB.Connection
Dim payrs As New ADODB.Recordset
Dim swmr As String
Dim searchopt As Integer
Dim fin_selrow As Integer
Dim glx_flag As String
Dim eligible_pfsalary, eligible_esisalary, esi_percentage, pf_percentage As Double

Function fillgrid()
    With flx_increment
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "Month"
        .TextMatrix(0, 1) = "Month"
        .TextMatrix(0, 2) = "Year"
        .TextMatrix(0, 3) = "Increment"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
    End With
End Function

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
''Private Sub ca_KeyPress(KeyAscii As Integer)
''  find_Grosspay
''  find_netpay
''  On Error GoTo err_handler
''    chk_keyascii ca, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

''Private Sub attall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''
'' On Error GoTo err_handler
''    chk_keyascii attall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub Basic_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
find_Grosspay
find_netpay
    chk_keyascii Basic, "N", 6, 2, KeyAscii
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub



Private Sub caste_cmb_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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

Private Sub cityall_Change()
  find_Grosspay
  find_netpay
End Sub

''Private Sub cityall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii cityall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub cmb_empstatus_Click()
       If cmb_empstatus.Text = "RESIGNED" Then
          frame_resigned.Visible = True
       Else
          frame_resigned.Visible = False
       End If
End Sub

Private Sub cmb_esi_eligible_Change()
    txt_grosspay_Change
End Sub

Private Sub cmb_esi_eligible_Click()
    txt_grosspay_Change
End Sub

Private Sub cmb_pf_eligible_Click()
   If cmb_pf_eligible.Text = "YES" Then
      PF.Enabled = True
      pfno.Enabled = True
      PF.Text = pf_percentage
    Else
      PF.Enabled = False
      pfno.Enabled = False
      PF.Text = ""
    End If
    txt_grosspay_Change
End Sub

Private Sub cmd_Add_Click()
    fin_selrow = flx_increment.Rows
    
    With flx_increment
         .TextMatrix(fin_selrow - 1, 0) = cmbMonth.Text
         .TextMatrix(fin_selrow - 1, 1) = Val(cmbMonth.ItemData(cmbMonth.ListIndex))
         .TextMatrix(fin_selrow - 1, 2) = cmbYear.Text
         .TextMatrix(fin_selrow - 1, 3) = Int(txtIncrement.Text)
    End With
    flx_increment.Rows = flx_increment.Rows + 1
       
End Sub

Private Sub cmd_getpf_Click()
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = "select max(emp_pfno)+1 as eno from emp_mas  where emp_company = '" & company_code & "'"
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
        payrs.MoveFirst
        While Not payrs.EOF
            pfno.Text = payrs("eno")
            payrs.MoveNext
        Wend
    End If
''    Set paydb = vbNullString
''    Set payrs = vbNullString
    
End Sub



Private Sub cmd_move_Click()
    
    
Dim adocmd_new As New ADODB.Command
Dim adocmd_old As New ADODB.Command


Dim adors_new As New ADODB.Recordset
Dim adors_old As New ADODB.Recordset


    Dim pst_qry As String
    
    
    paycopy = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=shvpm1;Data Source=10.0.0.252"
    
    Set gen_connection = New ADODB.Connection
    gen_connection.CursorLocation = adUseClient
    gen_connection.Open pay
    
    Set gen_connection_new = New ADODB.Connection
    gen_connection_new.CursorLocation = adUseClient
    gen_connection_new.Open paycopy

    
    
''    pst_qry = "delete from emp_mas"
''    adocmd_new.ActiveConnection = gen_connection_new
''
''    adocmd_new.CommandType = adCmdText
''    adocmd_new.CommandText = pst_qry
''    Set adors_new = adocmd_new.Execute
    


    pst_qry = "delete from emp_salary"
    adocmd_new.ActiveConnection = gen_connection_new

    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute


    
    pst_qry = "delete from bio_devicelogs"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from bio_attendlogs"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    

    
    pst_qry = "delete from bio_device_shiftlogs"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    
    pst_qry = "delete from bio_emp_permissions"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    
    pst_qry = "delete from bio_emp_oddetails"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    
    pst_qry = "delete from bio_empleave"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from bio_emp_chleave"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    
     pst_qry = "delete from bio_empmas"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
       
       
       
    pst_qry = "delete from emp_dec_holiday"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
       
    pst_qry = "delete from canteen_recovery"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from canteen_expenses"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from monthly_deduction"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
              


''    pst_qry = "select * from emp_mas where emp_pfeligible = 'Y'"
''    adocmd_old.ActiveConnection = gen_connection
''    adocmd_old.CommandType = adCmdText
''    adocmd_old.CommandText = pst_qry
''    Set adors_old = adocmd_old.Execute
''
''    If adors_old.RecordCount > 0 Then
''        For pin_cnt = 1 To adors_old.RecordCount
''            pst_qry = "insert into emp_mas   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Val(adors_old(2)) & "','" & adors_old(3) & "','" & adors_old(4) & "','" & adors_old(5) & "','" & adors_old(6) & "'," _
''            & " '" & adors_old(7) & "','" & Val(adors_old(8)) & "','" & Val(adors_old(9)) & "','" & Val(adors_old(10)) & "','" & adors_old(11) & "','" & adors_old(12) & "','" & adors_old(13) & "','" & adors_old(14) & "','" & adors_old(15) & "','" & adors_old(16) & "','" & adors_old(17) & "','" & adors_old(18) & "','" & adors_old(19) & "','" & Format(adors_old(20), "MM/dd/yyyy") & "','" & Format(adors_old(21), "MM/dd/yyyy") & "','" & Val(adors_old(22)) & "','" & Val(adors_old(23)) & "','" & Val(adors_old(24)) & "','" & adors_old(25) & "','" & adors_old(26) & "','" & Val(adors_old(27)) & "','" & Val(adors_old(28)) & "','" & adors_old(29) & "','" & adors_old(30) & "','" & adors_old(31) & "','" & Val(adors_old(32)) & "','" & Val(adors_old(33)) & "','" & Val(adors_old(34)) & "','" & Val(adors_old(35)) & "','" & Val(adors_old(36)) & "','" & Val(adors_old(37)) & "'," _
''            & " '" & Val(adors_old(38)) & "','" & Val(adors_old(39)) & "','" & Val(adors_old(40)) & "','" & Val(adors_old(41)) & "','" & Val(adors_old(42)) & "','" & Val(adors_old(43)) & "','" & Val(adors_old(44)) & "','" & Val(adors_old(45)) & "','" & Val(adors_old(46)) & "','" & Val(adors_old(47)) & "','" & Val(adors_old(48)) & "'," _
''            & " '" & Val(adors_old(49)) & "','" & Val(adors_old(50)) & "','" & Val(adors_old(51)) & "','" & Val(adors_old(52)) & "','" & Val(adors_old(53)) & "','" & Val(adors_old(54)) & "','" & Val(adors_old(55)) & "','" & Val(adors_old(56)) & "','" & Val(adors_old(57)) & "','" & adors_old(58) & "','" & adors_old(59) & "','" & adors_old(60) & "','" & adors_old(61) & "','" & Val(adors_old(62)) & "','" & adors_old(63) & "','" & adors_old(64) & "','" & adors_old(65) & "','" & Format(adors_old(66), "MM/dd/yyyy") & "','" & adors_old(67) & "','" & Format(adors_old(68), "MM/dd/yyyy") & "' , " _
''            & " '" & adors_old(69) & "','" & Val(adors_old(70)) & "','" & Val(adors_old(71)) & "','" & Val(adors_old(72)) & "','" & adors_old(73) & "','" & adors_old(74) & "','" & adors_old(75) & "','" & adors_old(76) & "','" & adors_old(77) & "','" & adors_old(78) & "','" & Val(adors_old(79)) & "','" & adors_old(80) & "','" & adors_old(81) & "','" & adors_old(82) & "')"
''
''            adocmd_new.ActiveConnection = gen_connection_new
''            adocmd_new.CommandType = adCmdText
''            adocmd_new.CommandText = pst_qry
''            adocmd_new.Execute
''            adors_old.MoveNext
''        Next pin_cnt
''    End If



'''' emp_salary
''        pst_qry = "select * from emp_salary where s_pf > 0"
''    adocmd_old.ActiveConnection = gen_connection
''''    adocmd_old.CommandType = adCmdText
''    adocmd_old.CommandText = pst_qry
''    Set adors_old = adocmd_old.Execute
''
''    If adors_old.RecordCount > 0 Then
''        For pin_cnt = 1 To adors_old.RecordCount
''            pst_qry = "insert into emp_salary   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Val(adors_old(2)) & "','" & Val(adors_old(3)) & "','" & adors_old(4) & "','" & adors_old(5) & "','" & adors_old(6) & "'," _
''            & " '" & Val(adors_old(7)) & "','" & Val(adors_old(8)) & "','" & Val(adors_old(9)) & "','" & Val(adors_old(10)) & "','" & Val(adors_old(11)) & "','" & Val(adors_old(12)) & "','" & Val(adors_old(13)) & "','" & Val(adors_old(14)) & "','" & Val(adors_old(15)) & "','" & Val(adors_old(16)) & "','" & Val(adors_old(17)) & "','" & Val(adors_old(18)) & "','" & Val(adors_old(19)) & "','" & Val(adors_old(20)) & "','" & Val(adors_old(21)) & "','" & Val(adors_old(22)) & "','" & Val(adors_old(23)) & "','" & Val(adors_old(24)) & "','" & Val(adors_old(25)) & "','" & Val(adors_old(26)) & "','" & Val(adors_old(27)) & "','" & Val(adors_old(28)) & "','" & Val(adors_old(29)) & "','" & Val(adors_old(30)) & "','" & Val(adors_old(31)) & "','" & Val(adors_old(32)) & "','" & Val(adors_old(33)) & "','" & Val(adors_old(34)) & "','" & Val(adors_old(35)) & "','" & Val(adors_old(36)) & "','" & Val(adors_old(37)) & "'," _
''            & " '" & Val(adors_old(38)) & "','" & Val(adors_old(39)) & "','" & Val(adors_old(40)) & "','" & Val(adors_old(41)) & "','" & Val(adors_old(42)) & "','" & Val(adors_old(43)) & "','" & Val(adors_old(44)) & "','" & Val(adors_old(45)) & "','" & Val(adors_old(46)) & "','" & Val(adors_old(47)) & "','" & Val(adors_old(48)) & "'," _
''            & " '" & Val(adors_old(49)) & "','" & Val(adors_old(50)) & "','" & Val(adors_old(51)) & "','" & Val(adors_old(52)) & "','" & Val(adors_old(53)) & "','" & Val(adors_old(54)) & "','" & Val(adors_old(55)) & "','" & Val(adors_old(56)) & "','" & Val(adors_old(57)) & "','" & Val(adors_old(58)) & "','" & Val(adors_old(59)) & "','" & Val(adors_old(60)) & "','" & Val(adors_old(61)) & "','" & Val(adors_old(62)) & "','" & Val(adors_old(63)) & "','" & Val(adors_old(64)) & "','" & Val(adors_old(65)) & "','" & Val(adors_old(66)) & "','" & Val(adors_old(67)) & "','" & Val(adors_old(68)) & "' , " _
''            & " '" & Val(adors_old(69)) & "','" & Val(adors_old(70)) & "','" & Val(adors_old(71)) & "','" & Val(adors_old(72)) & "','" & Val(adors_old(73)) & "','" & Val(adors_old(74)) & "','" & Val(adors_old(75)) & "','" & Val(adors_old(76)) & "','" & Val(adors_old(77)) & "','" & Val(adors_old(78)) & "')"
''
''            adocmd_new.ActiveConnection = gen_connection_new
''            adocmd_new.CommandType = adCmdText
''            adocmd_new.CommandText = pst_qry
''            adocmd_new.Execute
''            adors_old.MoveNext
''        Next pin_cnt
''    End If


    pst_qry = "select * from bio_devicelogs , emp_mas where ad_fpcode = emp_fpcode and  emp_pfeligible = 'Y'"
    adocmd_old.ActiveConnection = gen_connection
''    adocmd_old.CommandType = adCmdText
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_devicelogs   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Val(adors_old(2)) & "','" & Format(adors_old(3), "MM/dd/yyyy") & "','" & Format(adors_old(4), "MM/dd/yyyy") & "' , '" & adors_old(5) & "', '" & adors_old(6) & "', '" & adors_old(7) & "')"

            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If



     pst_qry = "select * from bio_attendlogs , emp_mas where a_fpcode = emp_fpcode and  emp_pfeligible = 'Y'"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            data113 = 0
            If IsNull(adors_old(113)) Then
               data1113 = 0
           Else
               data1113 = adors_old(113)
           End If
           
            pst_qry = "insert into bio_attendlogs   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Val(adors_old(2)) & "','" & Val(adors_old(3)) & "','" & adors_old(4) & "','" & Format(adors_old(5), "MM/dd/yyyy") & "','" & Format(adors_old(6), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(7) & "','" & Format(adors_old(8), "MM/dd/yyyy") & "','" & Format(adors_old(9), "MM/dd/yyyy") & "','" & adors_old(10) & "','" & Format(adors_old(11), "MM/dd/yyyy") & "','" & Format(adors_old(12), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(13) & "','" & Format(adors_old(14), "MM/dd/yyyy") & "','" & Format(adors_old(15), "MM/dd/yyyy") & "','" & adors_old(16) & "','" & Format(adors_old(17), "MM/dd/yyyy") & "','" & Format(adors_old(18), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(19) & "','" & Format(adors_old(20), "MM/dd/yyyy") & "','" & Format(adors_old(21), "MM/dd/yyyy") & "', '" & adors_old(22) & "','" & Format(adors_old(23), "MM/dd/yyyy") & "','" & Format(adors_old(24), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(25) & "','" & Format(adors_old(26), "MM/dd/yyyy") & "','" & Format(adors_old(27), "MM/dd/yyyy") & "', '" & adors_old(28) & "','" & Format(adors_old(29), "MM/dd/yyyy") & "','" & Format(adors_old(30), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(31) & "','" & Format(adors_old(32), "MM/dd/yyyy") & "','" & Format(adors_old(33), "MM/dd/yyyy") & "', '" & adors_old(34) & "','" & Format(adors_old(35), "MM/dd/yyyy") & "','" & Format(adors_old(36), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(37) & "','" & Format(adors_old(38), "MM/dd/yyyy") & "','" & Format(adors_old(39), "MM/dd/yyyy") & "', '" & adors_old(40) & "','" & Format(adors_old(41), "MM/dd/yyyy") & "','" & Format(adors_old(42), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(43) & "','" & Format(adors_old(44), "MM/dd/yyyy") & "','" & Format(adors_old(45), "MM/dd/yyyy") & "', '" & adors_old(46) & "','" & Format(adors_old(47), "MM/dd/yyyy") & "','" & Format(adors_old(48), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(49) & "','" & Format(adors_old(50), "MM/dd/yyyy") & "','" & Format(adors_old(51), "MM/dd/yyyy") & "', '" & adors_old(52) & "','" & Format(adors_old(53), "MM/dd/yyyy") & "','" & Format(adors_old(54), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(55) & "','" & Format(adors_old(56), "MM/dd/yyyy") & "','" & Format(adors_old(57), "MM/dd/yyyy") & "', '" & adors_old(58) & "','" & Format(adors_old(59), "MM/dd/yyyy") & "','" & Format(adors_old(60), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(61) & "','" & Format(adors_old(62), "MM/dd/yyyy") & "','" & Format(adors_old(63), "MM/dd/yyyy") & "', '" & adors_old(64) & "','" & Format(adors_old(65), "MM/dd/yyyy") & "','" & Format(adors_old(66), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(67) & "','" & Format(adors_old(68), "MM/dd/yyyy") & "','" & Format(adors_old(69), "MM/dd/yyyy") & "', '" & adors_old(70) & "','" & Format(adors_old(71), "MM/dd/yyyy") & "','" & Format(adors_old(72), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(73) & "','" & Format(adors_old(74), "MM/dd/yyyy") & "','" & Format(adors_old(75), "MM/dd/yyyy") & "', '" & adors_old(76) & "','" & Format(adors_old(77), "MM/dd/yyyy") & "','" & Format(adors_old(78), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(79) & "','" & Format(adors_old(80), "MM/dd/yyyy") & "','" & Format(adors_old(81), "MM/dd/yyyy") & "', '" & adors_old(82) & "','" & Format(adors_old(83), "MM/dd/yyyy") & "','" & Format(adors_old(84), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(85) & "','" & Format(adors_old(86), "MM/dd/yyyy") & "','" & Format(adors_old(87), "MM/dd/yyyy") & "', '" & adors_old(88) & "','" & Format(adors_old(89), "MM/dd/yyyy") & "','" & Format(adors_old(90), "MM/dd/yyyy") & "', " _
            & "'" & adors_old(91) & "','" & Format(adors_old(92), "MM/dd/yyyy") & "','" & Format(adors_old(93), "MM/dd/yyyy") & "', '" & adors_old(94) & "','" & Format(adors_old(95), "MM/dd/yyyy") & "','" & Format(adors_old(96), "MM/dd/yyyy") & "', " _
            & "'" & Val(adors_old(97)) & "','" & Val(adors_old(98)) & "','" & Val(adors_old(99)) & "','" & Val(adors_old(100)) & "', '" & Val(adors_old(101)) & "','" & Val(adors_old(102)) & "','" & Val(adors_old(103)) & "','" & Val(adors_old(104)) & "', " _
            & "'" & Val(adors_old(105)) & "','" & Val(adors_old(106)) & "','" & Val(adors_old(107)) & "','" & Val(adors_old(108)) & "','" & Val(adors_old(109)) & "','" & Val(adors_old(110)) & "','" & Val(adors_old(111)) & "','" & Val(adors_old(112)) & "', " _
            & "'" & data1113 & "','" & Val(adors_old(114)) & "','" & Val(adors_old(115)) & "')"

            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If


     pst_qry = "select * from bio_device_shiftlogs  , emp_mas where ds_fpcode = emp_fpcode and  emp_pfeligible = 'Y'"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    Dim data12 As Integer
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            data12 = 0
            If IsNull(adors_old(12)) Then
               data12 = 0
           Else
               data12 = adors_old(12)
           End If

            pst_qry = "insert into bio_device_shiftlogs   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Val(adors_old(2)) & "','" & Val(adors_old(3)) & "','" & Format(adors_old(4), "MM/dd/yyyy") & "' , " _
            & "'" & adors_old(5) & "', '" & adors_old(6) & "','" & adors_old(7) & "', '" & adors_old(8) & "','" & adors_old(9) & "', '" & adors_old(10) & "','" & adors_old(11) & "','" & Val(data12) & "'," _
            & "'" & Format(adors_old(13), "MM/dd/yyyy") & "','" & Format(adors_old(14), "MM/dd/yyyy") & "' ," _
            & "'" & Format(adors_old(15), "MM/dd/yyyy") & "','" & Format(adors_old(16), "MM/dd/yyyy") & "' ," _
            & "'" & Format(adors_old(17), "MM/dd/yyyy") & "','" & Format(adors_old(18), "MM/dd/yyyy") & "' ," _
            & "'" & Format(adors_old(19), "MM/dd/yyyy") & "','" & Format(adors_old(20), "MM/dd/yyyy") & "' ," _
            & "'" & Format(adors_old(21), "MM/dd/yyyy") & "','" & Format(adors_old(22), "MM/dd/yyyy") & "' ," _
            & "'" & Format(adors_old(23), "MM/dd/yyyy") & "','" & Format(adors_old(24), "MM/dd/yyyy") & "' ," _
            & "'" & adors_old(25) & "','" & Val(adors_old(26)) & "','" & Val(adors_old(27)) & "','" & Val(adors_old(28)) & "'," _
            & "'" & Val(adors_old(29)) & "','" & Val(adors_old(30)) & "','" & Val(adors_old(31)) & "','" & Val(adors_old(32)) & "','" & Val(adors_old(33)) & "', '" & Val(adors_old(34)) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
    
    
    
    
    
     pst_qry = "select * from bio_emp_permissions   , emp_mas where empp_fpcode = emp_fpcode and  emp_pfeligible = 'Y'"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_emp_permissions   values ( '" & Val(adors_old(0)) & "','" & Val(adors_old(1)) & "','" & Format(adors_old(2), "MM/dd/yyyy") & "' , " _
            & "'" & adors_old(3) & "', '" & Val(adors_old(4)) & "','" & Val(adors_old(5)) & "', '" & adors_old(6) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If


     pst_qry = "select * from bio_emp_oddetails   , emp_mas where empod_fpcode = emp_fpcode and  emp_pfeligible = 'Y' "
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_emp_oddetails   values ( '" & Val(adors_old(0)) & "','" & Format(adors_old(1), "MM/dd/yyyy") & "' ,'" & Val(adors_old(2)) & "'," _
            & "'" & Format(adors_old(3), "MM/dd/yyyy") & "', '" & Val(adors_old(4)) & "','" & Val(adors_old(5)) & "' , '" & adors_old(6) & "', '" & adors_old(7) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If

    pst_qry = "select * from bio_empleave a   , emp_mas b where a.emp_fpcode = b.emp_fpcode and  b.emp_pfeligible = 'Y' "
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_empleave   values ( '" & Val(adors_old(0)) & "','" & Format(adors_old(1), "MM/dd/yyyy") & "' ,'" & Val(adors_old(2)) & "', " _
            & "'" & adors_old(3) & "', '" & Format(adors_old(4), "MM/dd/yyyy") & "','" & adors_old(5) & "', '" & adors_old(6) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If

    pst_qry = "select * from bio_emp_chleave  a   , emp_mas b where a.empch_fpcode = b.emp_fpcode and  b.emp_pfeligible = 'Y' "
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute

    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_emp_chleave   values ( '" & Val(adors_old(0)) & "','" & Format(adors_old(1), "MM/dd/yyyy") & "' ,'" & Val(adors_old(2)) & "', " _
            & "'" & Format(adors_old(3), "MM/dd/yyyy") & "','" & adors_old(4) & "','" & Format(adors_old(5), "MM/dd/yyyy") & "', '" & adors_old(6) & "', '" & adors_old(7) & "', '" & adors_old(8) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If

    pst_qry = "select * from bio_empmas   a   , emp_mas b where a.bioemp_fpcode = b.emp_fpcode and  b.emp_pfeligible = 'Y' "
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount

            pst_qry = "insert into bio_empmas   values ( '" & adors_old(0) & "','" & Val(adors_old(1)) & "' ,'" & Val(adors_old(2)) & "', " _
            & "'" & adors_old(3) & "','" & adors_old(4) & "','" & adors_old(5) & "', '" & adors_old(6) & "', '" & adors_old(7) & "', '" & Val(adors_old(8)) & "', '" & Val(adors_old(9)) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
                


    pst_qry = "select * from emp_dec_holiday"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            pst_qry = "insert into emp_dec_holiday   values ( '" & Format(adors_old(0), "MM/dd/yyyy") & "' ,'" & adors_old(1) & "' )"

            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
                          

    pst_qry = "delete from canteen_recovery"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from canteen_expenses"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
    
    pst_qry = "delete from monthly_deduction"
    adocmd_new.ActiveConnection = gen_connection_new
    adocmd_new.CommandType = adCmdText
    adocmd_new.CommandText = pst_qry
    Set adors_new = adocmd_new.Execute
        
        
    pst_qry = "select * from canteen_recovery"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            pst_qry = "insert into canteen_recovery   values ( '" & Format(adors_old(0), "MM/dd/yyyy") & "' ,'" & adors_old(1) & "'  ,'" & adors_old(2) & "' ,'" & adors_old(3) & "'  ,'" & adors_old(4) & "'   )"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
                      
    pst_qry = "select * from canteen_expenses"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            pst_qry = "insert into canteen_expenses   values ( '" & Format(adors_old(0), "MM/dd/yyyy") & "' ,'" & adors_old(1) & "'  ,'" & adors_old(2) & "' ,'" & adors_old(3) & "'  ,'" & adors_old(4) & "' ,'" & adors_old(5) & "','" & adors_old(6) & "' ,'" & adors_old(7) & "','" & adors_old(8) & "','" & adors_old(9) & "' ,'" & adors_old(10) & "' ,'" & adors_old(11) & "','" & adors_old(12) & "','" & adors_old(13) & "','" & adors_old(14) & "','" & adors_old(15) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
                      
                      
     pst_qry = "select * from monthly_deduction"
    adocmd_old.ActiveConnection = gen_connection
    adocmd_old.CommandText = pst_qry
    Set adors_old = adocmd_old.Execute
    
    If adors_old.RecordCount > 0 Then
        For pin_cnt = 1 To adors_old.RecordCount
            pst_qry = "insert into monthly_deduction   values ( '" & adors_old(0) & "' ,'" & adors_old(1) & "'  ,'" & adors_old(2) & "' ,'" & adors_old(3) & "'  ,'" & adors_old(4) & "' ,'" & adors_old(5) & "','" & adors_old(6) & "' ,'" & adors_old(7) & "','" & adors_old(8) & "')"
            adocmd_new.ActiveConnection = gen_connection_new
            adocmd_new.CommandType = adCmdText
            adocmd_new.CommandText = pst_qry
            adocmd_new.Execute
            adors_old.MoveNext
        Next pin_cnt
    End If
    
 ''   adors_old.Close
  ''  adors_new.Close


    MsgBox ("Moved sucessfully..")
    
    
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

Private Sub cmd_Save_Click()
    
      sql = "delete from emp_increment where empi_idcode = '" & emp_idcode.Text & "'  and empi_millcode = '" & company_code & "'"
      gcn_shvpm.Execute sql
      
      For i = 1 To flx_increment.Rows - 1
         If Val(flx_increment.TextMatrix(i, 3)) > 0 Then
            sql = "insert into emp_increment values(" & company_code & " ," & emp_idcode.Text & " ,'" & flx_increment.TextMatrix(i, 0) & "'," & flx_increment.TextMatrix(i, 1) & "," & flx_increment.TextMatrix(i, 2) & "," & Val(flx_increment.TextMatrix(i, 3)) & ")"
            gcn_shvpm.Execute sql
         End If
      Next
      
      sql = "update emp_mas set emp_passed_year = '" & txt_Year_Passed.Text & "'   , emp_starting_salary = '" & Val(txt_Start_Salary.Text) & "' where emp_code = " & emp_idcode.Text & " and emp_company = " & company_code & ""
      gcn_shvpm.Execute sql
       
      MsgBox ("Record Saved..")

      
      
End Sub



Private Sub Combo1_Change()

End Sub


Private Sub cmd_Save2_Click()

   
      sql = "update emp_mas set emp_houserent_allow = '" & txt_houseRent.Text & "' , emp_petrol_allow = '" & Val(txt_Petrol.Text) & "'  , emp_other_allow = '" & Val(txt_other_benefit.Text) & "'  where emp_code = " & emp_idcode.Text & " and emp_company = " & company_code & ""
      gcn_shvpm.Execute sql
       
      MsgBox ("Record Saved..")
      
   txt_houseRent.Text = ""
   txt_Petrol.Text = ""
   txt_other_benefit.Text = ""

End Sub

Private Sub cmdr_Click()
    If poweruser = 1 Then
    cmd_move.Visible = True
 Else
    cmd_move.Visible = False
 End If
End Sub

Private Sub Community_cmb_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub eduall_Change()
  find_Grosspay
  find_netpay
End Sub

Private Sub emp_idcode_Change()
  If searchopt = 1 Then Exit Sub

End Sub
Private Sub emp_idcode_LostFocus()
  find_empcode
End Sub

Private Sub emp_idcode_KeyPress(KeyAscii As Integer)
  On Error GoTo err_handler
    If KeyAscii <> 8 Then chk_keyascii fda, "N", 5, 0, KeyAscii
 ''   find_empcode
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Function find_empcode()
    Set paydb2 = New ADODB.Connection
    Set payrs2 = New ADODB.Recordset
    paydb2.Open pay
    sql = "select * from emp_mas where emp_code = '" & Trim(emp_idcode.Text) & "' and emp_company = '" & company_code & "'"
    payrs2.Open sql, paydb2, adOpenDynamic, adLockOptimistic
    If Not payrs2.EOF Then
       MsgBox ("Employee code already avaiable")
       emp_idcode.Text = ""
    End If
End Function



Function search_empcode()


End Function
Private Sub emp_edit_Click()
     savechk = 1
''    emp_idcode.Enabled = False
''    EDIT_FRAME.Visible = True
    empedit_cmb.Visible = True
    opt_staff.Enabled = True
    opt_emp_Click
    
''    opt_staff.SetFocus
''    savechk = 1
''    If opt_staff.Value = True Then
''       opt_staff_Click
''    Else
''       opt_worker_Click
''    End If
End Sub



Private Sub emp_name_Change()
    If searchopt = 1 Then Exit Sub
    If emp_idcode.Text = "" Then
       MsgBox ("Enter Employee code & Continue ...")
       emp_name.Text = ""
       Exit Sub
    End If
End Sub

Private Sub flx_increment_DblClick()
fin_selrow = flx_increment.Row
pst_ans = MsgBox("Press YES-to Delete NO-to Cancel", vbYesNo, "Confirmation")
If pst_ans = vbYes Then
    If flx_increment.Rows <= 2 Then
       flx_increment.Rows = 1
       MsgBox "No rows to remove"
    Else
       flx_increment.RemoveItem fin_selrow
       flx_increment.Row = flx_increment.Rows - 1
    End If
End If
End Sub

Private Sub healthall_Change()
  find_Grosspay
  find_netpay
End Sub


Private Sub opt_emp_Click()
    emp_idcode.Text = ""
    empedit_cmb.Clear
    If savechk = 1 Then
        emp_name.Visible = False
        ''emp_idcode.Enabled = False
        empedit_cmb.Visible = True
        Set paydb = New ADODB.Connection
        Set payrs = New ADODB.Recordset
        If opt_Active.Value = True Then
            sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_status = 'A'  order by emp_name"
        ElseIf opt_resigned.Value = True Then
            sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_status = 'R'  order by emp_name"
        Else
            sql = "select * from emp_mas  where emp_company = '" & company_code & "' order by emp_name"
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


Private Sub opt_staff_Click()
    emp_idcode.Text = ""
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
    emp_idcode.Text = ""
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

Private Sub emp_save_Click()
   Dim ecat As String
   If Trim(emp_idcode) = "" Then
      MsgBox ("Employee ID code is blank ")
     '' emp_idcode.SetFocus
      Exit Sub
   End If
   If Trim(emp_name) = "" Then
      MsgBox ("Employee Name is blank - correct it ")
      emp_name.SetFocus
      Exit Sub
   End If
   If Trim(txt_aadhaar.Text) = "" Then
      MsgBox ("Employee Aadhaar Number is blank ")
      txt_aadhaar.SetFocus
      Exit Sub
   End If
   If Trim(txt_location.Text) = "" Then
      MsgBox ("Employee's Location is blank ")
      txt_location.SetFocus
      Exit Sub
   End If
''   If Trim(cmb_blood.Text) = "" Then
''      MsgBox ("Blood type is blank ")
''      cmb_blood.SetFocus
''      Exit Sub
''   End If

''
''   If Trim(cmb_cost) = "" Then
''      MsgBox ("Employee Cost type is blank - correct it ")
''      cmb_cost.SetFocus
''      Exit Sub
''   End If
''
''   If Trim(cmb_group) = "" Then
''      MsgBox ("Employee Group type is blank - correct it ")
''      cmb_group.SetFocus
''      Exit Sub
''   End If
''
''
''
   If Trim(qualify_cmb) = "" Then
      MsgBox ("Employee Qualification is blank - correct it ")
      qualify_cmb.SetFocus
      Exit Sub
   End If
''
''   If Trim(fathername) = "" Then
''      MsgBox ("Employee father name Name is blank - correct it ")
''      fathername.SetFocus
''      Exit Sub
''   End If
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
''   If Trim(cmb_mc.Text) = "" And data_source <> "H" Then
''      MsgBox ("Select Machine.. ")
''      cmb_mc.SetFocus
''      Exit Sub
''   End If
   

   If cmb_pf_eligible.Text = "YES" And Val(pfno.Text) = 0 Then
      MsgBox ("PF Number is Nil... check it..")
      PF.SetFocus
      Exit Sub
   End If
''   If txt_interviewername.Text = "" Then
''      MsgBox "Enter Interviewer name"
''      txt_interviewername.SetFocus
''      Exit Sub
''   End If
''
''   If txt_preinterviewby.Text = "" Then
''      MsgBox "Enter Preliminary Interviewer name"
''      txt_preinterviewby.SetFocus
''      Exit Sub
''   End If
''
''   If txt_appointedby.Text = "" Then
''      MsgBox "Enter Appointed by name"
''      txt_appointedby.SetFocus
''      Exit Sub
''   End If
''
   
   
   etype = "W"
   If emptype_cmb.Text = "STAFF" Then
      etype = "S"
   End If
   
   Set paydb = New ADODB.Connection
   Set payrs = New ADODB.Recordset
   find_Grosspay
   paydb.Open pay
   If savechk = 0 Then
      
      sql = "select * from emp_mas where emp_code = '" & emp_idcode.Text & "' and emp_company = '" & company_code & "'"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If Not payrs.EOF Then
         MsgBox ("Employee code Already Entered for ... " + payrs("emp_name"))
         payrs.Close
         paydb.Close
         Exit Sub
      End If
      payrs.Close
      sql = "select * from emp_mas where emp_Name = '" & Trim(UCase(emp_name.Text)) & "' and emp_company = '" & company_code & "'"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      If Not payrs.EOF Then
         MsgBox ("Employee Name Already Entered  ... " + payrs("emp_name"))
         payrs.Close
         paydb.Close
         Exit Sub
      End If
      
      
      payrs.Close
      sql = "Select * from emp_mas"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      payrs.AddNew
      payrs.Fields("emp_name") = Trim(UCase(emp_name.Text))
   Else
      sql = ("select * from emp_mas where emp_code = '" & emp_idcode & "' and emp_company = '" & company_code & "'")
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      payrs.Fields("emp_name") = Trim(UCase(emp_name.Text))
   End If
   payrs.Fields("emp_company") = company_code
   payrs.Fields("emp_code") = emp_idcode.Text
   payrs.Fields("emp_fpcode") = emp_idcode.Text
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
   payrs.Fields("emp_cadd1") = UCase(Trim(c_add1))
   payrs.Fields("emp_cadd2") = UCase(Trim(c_add2))
   payrs.Fields("emp_cadd3") = UCase(Trim(c_add3))
   payrs.Fields("emp_cpin") = c_pin
   payrs.Fields("emp_contactno") = txt_phoneno.Text
   payrs.Fields("emp_padd1") = UCase(Trim(p_add1))
   payrs.Fields("emp_padd2") = UCase(Trim(p_add2))
   payrs.Fields("emp_padd3") = UCase(Trim(p_add3))
   payrs.Fields("emp_ppin") = p_pin
      If emptype_cmb.Text = "STAFF" Then
         payrs.Fields("emp_cat") = "S"
         ecat = "S"
      Else
         payrs.Fields("emp_cat") = "W"
         ecat = "W"
      End If
      
   If savechk = 0 Then
      payrs.Fields("emp_dept") = dept_cmb.ItemData(dept_cmb.ListIndex)
      payrs.Fields("emp_design") = desi_cmb.ItemData(desi_cmb.ListIndex)
      payrs.Fields("emp_type") = etype
      payrs.Fields("emp_qualify") = qualify_cmb.ItemData(qualify_cmb.ListIndex)

  Else
      find_deptcode (dept_cmb.Text)
      payrs.Fields("emp_dept") = dcode
      find_designcode (desi_cmb.Text)
      payrs.Fields("emp_design") = dcode
      ''find_typecode (emptype_cmb.Text)
      payrs.Fields("emp_type") = etype

      find_qualifycode (qualify_cmb.Text)
      payrs.Fields("emp_qualify") = dcode
   End If
   Dim wplace As String
   
   If work_cmb.Text = "MILL" Then
      payrs.Fields("emp_workplace") = "MIL"
      wplace = "MIL"
   Else
''   ElseIf work_cmb.Text = "COIMBATORE" Then
''      payrs.Fields("emp_workplace") = "CBE"
''      wplace = "CBE"
''   ElseIf work_cmb.Text = "SIVAKASI" Then
      payrs.Fields("emp_workplace") = "OTH"
      wplace = "OTH"
   End If
   If cmb_pf_eligible.Text = "YES" Then
      payrs.Fields("emp_pfeligible") = "Y"
      payrs.Fields("emp_pfjoin_date") = dt_pf_join.Value
   Else
      payrs.Fields("emp_pfeligible") = "N"
      payrs.Fields("emp_pfjoin_date") = 0
   End If
   payrs.Fields("emp_grosspay") = Val(txt_grosspay.Text)
   payrs.Fields("emp_pfsalary") = Val(txt_pfsalary.Text)
   payrs.Fields("emp_pfp") = Val(PF)
   payrs.Fields("emp_pfno") = Val(pfno)
   payrs.Fields("emp_uan") = txt_uan.Text
   payrs.Fields("emp_basic") = Val(Basic)
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
   payrs.Fields("emp_itded") = Val(txt_itded.Text)
   ''payrs.Fields("emp_fpcode") = Val(FPCODE)
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
   
   If cmb_esi_eligible.Text = "YES" Then
       payrs.Fields("emp_esieligible") = "Y"
   Else
       payrs.Fields("emp_esieligible") = "N"
   End If
   
   payrs.Fields("emp_bank") = cmb_bank.ItemData(cmb_bank.ListIndex)
   payrs.Fields("emp_bank_acno") = txt_bank_acno.Text
   payrs.Fields("emp_bank_ifsc") = txt_bank_ifsc.Text
   payrs.Fields("emp_email") = txt_email.Text
   
   payrs.Fields("emp_esino") = txt_esino.Text
   payrs.Fields("emp_resign_reason") = txt_reason.Text
   
   If cmb_empstatus.Text = "RESIGNED" Then
       payrs.Fields("emp_resigneddate") = dt_resigned.Value
       payrs.Fields("emp_reason") = cmb_reason.Text
   Else
       payrs.Fields("emp_resigneddate") = 0
       payrs.Fields("emp_reason") = ""
      
   End If
   payrs.Fields("emp_ctc") = ctc
   payrs.Fields("emp_interview_by") = txt_interviewername.Text
   payrs.Fields("emp_final_interview_by") = txt_preinterviewby.Text
   payrs.Fields("emp_appointed_by") = txt_appointedby.Text
   
   If cmb_pi_eligbile_yn.Text = "YES" Then
      payrs.Fields("emp_pi_eligible_yn") = "Y"
   Else
      payrs.Fields("emp_pi_eligible_yn") = "N"
   End If
   
   If cmb_decholiday_eligbile_yn.Text = "YES" Then
      payrs.Fields("emp_dh_wages_yn") = "Y"
   Else
      payrs.Fields("emp_dh_wages_yn") = "N"
   End If
   
   
   
   payrs.Fields("emp_work_hrs") = Val(cmb_work_hrs.Text)
   
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
   payrs.Fields("emp_aadhaar") = Left(Trim(txt_aadhaar.Text), 12)
   payrs.Fields("emp_location") = Left(Trim(txt_location.Text), 30)
   payrs.Fields("emp_costtype") = cmb_cost.Text
   
   payrs.Fields("emp_passed_year") = Left(Trim(txt_Year_Passed.Text), 10)
   payrs.Fields("emp_starting_salary") = Val(txt_Start_Salary.Text)
   
   payrs.Fields("emp_houserent_allow") = Val(txt_houseRent.Text)
   payrs.Fields("emp_petrol_allow") = Val(txt_Petrol.Text)
   payrs.Fields("emp_other_allow") = Val(txt_other_benefit.Text)
   
   
   payrs.Update
   payrs.Close
   MsgBox ("Data updated")
         
   If savechk = 0 Then
       sql = "insert into emp_eligible_leave values (" & company_code & "," & finyear & "," & year(doj.Value) & "," & emp_idcode.Text & ",0,'" & ecat & "','" & wplace & "' )"
       paydb.Execute sql
       
       sql = "insert into attn_entry values (" & company_code & "," & finyear & ",1," & year(doj.Value) & "," & emp_idcode.Text & ",'" & ecat & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
       paydb.Execute sql
   End If
   refresh_data
End Sub

Private Sub empedit_cmb_Click()
    fillgrid
    searchopt = 1
''    txt_newname.Text = ""
  ''  txt_father_name.Text = ""
    dt_resigned.Value = Format(Now, "dd/mm/yyyy")
    savechk = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_mas where emp_name = '" & Trim(empedit_cmb.Text) & "' and emp_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
       Exit Sub
    Else
''       While Not payrs.EOF
''           CMB_EMPCODE.AddItem payrs.Fields("emp_code")
''           payrs.MoveNext
''       Wend
       payrs.MoveFirst
   
   cmd_Save.Visible = True
   cmd_Save2.Visible = True
          
       swmr = payrs.Fields("emp_cat")
      
       emp_idcode = payrs.Fields("emp_code")
       emp_name = empedit_cmb.Text 'empedit_cmb.ItemData(empedit_cmb.ListIndex)
       

''       txt_newname.Text = empedit_cmb.Text
  ''     txt_father_name.Text = payrs.Fields("emp_fname")
       
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
       txt_aadhaar.Text = payrs.Fields("emp_aadhaar")
       cmb_cost.Text = payrs.Fields("emp_costtype")
       weekly_off_lst.Text = payrs.Fields("emp_holiday")
       cmb_blood.Text = payrs.Fields("emp_blood")
       Religion_cmb.Text = payrs.Fields("emp_religion")
       Community_cmb.Text = payrs.Fields("emp_community")
       caste_cmb.Text = payrs.Fields("emp_caste")
       c_add1.Text = payrs.Fields("emp_cadd1")
       c_add2.Text = payrs.Fields("emp_cadd2")
       c_add3.Text = payrs.Fields("emp_cadd3")
       c_pin.Text = payrs.Fields("emp_cpin")
       txt_location.Text = payrs.Fields("emp_location")
       txt_phoneno.Text = payrs.Fields("emp_contactno")
       p_add1.Text = payrs.Fields("emp_padd1")
       p_add2.Text = payrs.Fields("emp_padd2")
       p_add3.Text = payrs.Fields("emp_padd3")
       p_pin.Text = payrs.Fields("emp_ppin")
       
       txt_grosspay.Text = payrs.Fields("emp_grosspay")
       txt_pfsalary.Text = payrs.Fields("emp_pfsalary")
       
       txt_reason.Text = payrs.Fields("emp_resign_reason")
       
       Basic = payrs.Fields("emp_basic")
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
       
       
       txt_Year_Passed.Text = payrs.Fields("emp_passed_year")
       txt_Start_Salary.Text = payrs.Fields("emp_starting_salary")
       
       
       txt_pfsalary.Text = payrs.Fields("emp_pfsalary")
       txt_grosspay.Text = payrs.Fields("emp_grosspay")
       txt_itded.Text = payrs.Fields("emp_itded")
       
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
       rd = payrs.Fields("emp_rd")
       PF = payrs.Fields("emp_pfp")
       pfno = payrs.Fields("emp_pfno")
       txt_uan.Text = payrs.Fields("emp_uan")
       find_deptname (payrs.Fields("emp_dept"))
       dept_cmb.Text = dname
       
       desi_cmb.ListIndex = find_index_item_data(desi_cmb, payrs!emp_design)
       
       ''emptype_cmb.ListIndex = find_index_item_data(emptype_cmb, payrs!emp_type)
       
       If payrs!emp_cat = "S" Then
           emptype_cmb.Text = "STAFF"
       Else
           emptype_cmb.Text = "WORKER"
       End If
       
       qualify_cmb.ListIndex = find_index_item_data(qualify_cmb, payrs!emp_qualify)
       
       Religion_cmb.ListIndex = find_index_item_data(Religion_cmb, payrs!emp_religion)
       Community_cmb.ListIndex = find_index_item_data(Community_cmb, payrs!emp_community)
       caste_cmb.ListIndex = find_index_item_data(caste_cmb, payrs!emp_caste)
      
       dname = ""
       
       
       If payrs.Fields("emp_workplace") = "MIL" Then
          work_cmb.Text = "MILL"
       Else
          work_cmb.Text = "OTHER AREA"
       End If
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
       If payrs.Fields("emp_esieligible") = "Y" Then
          cmb_esi_eligible.Text = "YES"
       Else
          cmb_esi_eligible.Text = "NO"
       End If
       
       If payrs.Fields("emp_dh_wages_yn") = "Y" Then
          cmb_decholiday_eligbile_yn.Text = "YES"
       Else
          cmb_decholiday_eligbile_yn.Text = "NO"
       End If
       
       
       
       
       ''cmb_bank.ListIndex = payrs.Fields("emp_bank")
       cmb_bank.ListIndex = find_index_item_data(cmb_bank, payrs.Fields("emp_bank"))
       txt_bank_acno.Text = payrs.Fields("emp_bank_acno")
       txt_bank_ifsc.Text = payrs.Fields("emp_bank_ifsc")
       txt_email.Text = payrs.Fields("emp_email")
       txt_esino.Text = payrs.Fields("emp_esino")
       
   
       txt_houseRent.Text = payrs.Fields("emp_houserent_allow")
       txt_Petrol.Text = payrs.Fields("emp_petrol_allow")
       txt_other_benefit.Text = payrs.Fields("emp_other_allow")



    If Val(txt_pfsalary.Text) > 0 Then
   
      pfamt.Text = Round((Val(txt_pfsalary.Text) * Val(PF) / 100), 0)
  End If
       
       If payrs.Fields("emp_esieligible") = "Y" Then
          cmb_esi_eligible.Text = "YES"
       Else
          cmb_esi_eligible.Text = "NO"
       End If
       If payrs.Fields("emp_pfeligible") = "Y" Then
          cmb_pf_eligible.Text = "YES"
       Else
          cmb_pf_eligible.Text = "NO"
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
''    If PF_NONELIGIBLE.Value = True Then
''       cmd_getpf.Visible = True
''    Else
''       cmd_getpf.Visible = False
''    End If
''
    txt_interviewername.Text = payrs.Fields("emp_interview_by")
    txt_preinterviewby.Text = payrs.Fields("emp_final_interview_by")
    txt_appointedby.Text = payrs.Fields("emp_appointed_by")
    txt_oe.Text = payrs.Fields("emp_preexp_outside")
    txt_ie.Text = payrs.Fields("emp_preexp_inside")
    
    If payrs.Fields("emp_pi_eligible_yn") = "Y" Then
       cmb_pi_eligbile_yn.Text = "YES"
    Else
       cmb_pi_eligbile_yn.Text = "NO"
    End If
    
    cmb_work_hrs.Text = payrs.Fields("emp_work_hrs")
    
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
    

    i = 1
    sql = "select * from emp_increment where empi_idcode = " & Trim(emp_idcode.Text) & " and empi_millcode = '" & company_code & "'  order by empi_year, empi_month"
    payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs2.EOF
      With flx_increment
      .Rows = .Rows + 1
      flx_increment.TextMatrix(i, 0) = payrs2!empi_month
      flx_increment.TextMatrix(i, 1) = payrs2!empi_monthcode
      flx_increment.TextMatrix(i, 2) = payrs2!empi_year
      flx_increment.TextMatrix(i, 3) = payrs2!empi_increment
      

      i = i + 1
      End With
      payrs2.MoveNext
    Wend
    
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
   cmd_Save.Visible = False
   cmd_Save2.Visible = False
 glx_flag = "ADD"
 fillgrid
 With cmbMonth
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
    With cmbYear
        .AddItem finyear + 2000
        .AddItem "2012"
        .AddItem "2013"
        .AddItem "2014"
        .AddItem "2015"
        .AddItem "2016"
        .AddItem "2017"
        .AddItem "2018"
        .AddItem "2019"
        .AddItem "2020"
        .AddItem "2021"
        .AddItem "2022"
        .AddItem "2023"
        .AddItem "2024"
        .AddItem "2025"
        .AddItem "2026"

        .Text = "2025"
    End With



 If poweruser = 1 Then
    cmdr.Visible = True
 Else
    cmdr.Visible = False
 End If

 
 searchopt = 0
    swmr = "S"
      
    cmb_cost.AddItem "FIXED COST"
    cmb_cost.AddItem "VARIABLE COST"
    
    cmb_group.AddItem "SENIOR"
    cmb_group.AddItem "ESSENTIAL"
    cmb_group.AddItem "REGULAR"
    
    cmb_reason.AddItem "CESSATION"
    cmb_reason.AddItem "SUPERANNUATION"
    cmb_reason.AddItem "RETIREMENT"
    cmb_reason.AddItem "DEATH IN SERVICE"
    cmb_reason.AddItem "PERMANENT DISABLEMENT"

    ''emp_idcode.Enabled = False
    cmb_esi_eligible.AddItem "YES"
    cmb_esi_eligible.AddItem "NO"
    cmb_esi_eligible.Text = "NO"
    cmb_pf_eligible.AddItem "YES"
    cmb_pf_eligible.AddItem "NO"
    cmb_pf_eligible.Text = "NO"

    
    cmb_pi_eligbile_yn.Clear
    cmb_pi_eligbile_yn.AddItem "YES"
    cmb_pi_eligbile_yn.AddItem "NO"
    
    cmb_decholiday_eligbile_yn.AddItem "YES"
    cmb_decholiday_eligbile_yn.AddItem "NO"
    cmb_work_hrs.Clear
    cmb_work_hrs.AddItem "8.00"
    cmb_work_hrs.AddItem "9.00"
    cmb_work_hrs.AddItem "11.00"
    cmb_work_hrs.AddItem "12.00"
    cmb_work_hrs.AddItem "16.00"
     cmb_work_hrs.Text = "9.00"
''    If company_code = 1 Then
''       cmb_mc.AddItem "PM1"
''       cmb_mc.AddItem "PM2"
''       cmb_mc.AddItem "PM3"
''    ElseIf company_code = 3 Then
''       cmb_mc.AddItem "PM"
''    ElseIf company_code = 5 Then
''       cmb_mc.AddItem "POWER"
''    Else
''       cmb_mc.AddItem "OIL"
''    End If
    dob.Value = Now
    doj.Value = Now
    dt_resigned.Value = Now
    dob.Value = Format(Now, "dd/mm/yyyy")
    doj.Value = Format(Now, "dd/mm/yyyy")
    dt_resigned.Value = Format(Now, "dd/mm/yyyy")
    dt_pf_join.Value = Format(Now, "dd/mm/yyyy")
    
    emp_mas_entry.Caption = emp_mas_entry.Caption
    savechk = 0
''    pay = "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=spl_others;Data Source=spplserver"
    MALE.Value = True
    ''PF_ELIGIBLE.Value = True
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    sql = ("select * from comp_mas where comp_code = 1")
    paydb.Open pay
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    payrs.MoveFirst
    While Not payrs.EOF
    
        eligible_pfsalary = payrs("comp_pf_eligible")
        eligible_esisalary = payrs("comp_esi_eligible")
        pf_percentage = payrs("comp_pf_emp1_contri")
        esi_percentage = payrs("comp_esi_emp1_contri")
        payrs.MoveNext
    Wend
    
    
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
    
    
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    sql = ("Select * from  pemptype_mas order by dtype_code")
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    payrs.MoveFirst
''    While Not payrs.EOF
''        emptype_cmb.AddItem payrs(1)
''        emptype_cmb.ItemData(emptype_cmb.NewIndex) = payrs(0)
''        payrs.MoveNext
''    Wend
''
    
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
    
    work_cmb.AddItem "MILL"
    work_cmb.AddItem "OTHER AREA"
''    work_cmb.AddItem "COIMBATORE"
''    work_cmb.AddItem "CHENNAI"
''    work_cmb.AddItem "SIVAKASI"
    

    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset

        sql = "select * from emp_mas  where emp_company = '" & company_code & "' and emp_status = 'A'  order by emp_name"
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
        
    
  ''  PF_ELIGIBLE.Value = True
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
    
    emptype_cmb.AddItem ("STAFF")
    emptype_cmb.AddItem ("WORKER")
    
''    cmb_classification.AddItem "BELOW MANAGER"
''    cmb_classification.AddItem "MANAGER & ABOVE"
''    cmb_classification.AddItem "MANAGEMENT"
''    cmb_classification.Text = "BELOW MANAGER"
''    If data_source = "A" Then
''       loc = ""
''       loc2 = ""
''    ElseIf data_source = "H" Then
''       loc = " and emp_workplace = 'CBE'"
''       loc2 = " and s_workplace = 'CBE'"
''    Else
''       loc = " and emp_workplace = 'MILL'"
''       loc2 = " and s_workplace = 'MILL'"
''    End If
''----
    loc = ""
    loc2 = ""
'-----
End Sub

Private Sub fuelall_Change()
  find_Grosspay
  find_netpay
End Sub

''Private Sub fuelall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii fuelall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub houserent_Change()
  find_Grosspay
  find_netpay
End Sub

''Private Sub houserent_KeyPress(KeyAscii As Integer)
'' On Error GoTo err_handler
''    chk_keyascii houserent, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

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

''Private Sub mazall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii mazall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

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
    searchopt = 0
    savechk = 0
    emp_save.Enabled = True
    emp_idcode.Enabled = True
    emp_name.Visible = True
    empedit_cmb.Visible = False
    refresh_data
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''
''    sql = "select max(emp_code)+1 as eno from emp_mas  where emp_company = '" & company_code & "' "
''        paydb.Open pay
''        payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''        If Not payrs.EOF Then
''            payrs.MoveFirst
''
''            While Not payrs.EOF
''                emp_idcode.Text = payrs("eno")
''                payrs.MoveNext
''            Wend
''        End If
''        payrs.Close
         
''
''    Set paydb = New ADODB.Connection
''    Set payrs = New ADODB.Recordset
''    sql = "select max(emp_pfno)+1 as eno from emp_mas  where emp_company = '" & company_code & "'"
''    paydb.Open pay
''    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
''    If Not payrs.EOF Then
''        payrs.MoveFirst
''        empedit_cmb.Clear
''        While Not payrs.EOF
''            If IsNull(payrs("eno")) = True Then
''               pfno.Text = 1
''            Else
''               pfno.Text = payrs("eno")
''            End If
''            payrs.MoveNext
''        Wend
''    End If
    
    
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
   Dim pfcalc As Double
 
   If Val(txt_pfsalary.Text) > 0 Then
   
      pfamt.Text = Round((Val(txt_pfsalary.Text) * Val(PF) / 100), 0)
  End If
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
''Private Sub phoneall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii phoneall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

''Private Sub profall_Change()
''  find_Grosspay
''
''End Sub
''
''Private Sub profall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''
''On Error GoTo err_handler
''    chk_keyascii profall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

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

Private Sub Religion_cmb_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
''
''Private Sub ser_wt_Change()
'' find_Grosspay
'' find_netpay
''End Sub
''
''Private Sub ser_wt_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii ser_wt, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub
''
''Private Sub spl_pay_Change()
''  find_Grosspay
''  find_netpay
''End Sub
''
''Private Sub spl_pay_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii spl_pay, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub
''
''Private Sub splall_Change()
''  find_Grosspay
'' find_netpay
''End Sub
''
''
''
''
''Private Sub splall_KeyPress(KeyAscii As Integer)
''  find_Grosspay
''  find_netpay
''  On Error GoTo err_handler
''    chk_keyascii splall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

''Private Sub teaall_Change()
''  find_Grosspay
''  find_netpay
''End Sub
''
''Private Sub teaall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''  On Error GoTo err_handler
''    chk_keyascii teaall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub
''
''Private Sub txt_deposit_Change()
''  find_Grosspay
''  find_netpay
''End Sub
''
''Private Sub txt_deposit_KeyPress(KeyAscii As Integer)
''
'' On Error GoTo err_handler
''    chk_keyascii txt_deposit, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub txt_mess_Change()

End Sub


Private Sub txt_empcode_Change()
    fillgrid
    searchopt = 1
    
''    txt_newname.Text = ""
    dt_resigned.Value = Format(Now, "dd/mm/yyyy")
    savechk = 1
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    paydb.Open pay
    sql = ("select * from emp_mas where emp_code = '" & Trim(txt_empcode.Text) & "' and emp_company = '" & company_code & "'")
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
''       MsgBox ("Data not avaiable")
       Exit Sub
    Else
''       While Not payrs.EOF
''           CMB_EMPCODE.AddItem payrs.Fields("emp_code")
''           payrs.MoveNext
''       Wend

   cmd_Save.Visible = True
   cmd_Save2.Visible = True
   
       payrs.MoveFirst
       swmr = payrs.Fields("emp_cat")
       emp_idcode = payrs.Fields("emp_code")
       emp_name = payrs.Fields("emp_name")  ''empedit_cmb.Text   'empedit_cmb.ItemData(empedit_cmb.ListIndex)
       empedit_cmb.Text = payrs.Fields("emp_name")
''       txt_newname.Text = empedit_cmb.Text
''       txt_father_name.Text = payrs.Fields("emp_fname")
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
       txt_aadhaar.Text = payrs.Fields("emp_aadhaar")
       cmb_cost.Text = payrs.Fields("emp_costtype")
       weekly_off_lst.Text = payrs.Fields("emp_holiday")
       cmb_blood.Text = payrs.Fields("emp_blood")
       Religion_cmb.Text = payrs.Fields("emp_religion")
       Community_cmb.Text = payrs.Fields("emp_community")
       caste_cmb.Text = payrs.Fields("emp_caste")
       c_add1.Text = payrs.Fields("emp_cadd1")
       c_add2.Text = payrs.Fields("emp_cadd2")
       c_add3.Text = payrs.Fields("emp_cadd3")
       c_pin.Text = payrs.Fields("emp_cpin")
       txt_location.Text = payrs.Fields("emp_location")
       txt_phoneno.Text = payrs.Fields("emp_contactno")
       p_add1.Text = payrs.Fields("emp_padd1")
       p_add2.Text = payrs.Fields("emp_padd2")
       p_add3.Text = payrs.Fields("emp_padd3")
       p_pin.Text = payrs.Fields("emp_ppin")
       
       txt_grosspay.Text = payrs.Fields("emp_grosspay")
       txt_pfsalary.Text = payrs.Fields("emp_pfsalary")
       
       Basic = payrs.Fields("emp_basic")
     ''  ser_wt = payrs.Fields("emp_serwt")
    ''   spl_pay = payrs.Fields("emp_splpay")
       fda = payrs.Fields("emp_fda")
    ''   vda = payrs.Fields("emp_vda")
       hra = payrs.Fields("emp_hra")
''       attall = payrs.Fields("emp_attall")
''       ca = payrs.Fields("emp_convall")
''       splall = payrs.Fields("emp_splall")
''       teaall = payrs.Fields("emp_teaall")
       medall = payrs.Fields("emp_medall")
''       washall = payrs.Fields("emp_washall")
       lta = payrs.Fields("emp_lta")

       txt_itded.Text = payrs.Fields("emp_itded")
       
''       mazall = payrs.Fields("emp_magall")
''       fuelall = payrs.Fields("emp_fuelall")
''       profall = payrs.Fields("emp_profall")
''       cityall = payrs.Fields("emp_cityall")
''       phoneall = payrs.Fields("emp_phoneall")
''       healthall = payrs.Fields("emp_healthall")
   ''    FPCODE = payrs.Fields("emp_fpcode")
''       eduall = payrs.Fields("emp_eduall")
''       mealsall = payrs.Fields("emp_mealsall")
       othall.Text = payrs.Fields("emp_othall")
       lic = payrs.Fields("emp_lic")
       rd = payrs.Fields("emp_rd")
       PF = payrs.Fields("emp_pfp")
       pfno = payrs.Fields("emp_pfno")
       txt_uan.Text = payrs.Fields("emp_uan")
       find_deptname (payrs.Fields("emp_dept"))
       dept_cmb.Text = dname
       
       desi_cmb.ListIndex = find_index_item_data(desi_cmb, payrs!emp_design)
       
       ''emptype_cmb.ListIndex = find_index_item_data(emptype_cmb, payrs!emp_type)
       
       If payrs!emp_cat = "S" Then
           emptype_cmb.Text = "STAFF"
       Else
           emptype_cmb.Text = "WORKER"
       End If
       
       qualify_cmb.ListIndex = find_index_item_data(qualify_cmb, payrs!emp_qualify)
       
       Religion_cmb.ListIndex = find_index_item_data(Religion_cmb, payrs!emp_religion)
       Community_cmb.ListIndex = find_index_item_data(Community_cmb, payrs!emp_community)
       caste_cmb.ListIndex = find_index_item_data(caste_cmb, payrs!emp_caste)
       
       
       txt_Year_Passed.Text = payrs.Fields("emp_passed_year")
       txt_Start_Salary.Text = payrs.Fields("emp_starting_salary")

   
   
       
''      find_desiname (payrs.Fields("emp_design"))
''       desi_cmb.Text = dname
       
''       find_etypename (payrs.Fields("emp_type"))
''       emptype_cmb.Text = dname
      
       dname = ""
''       find_qualifyname (payrs.Fields("emp_qualify"))
       ''qualify_cmb.Text = dname
       
''       work_cmb.Text = payrs.Fields("emp_workplace")
''
''       Set payrs2 = New ADODB.Recordset
''       sql = ("select * from preli_mas where preli_code = " & Val(Religion_cmb.Text))
''       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
''       If Not payrs.EOF Then
''          Religion_cmb.Text = payrs2.Fields("preli_name")
''       End If
''       Set payrs2 = New ADODB.Recordset
''       sql = ("select * from pcomm_mas where pcomm_code = " & Val(Community_cmb.Text))
''       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
''       If Not payrs.EOF Then
''          Community_cmb.Text = payrs2.Fields("pcomm_name")
''       End If
''       Set payrs2 = New ADODB.Recordset
''       sql = ("select * from pcast_mas where pcast_code = " & Val(caste_cmb.Text))
''       payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
''       If Not payrs.EOF Then
''          caste_cmb.Text = payrs2.Fields("pcast_name")
''       End If
       
       
       
       If payrs.Fields("emp_workplace") = "MIL" Then
          work_cmb.Text = "MILL"
       Else
          work_cmb.Text = "OTHER AREA"
       End If
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
       

       
       
       If payrs.Fields("emp_pfeligible") = "Y" Then

          If IsNull(payrs.Fields("emp_pfjoin_date")) = False Then
             dt_pf_join.Value = payrs.Fields("emp_pfjoin_date")
          End If
       End If
       
'''       weekly_off_lst.AddItem payrs.Fields("emp_holiday")
'''       weekly_off_lst.AddItem ("SUNDAY")
'''       weekly_off_lst.AddItem ("MONDAY")
'''       weekly_off_lst.AddItem ("TUESDAY")
'''       weekly_off_lst.AddItem ("WEDNESDAY")
'''       weekly_off_lst.AddItem ("THURSDAY")
'''       weekly_off_lst.AddItem ("FRIDAY")
'''       weekly_off_lst.AddItem ("SATURDAY")
'''
       
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
       If payrs.Fields("emp_esieligible") = "Y" Then
          cmb_esi_eligible.Text = "YES"
       Else
          cmb_esi_eligible.Text = "NO"
       End If
       
       If payrs.Fields("emp_dh_wages_yn") = "Y" Then
          cmb_decholiday_eligbile_yn.Text = "YES"
       Else
          cmb_decholiday_eligbile_yn.Text = "NO"
       End If
       
       
       
       
       ''cmb_bank.ListIndex = payrs.Fields("emp_bank")
       cmb_bank.ListIndex = find_index_item_data(cmb_bank, payrs.Fields("emp_bank"))
       txt_bank_acno.Text = payrs.Fields("emp_bank_acno")
       txt_bank_ifsc.Text = payrs.Fields("emp_bank_ifsc")
       txt_email.Text = payrs.Fields("emp_email")
       txt_esino.Text = payrs.Fields("emp_esino")
       txt_reason.Text = payrs.Fields("emp_resign_reason")
   
   
    If Val(txt_pfsalary.Text) > 0 Then
   
      pfamt.Text = Round((Val(txt_pfsalary.Text) * Val(PF) / 100), 0)
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
''    If PF_NONELIGIBLE.Value = True Then
''       cmd_getpf.Visible = True
''    Else
''       cmd_getpf.Visible = False
''    End If
    cmd_getpf.Visible = False
    txt_interviewername.Text = payrs.Fields("emp_interview_by")
    txt_preinterviewby.Text = payrs.Fields("emp_final_interview_by")
    txt_appointedby.Text = payrs.Fields("emp_appointed_by")
    txt_oe.Text = payrs.Fields("emp_preexp_outside")
    txt_ie.Text = payrs.Fields("emp_preexp_inside")
    
    If payrs.Fields("emp_pi_eligible_yn") = "Y" Then
       cmb_pi_eligbile_yn.Text = "YES"
    Else
       cmb_pi_eligbile_yn.Text = "NO"
    End If
    
    
    If payrs.Fields("emp_pfeligible") = "Y" Then
       cmb_pf_eligible.Text = "YES"
    Else
       cmb_pf_eligible.Text = "NO"
    End If
    
    If payrs.Fields("emp_esieligible") = "Y" Then
       cmb_esi_eligible.Text = "YES"
    Else
       cmb_esi_eligible.Text = "NO"
    End If
    
    
    
    cmb_work_hrs.Text = payrs.Fields("emp_work_hrs")
    
    txt_houseRent.Text = payrs.Fields("emp_houserent_allow")
    txt_Petrol.Text = payrs.Fields("emp_petrol_allow")
    txt_other_benefit.Text = payrs.Fields("emp_other_allow")

   
 
    
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

    i = 1
    sql = "select * from emp_increment where empi_idcode = " & Trim(emp_idcode.Text) & " and empi_millcode = '" & company_code & "' order by empi_year, empi_month"
    payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs2.EOF
      With flx_increment
      .Rows = .Rows + 1
      flx_increment.TextMatrix(i, 0) = payrs2!empi_month
      flx_increment.TextMatrix(i, 1) = payrs2!empi_monthcode
      flx_increment.TextMatrix(i, 2) = payrs2!empi_year
      flx_increment.TextMatrix(i, 3) = payrs2!empi_increment
      

      i = i + 1
      End With
      payrs2.MoveNext
    Wend

    payrs2.Close


End Sub

Private Sub txt_grosspay_Change()
Dim ESIAllowed As Double
''    If searchopt = 0 Then
        Basic.Text = ""
        hra.Text = ""
        fda.Text = ""
        medall.Text = ""
        lta.Text = ""
 ''   End If
 ''  txt_pfsalary.Text = "0"
    
    If cmb_pf_eligible.Text = "YES" Then
          If Val(txt_grosspay.Text) > 0 Then
           Basic.Text = Round(Val(txt_grosspay.Text) * 50 / 100, 0)
           fda.Text = Round(Val(txt_grosspay.Text) * 20 / 100, 0)
           If emptype_cmb.Text = "STAFF" Then
                hra.Text = Round(Val(txt_grosspay.Text) * 30 / 100, 0)
                lta.Text = 0
                medall.Text = 0
           Else
                hra.Text = 0
                lta.Text = Round(Val(txt_grosspay.Text) * 30 / 100, 0)
                medall.Text = 0
           
           End If
        End If
        
       If Val(Basic.Text) + Val(fda.Text) >= eligible_pfsalary Then
          txt_pfsalary.Text = eligible_pfsalary
       Else
          txt_pfsalary.Text = Val(Basic.Text) + Val(fda.Text)
          ''txt_pfsalary.Text = "0"
       End If
    Else
       Basic.Text = txt_grosspay.Text
    End If
    
    ESIAllowed = Val(Basic.Text) + Val(fda.Text)
    
    If cmb_esi_eligible.Text = "YES" Then
    
    
       
       
    
       If ESIAllowed >= eligible_esisalary Then
          txt_esisalary.Text = 0
       Else
          txt_esisalary.Text = ESIAllowed
       
       
       End If
    Else
      txt_esisalary.Text = "0"
    End If
     
    find_Grosspay
    find_netpay
     
     
End Sub

Private Sub txt_itded_Change()
    find_Grosspay
    find_netpay
End Sub

Private Sub txt_itded_KeyPress(KeyAscii As Integer)

On Error GoTo err_handler
    chk_keyascii txt_itded, "N", 5, 2, KeyAscii
    find_Grosspay
    find_netpay
    Exit Sub
err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_mess_subsidy_KeyPress(KeyAscii As Integer)
On Error GoTo err_handler
    chk_keyascii txt_mess_subsidy, "N", 8, 2, KeyAscii
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
''
''Private Sub vda_Change()
''    find_Grosspay
''    find_netpay
''End Sub
''
''Private Sub vda_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii vda, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub

Private Sub washall_Change()
  find_Grosspay
  find_netpay
End Sub

''Private Sub washall_KeyPress(KeyAscii As Integer)
''find_Grosspay
''find_netpay
''On Error GoTo err_handler
''    chk_keyascii washall, "N", 5, 2, KeyAscii
''    Exit Sub
''err_handler:
''    chk = gen_Validation(Err.Number, Err.Description)
''    If chk = 1 Then Resume
''End Sub
Private Sub find_Grosspay()
   Gross.Text = Val(Basic) + Val(ser_wt) + Val(spl_pay) + Val(fda) + Val(vda) + Val(hra) + Val(attall) + Val(ca) + Val(splall) + Val(teaall) + _
                Val(medall) + Val(washall) + Val(lta) + Val(mazall) + Val(fuelall) + Val(profall) + Val(phoneall) + Val(cityall) + Val(othall) + _
                Val(mealsall) + Val(eduall) + Val(healthall)
                
   Gross.Text = Format$(Val(Gross), "0.00")
End Sub

Private Sub find_netpay()
   Dim pfcalc, bonus, gratuity As Double
   
   pfamt.Text = ""
   txt_esiamt.Text = ""
   
   
  If Val(txt_pfsalary.Text) > 0 Then
      pfamt.Text = Round((Val(txt_pfsalary.Text) * Val(PF) / 100), 0)
  End If
   
  If Val(txt_esisalary.Text) > 0 Then
      txt_esiamt.Text = Round((Val(txt_esisalary.Text) * esi_percentage / 100), 0)
  End If
   
   
   
   NET_PAY.Text = Val(Gross.Text) - Val(pfamt) - Val(pfdeduction) - Val(rd) - Val(lic) - Val(txt_itded.Text) - Val(txt_esiamt.Text)
   ctc.Text = Round(Val(Gross.Text) + Val(pfamt) + bonus + gratuity, 0)
End Sub


Public Sub refresh_data()
   cmd_Save.Visible = False
   cmd_Save2.Visible = False
   fillgrid
   txt_location.Text = ""
   
   txt_empcode.Text = ""
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
''   txt_teadeduction.Text = ""
  '' txt_wfund.Text = ""
   txt_bank_acno.Text = ""
   emp_name = ""
   emp_idcode = ""
   txt_email.Text = ""
   cmd_getpf.Visible = False
   empedit_cmb.Text = ""
''   FPCODE.Text = ""
   txt_esino.Text = ""
''   eduall.Text = ""
   txt_interviewername.Text = ""
   txt_preinterviewby.Text = ""
   txt_appointedby.Text = ""
   txt_oe.Text = ""
   txt_ie.Text = ""
   txt_mess_subsidy.Text = ""
   txt_pfsalary.Text = ""
   txt_esisalary.Text = ""
   txt_grosspay.Text = ""
   txt_itded.Text = ""
   txt_esiamt.Text = ""
   pfamt.Text = ""
   txt_houseRent.Text = ""
   txt_Petrol.Text = ""
   txt_other_benefit.Text = ""

End Sub


