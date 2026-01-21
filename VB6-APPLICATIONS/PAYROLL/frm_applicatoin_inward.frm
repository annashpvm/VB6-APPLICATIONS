VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_applicatoin_inward 
   Caption         =   "APPLICATION INWARD"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "frm_applicatoin_inward.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "frm_applicatoin_inward.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "frm_applicatoin_inward.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_applicatoin_inward.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   240
         MaskColor       =   &H000000FF&
         Picture         =   "frm_applicatoin_inward.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   18975
      Begin MSComCtl2.DTPicker dt_dob 
         Height          =   375
         Left            =   5760
         TabIndex        =   62
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   109379585
         CurrentDate     =   44821
      End
      Begin MSComCtl2.DTPicker dt_entry 
         Height          =   375
         Left            =   4920
         TabIndex        =   60
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   109379585
         CurrentDate     =   44821
      End
      Begin VB.ComboBox cmb_entryno 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmb_dept 
         Height          =   315
         Left            =   2280
         TabIndex        =   58
         Top             =   6000
         Width           =   5055
      End
      Begin VB.ComboBox cmb_gender 
         Height          =   315
         ItemData        =   "frm_applicatoin_inward.frx":1DEA
         Left            =   2280
         List            =   "frm_applicatoin_inward.frx":1DF4
         TabIndex        =   56
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmb_sw 
         Height          =   315
         ItemData        =   "frm_applicatoin_inward.frx":1E06
         Left            =   13560
         List            =   "frm_applicatoin_inward.frx":1E10
         TabIndex        =   54
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmb_religion 
         Height          =   315
         Left            =   2280
         TabIndex        =   52
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox txt_entryno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmb_caste 
         Height          =   315
         Left            =   2280
         TabIndex        =   49
         Top             =   7560
         Width           =   2895
      End
      Begin VB.ComboBox cmb_community 
         Height          =   315
         Left            =   2280
         TabIndex        =   47
         Top             =   6960
         Width           =   2775
      End
      Begin VB.TextBox txt_ContactNo 
         Height          =   400
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   46
         Top             =   3960
         Width           =   4095
      End
      Begin VB.TextBox txt_Reference 
         Height          =   400
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   44
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txt_reason 
         Height          =   400
         Left            =   13560
         MaxLength       =   100
         TabIndex        =   42
         Top             =   6720
         Width           =   4095
      End
      Begin VB.TextBox txt_interview_by 
         Height          =   400
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   40
         Top             =   5160
         Width           =   4095
      End
      Begin VB.ComboBox cmb_interview_status 
         Height          =   315
         ItemData        =   "frm_applicatoin_inward.frx":1E23
         Left            =   13560
         List            =   "frm_applicatoin_inward.frx":1E30
         TabIndex        =   38
         Top             =   6000
         Width           =   1815
      End
      Begin VB.ComboBox cmb_interview 
         Height          =   315
         ItemData        =   "frm_applicatoin_inward.frx":1E52
         Left            =   13560
         List            =   "frm_applicatoin_inward.frx":1E5C
         TabIndex        =   36
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txt_expected_salary 
         Height          =   400
         Left            =   13560
         TabIndex        =   34
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txt_posting 
         Height          =   400
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   32
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txt_workedas 
         Height          =   400
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txt_workedcompany 
         Height          =   400
         Left            =   13560
         MaxLength       =   100
         TabIndex        =   28
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txt_experience 
         Height          =   400
         Left            =   2280
         TabIndex        =   26
         Top             =   8040
         Width           =   1095
      End
      Begin VB.TextBox txt_qualification 
         Height          =   400
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   23
         Top             =   5400
         Width           =   4095
      End
      Begin VB.TextBox txt_aadhaar 
         Height          =   400
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   21
         Top             =   4920
         Width           =   4095
      End
      Begin VB.TextBox txt_email 
         Height          =   400
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   19
         Top             =   4440
         Width           =   4095
      End
      Begin VB.TextBox txt_Pincode 
         Height          =   400
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   16
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txt_Addr2 
         Height          =   400
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2400
         Width           =   6975
      End
      Begin VB.TextBox txt_Place 
         Height          =   400
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox txt_Addr1 
         Height          =   400
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1920
         Width           =   6975
      End
      Begin VB.TextBox txt_Name 
         Height          =   375
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   8
         Top             =   840
         Width           =   6975
      End
      Begin VB.Label Label27 
         Caption         =   "Date of Birth"
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
         Left            =   4320
         TabIndex        =   61
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Date"
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
         Left            =   4440
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Gender"
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
         Left            =   360
         TabIndex        =   57
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label24 
         Caption         =   "Staff / Worker"
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
         Left            =   10920
         TabIndex        =   55
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label23 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label22 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   7560
         Width           =   2295
      End
      Begin VB.Label Label21 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   6960
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Reference"
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
         Left            =   10920
         TabIndex        =   45
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Reason for Rejection"
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
         Left            =   10920
         TabIndex        =   43
         Top             =   6840
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Inverviewed By"
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
         Left            =   10920
         TabIndex        =   41
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Interview Status"
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
         Left            =   10920
         TabIndex        =   39
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Interview Contacted (Y/N)"
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
         Left            =   10920
         TabIndex        =   37
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Expected Salary"
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
         Left            =   10920
         TabIndex        =   35
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Apply for the Posting"
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
         Left            =   10920
         TabIndex        =   33
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Worked as"
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
         Left            =   10920
         TabIndex        =   31
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Worked Company"
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
         Left            =   10920
         TabIndex        =   29
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Experiance"
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
         Left            =   240
         TabIndex        =   27
         Top             =   8040
         Width           =   2295
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Edu.Qualification"
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
         Left            =   240
         TabIndex        =   24
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Aadhaar No"
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
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "E-mail ID"
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
         Left            =   240
         TabIndex        =   20
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Contact No"
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
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Pin Code"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Place"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Entry No"
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
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Name "
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
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frm_applicatoin_inward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_entryno_Click()
    savechk = 1
    Dim payrs As New ADODB.Recordset
    pst_qry = "select  * from mas_applications  where a_entryno = " & Val(cmb_entryno.Text)
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    If payrs.EOF Then
       MsgBox ("Data not avaiable")
       Exit Sub
    Else
    
       txt_entryno.Text = Val(cmb_entryno.Text)
       txt_Name.Text = payrs.Fields("a_name")
       
       If payrs.Fields("a_name") = "M" Then
         cmb_gender.Text = "MALE"
       Else
         cmb_gender.Text = "FEMALE"
       End If
       

   
       dt_entry = payrs.Fields("a_date")
       dt_dob = payrs.Fields("a_dob")
       txt_Addr1.Text = payrs.Fields("a_address1")
       txt_Addr2.Text = payrs.Fields("a_address2")
       txt_Place.Text = payrs.Fields("a_place")
       txt_Pincode.Text = payrs.Fields("a_pin")
       
       txt_ContactNo.Text = payrs.Fields("a_contactno")
       txt_email.Text = payrs.Fields("a_emailed")
       txt_aadhaar.Text = payrs.Fields("a_aadhaar")
       txt_qualification.Text = payrs.Fields("a_qualification")
       
       txt_experience.Text = payrs.Fields("a_experiance")
       txt_workedcompany.Text = payrs.Fields("a_workedcompany")
       txt_workedas.Text = payrs.Fields("a_workedas")
       txt_Reference.Text = payrs.Fields("a_reference")
       txt_posting.Text = payrs.Fields("a_posting")
       txt_expected_salary.Text = payrs.Fields("a_expected_salary")
       txt_interview_by.Text = payrs.Fields("a_interview_by")
       txt_reason.Text = payrs.Fields("a_reason")
       If payrs.Fields("a_inverview_status") = "S" Then
         cmb_interview_status.Text = "SELECTED"
       ElseIf payrs.Fields("a_inverview_status") = "P" Then
         cmb_interview_status.Text = "PENDING"
       Else
         cmb_interview_status.Text = "REJECTED"
       End If
       
        If payrs.Fields("a_interview_yn") = "Y" Then
         cmb_interview.Text = "YES"
       Else
         cmb_interview.Text = "NO"
       End If
       
        If payrs.Fields("a_sw") = "S" Then
         cmb_sw.Text = "STAFF"
       Else
         cmb_sw.Text = "WORKER"
       End If
       
       cmb_religion.ListIndex = find_index_item_data(cmb_religion, payrs.Fields("a_religion"))
       cmb_community.ListIndex = find_index_item_data(cmb_community, payrs.Fields("a_community"))
       cmb_caste.ListIndex = find_index_item_data(cmb_caste, payrs.Fields("a_caste"))
       cmb_dept.ListIndex = find_index_item_data(cmb_dept, payrs.Fields("a_dept"))
       
    End If
    payrs.Close

End Sub

Private Sub edit_Click()
    cmb_entryno.Clear
    cmb_entryno.Visible = True
    Dim payrs As New ADODB.Recordset
    pst_qry = "select a_entryno from mas_applications order by a_entryno desc"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_entryno.AddItem payrs("a_entryno")
        payrs.MoveNext
    Wend
    payrs.Close
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmb_entryno.Visible = False
    
    savechk = 0
    Dim payrs As New ADODB.Recordset
    pst_qry = "select max(a_entryno)+1 as entno from mas_applications"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    no = 1
    txt_entryno.Text = no
    If Not IsNull(payrs!entno) Then
        If Not payrs.EOF Then
            txt_entryno.Text = payrs!entno
        End If
    End If
    payrs.Close
    

    

    sql = "Select * from  pdept_mas order by dept_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_dept.AddItem payrs(1)
        cmb_dept.ItemData(cmb_dept.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
    
    
    sql = "Select * from  pcomm_mas order by pcomm_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_community.AddItem payrs(1)
        cmb_community.ItemData(cmb_community.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
    
    
    sql = "Select * from  pcast_mas order by pcast_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_caste.AddItem payrs(1)
        cmb_caste.ItemData(cmb_caste.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
    
    
    sql = "Select * from  preli_mas order by preli_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_religion.AddItem payrs(1)
        cmb_religion.ItemData(cmb_religion.NewIndex) = payrs(0)
        payrs.MoveNext
    Wend
    payrs.Close
        
    
End Sub

Private Sub refresh_Click()
    cmb_entryno.Visible = False
    savechk = 0
    Dim payrs As New ADODB.Recordset
    pst_qry = "select max(a_entryno)+1 as entno from mas_applications"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    no = 1
    txt_entryno.Text = no
    If Not IsNull(payrs!entno) Then
        If Not payrs.EOF Then
            txt_entryno.Text = payrs!entno
        End If
    End If
    payrs.Close
    
       txt_Name.Text = ""
       txt_Addr1.Text = ""
       txt_Addr2.Text = ""
       txt_Place.Text = ""
       txt_Pincode.Text = ""
       
       txt_ContactNo.Text = ""
       txt_email.Text = ""
       txt_aadhaar.Text = ""
       txt_qualification.Text = ""
       
       txt_experience.Text = ""
       txt_workedcompany.Text = ""
       txt_workedas.Text = ""
       txt_Reference.Text = ""
       txt_posting.Text = ""
       txt_expected_salary.Text = ""
       txt_interview_by.Text = ""
       txt_reason.Text = ""
    
    
End Sub

Private Sub SAVE_Click()
   Dim ecat As String
   If Trim(txt_entryno.Text) = "" Then
      MsgBox ("Entry Number is blank ")
     '' emp_idcode.SetFocus
      Exit Sub
   End If
   If Trim(txt_Name.Text) = "" Then
      MsgBox ("Employee Name is blank - correct it ")
      txt_Name.SetFocus
      Exit Sub
   End If
   If Trim(txt_aadhaar.Text) = "" Then
      MsgBox ("Employee Aadhaar Number is blank ")
      txt_aadhaar.SetFocus
      Exit Sub
   End If
   If Trim(txt_Addr1.Text) = "" Then
      MsgBox ("Address Line is  blank ")
      txt_Addr1.SetFocus
      Exit Sub
   End If

''
''
   If Trim(txt_Place.Text) = "" Then
      MsgBox ("Place is Empty ")
      txt_Place.SetFocus
      Exit Sub
   End If


   If Trim(cmb_religion.Text) = "" Then
      MsgBox ("Religion is blank - correct it ")
      cmb_religion.SetFocus
      Exit Sub
   End If

   If Trim(cmb_community.Text) = "" Then
      MsgBox ("Community is blank - correct it ")
      cmb_community.SetFocus
      Exit Sub
   End If
   If Trim(cmb_caste.Text) = "" Then
      MsgBox ("Employee caste is blank - correct it ")
      cmb_caste.SetFocus
      Exit Sub
   End If

   

   If Trim(cmb_dept.Text) = "" Then
      MsgBox ("Department name is blank - Select department ")
      cmb_dept.SetFocus
      Exit Sub
   End If
   
   
   If Trim(cmb_sw.Text) = "" Then
      MsgBox ("Select Staff / Worker ")
      cmb_sw.SetFocus
      Exit Sub
   End If
   
   
   
   If Trim(cmb_interview.Text) = "" Then
      MsgBox ("Select Interview By")
      cmb_interview.SetFocus
      Exit Sub
   End If
   
   
      
   If Trim(cmb_interview_status.Text) = "" Then
      MsgBox ("Select Interview Status ")
      cmb_interview_status.SetFocus
      Exit Sub
   End If
   
   
    Dim payrs As New ADODB.Recordset
    
    If savechk = 0 Then
    

      sql = "Select * from mas_applications"
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
      payrs.AddNew
      payrs.Fields("a_entryno") = Val(txt_entryno.Text)
   Else
      sql = "select * from mas_applications where a_entryno = " & Val(txt_entryno.Text)
      payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
   End If
   




   

   payrs.Fields("a_name") = UCase(txt_Name.Text)
   payrs.Fields("a_sex") = IIf(cmb_gender.Text = "MALE", "M", "F")
   payrs.Fields("a_address1") = UCase(txt_Addr1.Text)
   payrs.Fields("a_address2") = UCase(txt_Addr2.Text)
   payrs.Fields("a_place") = UCase(txt_Place.Text)
   payrs.Fields("a_pin") = txt_Pincode.Text
   payrs.Fields("a_contactno") = txt_ContactNo.Text
   payrs.Fields("a_emailed") = txt_email.Text
   payrs.Fields("a_aadhaar") = txt_aadhaar.Text
   
   payrs.Fields("a_qualification") = txt_qualification.Text
   payrs.Fields("a_religion") = cmb_religion.ItemData(cmb_religion.ListIndex)
   payrs.Fields("a_community") = cmb_community.ItemData(cmb_community.ListIndex)
   payrs.Fields("a_caste") = cmb_caste.ItemData(cmb_caste.ListIndex)
   payrs.Fields("a_dept") = cmb_dept.ItemData(cmb_dept.ListIndex)
   payrs.Fields("a_experiance") = Val(txt_experience.Text)
   payrs.Fields("a_workedcompany") = UCase(txt_workedcompany.Text)
   payrs.Fields("a_workedas") = UCase(txt_workedas.Text)
   payrs.Fields("a_sw") = Left(cmb_sw.Text, 1)
   payrs.Fields("a_reference") = UCase(txt_Reference.Text)
   payrs.Fields("a_posting") = UCase(txt_posting.Text)
   payrs.Fields("a_expected_salary") = Val(txt_expected_salary.Text)
   payrs.Fields("a_interview_yn") = Left(cmb_interview.Text, 1)
   payrs.Fields("a_interview_by") = UCase(txt_interview_by.Text)
   payrs.Fields("a_inverview_status") = Left(cmb_interview_status.Text, 1)
   payrs.Fields("a_reason") = txt_reason.Text
   payrs.Fields("a_date") = dt_entry.Value
   payrs.Fields("a_dob") = dt_dob.Value
   
  
   payrs.Update
   payrs.Close
   MsgBox ("Data updated")
   refresh_Click

End Sub
