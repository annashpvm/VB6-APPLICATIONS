VERSION 5.00
Begin VB.Form Weeklyoffmaster 
   Caption         =   "WEEKLY OFF"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton SAVE 
      Caption         =   "SAVE"
      Height          =   855
      Left            =   3720
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Frame LEAVE 
      Height          =   4695
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   9735
      Begin VB.ComboBox cmb_leave 
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
         Left            =   2520
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
      End
      Begin VB.ComboBox employee_cmb 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label2 
         Caption         =   "LEAVE DAY"
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
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "SELECT EMPLOYEE"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Weeklyoffmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXIT_Click()
   Unload Me
End Sub

Private Sub Form_Load()
        cmb_leave.AddItem "SUNDAY"
        cmb_leave.AddItem "MONDAY"
        cmb_leave.AddItem "TUESDAY"
        cmb_leave.AddItem "WEDNESDAY"
        cmb_leave.AddItem "THURSDAY"
        cmb_leave.AddItem "FRIDAY"
        cmb_leave.AddItem "SATURDAY"
        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\personnel\plat04s\Acu301.mdb"
        Dim mdbrs As New ADODB.Recordset
        mdb_qry = "Select * from Member where companyID <> '9999'"
        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
        totemp = 0
        While Not mdbrs.EOF
             employee_cmb.AddItem (mdbrs(1))
             employee_cmb.ItemData(employee_cmb.NewIndex) = mdbrs(0)
             mdbrs.MoveNext
        Wend
End Sub

