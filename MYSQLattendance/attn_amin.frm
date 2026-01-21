VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form attn_main 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ATTENDANCE REPORTS"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=servall;Initial Catalog=servall"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=servall;Initial Catalog=servall"
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
   Begin VB.CommandButton CMD_EXIT 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&EXIT"
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
      Left            =   3240
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_continue 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&CONTINUE"
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
      Left            =   1680
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   4335
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   9255
      Begin VB.TextBox txt_pw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "#"
         TabIndex        =   6
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox cmb_fin 
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
         Left            =   2520
         TabIndex        =   4
         Text            =   "cmb_fin"
         Top             =   1800
         Width           =   4095
      End
      Begin VB.ComboBox cmb_mill 
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
         Left            =   2520
         TabIndex        =   2
         Text            =   "cmb_mill"
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fin-Year"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Mill Name"
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
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
   End
End
Attribute VB_Name = "attn_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_continue_Click()
    millcode = cmb_mill.ItemData(cmb_mill.ListIndex)
    fincode = cmb_fin.ItemData(cmb_fin.ListIndex)
    millname = Trim(cmb_mill.Text)
    Dim passchk As New ADODB.Recordset
    sql_qry = "select * from mas_company where company_code = " & cmb_mill.ItemData(cmb_mill.ListIndex)
    passchk.Open sql_qry, attndb, 1, 2
    With passchk
         Dim pw As String
         pw = Trim(txt_pw.Text)
         If passchk!company_pass <> pw Then
            MsgBox ("Passwozrd is wrong!..")
            txt_pw.SetFocus
            Exit Sub
         End If
         attn_mainmenu.Show
         attn_mainmenu.ZOrder
    End With
End Sub

Private Sub cmd_exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  attndb.Open "Provider=SQLOLEDB.1;Password=serdat;Persist Security Info=True;User ID=sa;DATABASE=servall;Data Source=servalldata"
  ''attndb.Open "Provider=SQLOLEDB.1;Password=hodat;Persist Security Info=True;User ID=sa;DATABASE=servall;Data Source=hodata"
  Fill_Combo_mill "Select company_name,company_code from mas_company order by company_code", cmb_mill
  Fill_Combo "Select fin_year,fin_code from mas_finyear order by fin_year", cmb_fin
  txt_pw.Text = "DPM"
End Sub

