VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_weighbridge_entry 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WEIGH BRIDGE ENTRY"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15450
   Icon            =   "frm_weighbridge_entry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   15450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Party Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton opt_party_manual 
         Caption         =   "Manual"
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
         Left            =   1920
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton opt_party_auto 
         Caption         =   "Auto"
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
         Left            =   600
         TabIndex        =   47
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtReason 
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
      Left            =   10560
      MaxLength       =   49
      TabIndex        =   44
      Top             =   9000
      Width           =   3495
   End
   Begin VB.TextBox txt_password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   10440
      PasswordChar    =   "#"
      TabIndex        =   43
      Top             =   8520
      Width           =   2055
   End
   Begin VB.CommandButton cmd_close 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8400
      Width           =   795
   End
   Begin Crystal.CrystalReport Cry_rep1 
      Left            =   120
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   13560
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   12480
      TabIndex        =   30
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtnew3 
      Height          =   495
      Left            =   12480
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtnew 
      Height          =   615
      Left            =   10800
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd_on 
      Caption         =   "ON"
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2400
      TabIndex        =   25
      Top             =   8040
      Width           =   4695
      Begin VB.CommandButton cmd_Weight 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   8520
   End
   Begin VB.TextBox txt_getwt2 
      Height          =   375
      Left            =   11520
      TabIndex        =   24
      Top             =   5280
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
      Left            =   11520
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_wt_from_serailport 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txt_getwt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1455
      Left            =   10920
      TabIndex        =   19
      Top             =   720
      Width           =   3495
      Begin VB.Label lbl_getwt 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1095
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   9615
      Begin VB.ComboBox cmb_vehicle 
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
         TabIndex        =   36
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dt_ticket2 
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy hh:mm:ss  tt"
         Format          =   120848387
         CurrentDate     =   45257
      End
      Begin MSComCtl2.DTPicker dt_ticket 
         Height          =   495
         Left            =   2400
         TabIndex        =   33
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
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
         Format          =   120848385
         CurrentDate     =   45257
      End
      Begin VB.TextBox txt_NetWt 
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txt_EmptyWt 
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
         Left            =   2400
         TabIndex        =   15
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox txt_LoadWt 
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
         Left            =   2400
         TabIndex        =   13
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ADD"
         Height          =   195
         Left            =   7800
         TabIndex        =   12
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmb_supplier 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3840
         Width           =   5295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ADD"
         Height          =   195
         Left            =   7800
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmb_material 
         Appearance      =   0  'Flat
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
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   5295
      End
      Begin VB.ComboBox cmb_load 
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
         ItemData        =   "frm_weighbridge_entry.frx":000C
         Left            =   2400
         List            =   "frm_weighbridge_entry.frx":0016
         TabIndex        =   6
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txt_vehicle 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txt_ticketNo 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lbltime 
         BackColor       =   &H8000000B&
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
         Height          =   255
         Left            =   5040
         TabIndex        =   41
         Top             =   2520
         Width           =   4095
      End
      Begin VB.OLE OLE1 
         Height          =   30
         Left            =   3480
         TabIndex        =   34
         Top             =   5520
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ticket Date"
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
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Net Weight"
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
         Left            =   360
         TabIndex        =   18
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lbl_emptywt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empty Weight"
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
         Left            =   360
         TabIndex        =   16
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label lbl_loadwt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loaded Weight"
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
         Left            =   360
         TabIndex        =   14
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lbl_supplier 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Party Name"
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
         Left            =   360
         TabIndex        =   10
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Material"
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
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load Status"
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
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl_vechicle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vehicle Number"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbl_ticket 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ticket Number"
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
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   1200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REASON"
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
      Left            =   9120
      TabIndex        =   45
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FIRST WEIGHT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   3120
      TabIndex        =   37
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frm_weighbridge_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mwt As Integer
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
Dim pst_qry, sizecode, destag As String
Dim mm, yy, runyear, rno, yr, dd  As String
Dim varietycode, custcode, codesize, pin_cnt As Integer

Dim dmy As String

Dim partycode As Integer

Dim wtchk As Integer
Dim addwt As Integer
    Dim wbtype As String
    
    Dim firstno, lastno As Double
Dim winderno As String

Dim portnumber, baudRate As Integer

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
    Dim fincode As Integer
    Dim saveflag As String
    
Private Sub cmb_load_Click()
    If cmb_load.Text = "Loaded" Then
''       cmb_material.Enabled = True
       cmb_supplier.Enabled = True
    Else
''       cmb_material.Enabled = False
''       cmb_supplier.Enabled = False
    End If
    Dim truck As String
    
     If load_first_sec_type = "F" Then
     
        truck = UCase(Replace(txt_vehicle.Text, " ", ""))
         
        pst_qry = "select *  from trn_weighbridge_entry where t_wb_vehicle =  '" & truck & "' and t_wb_compcode = " & compcode & " and  t_wb_net_weight = 0 and t_wb_upd = 'N' and t_wb_year = " & gin_finid & ""
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
           MsgBox ("Already this Truck was Accounted.. in the Ticket Number : " + CStr(adors("t_wb_ticketno")))
        End If
     
     End If
End Sub

Private Sub cmb_load_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmb_material.SetFocus
  End If
End Sub

Private Sub cmb_supplier_Click()
      partycode = cmb_supplier.ItemData(cmb_supplier.ListIndex)
End Sub

Private Sub cmb_vehicle_Click()
        pst_qry = "select *  from trn_weighbridge_entry where t_wb_vehicle =  '" & cmb_vehicle.Text & "' and t_wb_compcode = " & compcode & " and  t_wb_net_weight = 0  and t_wb_upd = 'N' and t_wb_year = " & gin_finid & ""
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
            txt_ticketNo.Text = adors("t_wb_ticketno")
            If Trim(adors("t_wb_item")) <> "" Then
               cmb_material.Text = RTrim(LTrim(adors("t_wb_item")))
            End If
            If adors("t_wb_party") <> "" Then
            cmb_supplier.Text = adors("t_wb_party")
            End If
            txt_vehicle.Text = cmb_vehicle.Text
            lbltime.Caption = adors("t_wb_1st_time")
            If adors("t_wb_1st_loadtype") = "L" Then
               txt_LoadWt.Text = adors("t_wb_1st_weight")
               txt_EmptyWt.Text = ""
               cmb_load.Text = "Empty"
            Else
               txt_EmptyWt.Text = adors("t_wb_1st_weight")
               txt_LoadWt.Text = ""
               cmb_load.Text = "Loaded"
            End If
        End If
End Sub

Private Sub cmb_vehicle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmb_load.SetFocus
End If
End Sub

Private Sub cmd_close_Click()
       Dim adocmd_mysql As New ADODB.Command
       Dim adors As New ADODB.Recordset
       If Trim(txt_ticketNo.Text) = "" Then
          MsgBox ("Select Ticket Number... ")
          Exit Sub
       End If

       If Trim(txtReason.Text) = "" Then
          MsgBox ("Enter Reason For Cancel... ")
          Exit Sub
       End If


       If Trim(txt_password.Text) = "Close1" Then
        adocmd_mysql.ActiveConnection = gen_connection_mysql
         Dim pst_ans As String
         pst_ans = MsgBox("Kindly confirm Again. Do you want to CLOSE this Ticket...  ", vbYesNo)
         If pst_ans = vbYes Then
                 
            pst_qry = "update trn_weighbridge_entry set t_wb_cancel_reason = '" & txtReason.Text & "'  , t_wb_upd = 'C' ,t_wb_1st_weight = 0, t_wb_2nd_weight = 0  Where t_wb_compcode = " & compcode & " And t_wb_year = " & gin_finid & " And t_wb_ticketno = " & Val(txt_ticketNo.Text) & " and t_wb_net_weight = 0"
            adocmd_mysql.CommandText = pst_qry
            Set adors = adocmd_mysql.Execute
            MsgBox ("Ticket Number : " + txt_ticketNo.Text + " has been closed ")
             txt_vehicle.Text = ""
             cmb_load.Text = ""
             cmb_material.Text = ""
    ''         cmb_supplier.Text = ""
             txt_LoadWt.Text = ""
             txt_EmptyWt.Text = ""
             txt_NetWt.Text = ""
            
             refresh_data
        End If
        Else
            MsgBox ("Password Error")
            txt_password.SetFocus
        End If

End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Function refresh_data()
    Dim pst_qry As String
    Dim pdb_seqno As Long
    Dim pdb_seqno_mysql As Long
    Dim pin_cnt As Long
    
    dt_ticket.Value = Now
    dt_ticket2.Value = Now
    
    adocmd_mysql.ActiveConnection = gen_connection_mysql



    pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_type = 'A' and t_wb_year = " & gin_finid & "  and t_wb_compcode = " & compcode
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
         txt_ticketNo.Text = adors("tikcetno")
    End If
         
         
         
    
    If load_first_sec_type = "S" Then
        cmb_vehicle.Clear
        txt_ticketNo.Text = ""
        pst_qry = "select *  from trn_weighbridge_entry where t_wb_compcode = " & compcode & " and t_wb_net_weight = 0 and t_wb_upd = 'N'"
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
            For pin_cnt = 1 To adors.RecordCount
                     cmb_vehicle.AddItem (adors("t_wb_vehicle"))
                     adors.MoveNext
            Next
        End If
    End If
    
    

End Function

Private Sub cmd_Refresh_Click()
        
    
 
On Error GoTo err_handler:
    
    If mwt = 0 Then
    
        
    Dim baudRate1 As String
    Dim parity As String
    Dim dataBits As String
    Dim stopBits As String
    Dim commSettings As String
    
    baudRate1 = baudRate
    parity = "N"   ' e.g., "N"
    dataBits = "8"  ' e.g., "8"
    stopBits = "1"  ' e.g., "1"
    
    ' Combine into one settings string
    commSettings = baudRate & ", " & parity & ", " & dataBits & ", " & stopBits
    
        With MSComm1
''              .Settings = "1200, N, 8, 1"
                .Settings = commSettings
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

Private Sub Command6_Click()

End Sub

Private Sub cmd_Save_Click()

    dt_ticket2.Value = Now
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    Dim pst_qry, wt_type, truck As String

    Dim savemsg As String
    Dim first_sec_wt  As Long
    
    Dim nwt As Long

    adocmd_mysql.ActiveConnection = gen_connection_mysql

    
    If cmb_load.Text = "Loaded" And opt_party_auto.Value = True And partycode = 0 Then
       MsgBox (" Error in Paryt Name .. Select Party Name ")
       cmb_supplier.SetFocus
       Exit Sub
    
    End If
    
    dt_ticket2.Value = Now
    
 
    

    If cmb_supplier.ListIndex = -1 Then
       MsgBox (" Select Party Name .... ")
       cmb_supplier.SetFocus
       Exit Sub
    End If

    If cmb_material.ListIndex = -1 Then
       MsgBox (" Select Item Name .... ")
       cmb_material.SetFocus
       Exit Sub
    End If
    
    
    If txt_vehicle.Text = "" Then
       MsgBox (" Enter Truck Number .... ")
       txt_vehicle.SetFocus
       Exit Sub
    End If
    
    
    If cmb_load.Text = "" Then
       MsgBox ("Select Load Type.... ")
       cmb_load.SetFocus
       Exit Sub
    End If
    

    If Val(txt_LoadWt.Text) + Val(txt_EmptyWt.Text) = 0 Then
        MsgBox ("Weight is  Empty . Please check ... ")
        Exit Sub
    End If


    
    If cmb_load.Text = "Loaded" Then
       first_sec_wt = Val(txt_LoadWt.Text)
       wt_type = "L"
       If cmb_material.Text = "" Then
           MsgBox ("Select Material Type.... ")
           cmb_material.SetFocus
           Exit Sub
       End If
        
       If cmb_supplier.Text = "" Then
           MsgBox ("Select Supplier Name .... ")
           cmb_supplier.SetFocus
           Exit Sub
       End If
       
    Else
       first_sec_wt = Val(txt_EmptyWt.Text)
       wt_type = "E"
    End If
    
    If first_sec_wt = 0 Then
       MsgBox ("Error in Weight .. Click WEIGHT Button Again.... ")
       Exit Sub
    End If
    
    truck = UCase(Replace(txt_vehicle.Text, " ", ""))
    
    
     If load_first_sec_type = "F" Then
     
         
        pst_qry = "select *  from trn_weighbridge_entry where t_wb_vehicle =  '" & truck & "' and t_wb_compcode = " & compcode & " and  t_wb_net_weight = 0 and t_wb_upd = 'N' and t_wb_year = " & gin_finid & ""
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
           MsgBox ("Already this Truck was Accounted.. in the Ticket Number : " + CStr(adors("t_wb_ticketno")))
           Exit Sub
        End If
     
     End If
     
    
    If cmb_load.Text = "Loaded" Then
       If Val(txt_LoadWt.Text) <> Val(lbl_getwt.Caption) Then
          MsgBox ("Load Weight is Difference - Please click WEIGHT Button Again... ")
          Exit Sub
       End If

    Else
       If Val(txt_EmptyWt.Text) <> Val(lbl_getwt.Caption) Then
          MsgBox ("Empty Weight is Difference - Please click WEIGHT Button Again... ")
          Exit Sub
       End If

    End If

        
        
    If load_first_sec_type = "F" Then
        wbtype = "A"
        


       pst_qry = "select * from mas_wb_party where  party_type = 1 and party_name = '" & LTrim(RTrim(cmb_supplier.Text)) & "'"
       adocmd_mysql.CommandText = pst_qry
       Set adors = adocmd_mysql.Execute
       If adors.RecordCount > 0 Then
           wbtype = "Z"
       End If
       adors.Close
       
''
''       pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_type = '" & wbtype & "' and t_wb_year = " & gin_finid & "  and t_wb_compcode = " & compcode
''       adocmd_mysql.CommandText = pst_qry
''       Set adors = adocmd_mysql.Execute
''       If adors.RecordCount > 0 Then
''           If wbtype = "A" Then
''               txt_ticketNo.Text = adors("tikcetno")
''           Else
''              If adors("tikcetno") = 1 Then
''                  txt_ticketNo.Text = Trim(dmy) + "01"
''              Else
''                txt_ticketNo.Text = adors("tikcetno")
''              End If
''           End If
''      End If
      
      If wbtype = "A" Then
         pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_type = 'A' and t_wb_year = " & gin_finid & "  and t_wb_compcode = " & compcode
         adocmd_mysql.CommandText = pst_qry
         Set adors = adocmd_mysql.Execute
         If adors.RecordCount > 0 Then
            txt_ticketNo.Text = adors("tikcetno")
         End If
      Else
         pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_type = 'Z' and t_wb_year = " & gin_finid & "  and t_wb_compcode = " & compcode & "  and t_wb_date = '" & Format(dt_ticket, "yyyy-MM-dd") & "'"
         adocmd_mysql.CommandText = pst_qry
         Set adors = adocmd_mysql.Execute
         If adors("tikcetno") = 1 Then
            txt_ticketNo.Text = Trim(dmy) + "01"
         Else
           txt_ticketNo.Text = adors("tikcetno")
         End If
      End If
      
      
      pst_qry = "insert into trn_weighbridge_entry (t_wb_year, t_wb_compcode, t_wb_ticketno, t_wb_date, t_wb_vehicle," _
        & " t_wb_item, t_wb_party, t_wb_area, t_wb_1st_loadtype, t_wb_1st_weight, t_wb_1st_time ,t_wb_type) values ( " & gin_finid & " ," & compcode & " ," & Val(txt_ticketNo.Text) & " ,'" & Format(dt_ticket, "yyyy-MM-dd") & "' , '" & truck & "'," _
        & " '" & Trim(Left(cmb_material.Text, 39)) & " ' , '" & Left(cmb_supplier.Text, 49) & "' , '', '" & wt_type & "' , " & first_sec_wt & " ,'" & Format(dt_ticket2.Value, "yyyy-MM-dd HH:MM:SS") & "','" & wbtype & " ')  "
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
''        MsgBox ("First Entry Saved...")
        savemsg = "First Entry Saved...  "
       
       

    Else
    
        nwt = Val(txt_LoadWt.Text) - Val(txt_EmptyWt.Text)
        
''Modified on 27/05/2024
        If nwt < -100 Then
            MsgBox ("Weight is Error .. Click Weight Button Again")
            Exit Sub
        End If
        
        pst_qry = "update trn_weighbridge_entry set  t_wb_party =  '" & Left(cmb_supplier.Text, 49) & "' , t_wb_item = '" & Left(cmb_material.Text, 39) & "' , t_wb_2nd_loadtype = '" & wt_type & "' , t_wb_2nd_time = '" & Format(dt_ticket2.Value, "yyyy-MM-dd HH:MM:SS") & "', t_wb_2nd_weight = " & first_sec_wt & " , t_wb_net_weight = " & nwt & "  where t_wb_year = " & gin_finid & "  and t_wb_ticketno = " & Val(txt_ticketNo.Text) & ""
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
      '' MsgBox ("Second Entry Saved...")
        savemsg = "Second Entry Saved..."
    End If
    Dim pst_ans, gst_repconnect, qry As String
    

        pst_ans = MsgBox(savemsg + " Do you want to Print the Ticket...  ", vbYesNo)
        If pst_ans = vbYes Then
        
            Cry_rep1.Formulas(0) = "opt = 0"
            qry = "{trn_weighbridge_entry.t_wb_compcode}  =  " & compcode & " and  {trn_weighbridge_entry.t_wb_year}  =  " & gin_finid & " and   {trn_weighbridge_entry.t_wb_ticketno}  =  " & txt_ticketNo.Text
            MousePointer = vbDefault
            ''gst_repconnect = "dsn=pay_new;uid=sa;pwd=serdat"
           gst_repconnect = "dsn=mysql;uid=root;pwd=P@ssw0rD;database=shvpm"
            
''            Cry_rep1.PrinterSelect
            Cry_rep1.ReportFileName = "d:\wbreports\weigh_slip.rpt"
''            Cry_rep1.ReportFileName = "\\10.0.0.243\wbreports\weigh_slip.rpt"
            Cry_rep1.ReplaceSelectionFormula (qry)
            Cry_rep1.WindowState = crptMaximized
            Cry_rep1.Connect = gst_repconnect
            
            Cry_rep1.Action = 1
        End If
        pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_year = " & gin_finid & "  and t_wb_compcode = " & compcode
        adocmd_mysql.CommandText = pst_qry
        Set adors = adocmd_mysql.Execute
        If adors.RecordCount > 0 Then
            txt_ticketNo.Text = adors("tikcetno")
        End If
    txt_vehicle.Text = ""
    cmb_load.Text = ""
''    cmb_material.Text = ""
''    cmb_supplier.Text = ""
    txt_LoadWt.Text = ""
    txt_EmptyWt.Text = ""
    txt_NetWt.Text = ""
   
    refresh_data
             
    
End Sub

Private Sub cmd_Weight_Click()
    
    If cmb_load.Text = "Loaded" Then
       txt_LoadWt.Text = Val(lbl_getwt.Caption)
    Else
       txt_EmptyWt.Text = Val(lbl_getwt.Caption)
    End If
    
    If Val(txt_LoadWt.Text) > 0 And Val(txt_EmptyWt.Text) > 0 Then
       txt_NetWt.Text = Val(txt_LoadWt.Text) - Val(txt_EmptyWt.Text)
    Else
       txt_NetWt.Text = ""
    End If
    
    
        
End Sub



Private Sub Command7_Click()
    wtchk = 1
End Sub

''Private Sub Command6_Click()
'' Dim xxx As String
''
'' '' AlphaNum (Text1.Text)
''xxx = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, ":") + 1, 8)
''MsgBox (xxx)
''xxx = Mid(xxx, Len(xxx) - 4)
''MsgBox (xxx)
''
''
''End Sub

''Function AlphaNum(strBefore As String)
''  Dim CleanString As String
''  CleanString = strBefore
''  Dim strAfter As String
''  Dim intAscii As Integer
''  Dim strTest As String
''  Dim dblX As Double
''  Dim dblLen As Integer
''  Dim intLen As Integer
''  dblLen = Len(strBefore)
''  For dblX = 1 To dblLen
''    strTest = Mid(strBefore, dblX, 1)
''    If Asc(strTest) < 48 Or Asc(strTest) > 57 Then
''      strTest = ""
''    End If
''    strAfter = strAfter & strTest
''  Next dblX
''  CleanString = strAfter
''  txt_getwt2.Text = CInt(Format(Val(CleanString), "#0"))
''''  MsgBox (CleanString)
''
''''  MsgBox (CInt(Format(Val(CleanString), "#0")))
''
''End Function


Private Sub Form_Load()
On Error GoTo err_handler:
    dt_ticket.Value = Now
    dt_ticket2.Value = Now
    
    partycode = 0
    compcode = 1
    yy = Year(Date) - 2000
    mm = Month(Date)
    dd = Day(Date)

    
    dmy = Trim(Str(yy)) + Right("0" + Trim(Str(mm)), 2) + Right("0" + Trim(Str(dd)), 2)
''    MsgBox (dmy)


'' change mwt = 1 for bypass the serialport , mwt = 0 for serial port mode
    mwt = 0
    
    If load_first_sec_type = "F" Then
       txt_vehicle.Visible = True
       cmb_vehicle.Visible = False
       cmb_load.Enabled = True
       lbl.Caption = "FIRST WEIGHT"
    Else
       txt_vehicle.Visible = False
       cmb_vehicle.Visible = True
       cmb_load.Enabled = False
       lbl.Caption = "SECOND WEIGHT"
    End If
    
    Call gen_dbconnection




    adocmd_mysql.ActiveConnection = gen_connection_mysql
    cmb_material.Clear


''    pst_qry = "select ifnull(max(t_wb_ticketno),0)+1 as tikcetno  from trn_weighbridge_entry where t_wb_type = 'A' and t_wb_year = " & yy & "  and t_wb_compcode = " & compcode
''    adocmd_mysql.CommandText = pst_qry
''    Set adors = adocmd_mysql.Execute
''    If adors.RecordCount > 0 Then
''         txt_ticketNo.Text = adors("tikcetno")
''    End If
         
         
    ''pst_qry = "select * from mas_wb_item order by item_name"
    pst_qry = "select * from (select item_code, item_name, item_group , '99' as ordprint from mas_wb_item Union All select 22 as item_code, 'WASTE PAPER' as item_name, 1 as item_group , '1' as ordprint Union All select 12 as item_code, 'PAPER REEL' as item_name, 0 item_group , '2' as ordprint ) a1 order by ordprint,item_name"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_material.AddItem (adors("item_name"))
                 cmb_material.ItemData(cmb_material.NewIndex) = adors("item_code")
                 adors.MoveNext
        Next
    End If
         
         
         
    adors.Close
    
    cmb_supplier.Clear

    pst_qry = "select * from mas_wb_party order by party_name"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_supplier.AddItem (adors("party_name"))
                 cmb_supplier.ItemData(cmb_supplier.NewIndex) = adors("party_code")
                 adors.MoveNext
        Next
    End If
    adors.Close
    
    
    pst_qry = "select * from mas_wb_settings_timeoffice "
    adocmd_mysql.ActiveConnection = gen_connection_mysql
 ''    cmbPort.Clear

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
    
                 baudRate = adors("wb_baudrate")
                 portnumber = adors("wb_port")
                 adors.MoveNext
        Next
    End If
    adors.Close
          
    
    
    cmd_Refresh_Click
    
refresh_data
    

    
    saveflag = "new"
    
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
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
''
''Private Sub MSComm1_OnComm()
'' On Error GoTo err_handler
''''     If mwt = 0 Then
''        Dim Buffer, chkwt, chkwt3 As String
''        Buffer = MSComm1.Input
''        txt_wt_from_serailport.SelText = Buffer
''
''''        txt_getwt = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, " ") + 2, 6)
''''        txt_getwt = Mid(txt_wt_from_serailport.Text, 1, Len(txt_wt_from_serailport.Text) - 3)
''
''Dim cnt, cnt2 As Integer
''
''        cnt = InStr(txt_wt_from_serailport.Text, ":")
''        MsgBox (Len(txt_wt_from_serailport.Text))
''        MsgBox (cnt)
''
''''        txt_getwt.Text = Mid(txt_wt_from_serailport.Text, 1, 8)
''
''''        txt_getwt.Text = Mid(txt_getwt.Text, Len(txt_getwt.Text) - 4)
''
''        chkwt = Mid(txt_wt_from_serailport.Text, 1, 8)
''        txt_getwt.Text = Mid(chkwt, Len(chkwt) - 2)
''''
''''
''''Dim cnt, cnt2 As Integer
''''Dim findchar As String
''''Dim newchar As String
''''cnt = InStr(txt_wt_from_serailport.Text, Chr(172))
''''
''''findchar = Mid(txt_wt_from_serailport.Text, cnt + 1)
''''cnt2 = InStr(findchar, Chr(172))
''''
''''
''''
''''findchar = Mid(findchar, 1, cnt2 - 3)
''''
''''newchar = Mid(txt_wt_from_serailport.Text, 1, Len(txt_wt_from_serailport.Text) - 2)
''''
''''cnt = Len(newchar)
''''txtnew3.Text = newchar
''''
''''txt.Text = Mid(newchar, 1, cnt - 2)
''''
''''If Len(newchar) = 5 Then
''''   txt.Text = Right(newchar, 3)
''''ElseIf Len(newchar) = 7 Then
''''   txt.Text = Right(newchar, 4)
''''ElseIf Len(newchar) = 9 Then
''''   txt.Text = Right(newchar, 5)
''''ElseIf Len(newchar) = 10 Then
''''   txt.Text = Right(newchar, 6)
''''ElseIf Len(newchar) = 11 Then
''''   txt.Text = Right(newchar, 6)
''''
''''End If
''
''
''  ''  chkwt3 = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, ":") + 1, 8)
''  ''    txtnew.Text = chkwt3
''
''
''
''        txt_wt2.Text = CInt(Format(Val(txt_getwt.Text), "#0"))
''''        Round(Val(txt_getwt2.Text), 0)
'''' lbl_getwt.Caption = Round(Val(txt_getwt.Text), 0)
''        lbl_getwt.Caption = CInt(Format(Val(txt_getwt.Text), "#0"))
''''     End If
''    Exit Sub
''err_handler:
''    MsgBox ("Port Not opened..")
''    Exit Sub
''End Sub

''
''Private Sub MSComm1_OnComm()
'' On Error GoTo err_handler
'' Dim cnt, cnt2, cnt3, cnt4 As String
'' Dim ii As Long
'' Dim fwt As String
''
''       If mwt = 0 Then
''        Dim Buffer As String
''        Buffer = MSComm1.Input
''        txt_wt_from_serailport.SelText = Buffer
''
''''                cnt = InStr(txt_wt_from_serailport.Text, ":")
''
''''MsgBox (cnt)
''''If cnt = 8 Then
''''MsgBox ("ok")
''''Else
''''MsgBox ("zero")
''''End If
''
''
''''        MsgBox (cnt)
''
''''        txt_getwt.Text = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, vbCrLf) + 1, 6)
''''        txt_getwt.Text = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, " ") + 1, 6)
''        cnt = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, ":") + 1, 15)
''        cnt2 = LTrim(RTrim(Mid(cnt, InStr(cnt, ":") + 1, 15)))
''
''
''   ''   ii = InStr(cnt2, ":")
''
''      cnt2 = Mid(LTrim(RTrim(cnt2)), 1, Len(cnt2) - 2)
''
''     ''   cnt3 = Mid(LTrim(RTrim(cnt2)), 1, ii - 2)
''
''
''
''''        MsgBox (cnt3)
''''                  MsgBox (cnt2)
''''If Len(cnt2) = 5 Then
''''   cnt3 = 0
''''If Len(cnt2) = 9 Then
''''
''''
''''        End If
''''MsgBox (Len(cnt2))
''       If Len(cnt2) = 5 Then
''          fwt = Mid(cnt2, 3, 3)
''      ''    fwt = Mid(fwt, 2, 3)
''       ElseIf Len(cnt2) = 7 Then
''          fwt = Mid(cnt2, 3, 3)
''        ElseIf Len(cnt2) = 9 Then
''          fwt = Mid(cnt2, 6, 6)
''       End If
''
''
'''
''       txt_getwt.Text = cnt2
''''        txt_getwt = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, ":") + 1, 10)
''        txt_wt2.Text = CInt(Format(Val(txt_getwt.Text), "#0"))
''''        Round(Val(txt_getwt2.Text), 0)
'''' lbl_getwt.Caption = Round(Val(txt_getwt.Text), 0)
''        lbl_getwt.Caption = CInt(Format(Val(txt_getwt.Text), "#0"))
''
''    ''    mwt = 1
''
''       End If
''    Exit Sub
''err_handler:
''   '' MsgBox ("Port Not opened..")
''    Exit Sub
''End Sub

Private Sub Text2_Change()

End Sub

Private Sub MSComm1_OnComm()
 On Error GoTo err_handler
 Dim cnt, cnt2, cnt3, cnt4 As String
 Dim ii As Long
 Dim fwt As String

       If mwt = 0 Then
        Dim Buffer As String
        Buffer = MSComm1.Input
        txt_wt_from_serailport.SelText = Buffer




    Dim f As Integer
    Dim str1, str2, wt  As String
    str1 = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, "=") + 2, 5)
    
    
    
    str1 = StrReverse(str1)
    
''  lll7MsgBox (str1)
    wt = str1

    txt_getwt.Text = wt
''        txt_getwt = Mid(txt_wt_from_serailport.Text, InStr(txt_wt_from_serailport.Text, ":") + 1, 10)
''        txt_wt2.Text = CInt(Format(Val(txt_getwt.Text), "#0"))
''''        Round(Val(txt_getwt2.Text), 0)
'''' lbl_getwt.Caption = Round(Val(txt_getwt.Text), 0)
''        lbl_getwt.Caption = CInt(Format(Val(txt_getwt.Text), "#0"))
''
''    ''    mwt = 1


        txt_wt2.Text = CLng(txt_getwt.Text)
''        Round(Val(txt_getwt2.Text), 0)
'' lbl_getwt.Caption = Round(Val(txt_getwt.Text), 0)
        lbl_getwt.Caption = CLng(txt_getwt.Text)
       End If
    Exit Sub
err_handler:
   '' MsgBox ("Port Not opened..")U
    Exit Sub
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Option1_Click()

End Sub

Private Sub Timer1_Timer()
   If mwt = 1 Then Exit Sub
   REFRESH_WEIGHT
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
                  Delay (500)
                Next i
                .Output = "AT+CSQ" & vbCr
                Delay (1000)
                .Output = "ATO & vbCr   'return to connected mode"
            End If
        End With
   
    Exit Sub
err_handler:
    MsgBox ("PORT -  Not opened..")
    MsgBox ("ALREADY THE SAME PROGRAM WAS OPENED.. PLEASE CHECK ")
    MsgBox ("NOW SYSTEM WILL EXIT THIS PROGRAM. PLEASE CHECK AND REOPEN THE PROGRAM")
    End
''    Exit Sub

End Sub
Private Sub cmd_on_Click()
    mwt = 1
End Sub


Private Sub txt_LoadWt_Change()
''       txt_NetWt.Text = Val(txt_LoadWt.Text) - Val(txt_EmptyWt.Text)
End Sub

Private Sub txt_vehicle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmb_load.SetFocus
End If
End Sub
