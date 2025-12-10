VERSION 5.00
Begin VB.Form PortSetting 
   Caption         =   "COM PORT SETTING"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txt_baudrate 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox cmbPort 
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
      ItemData        =   "PortSetting.frx":0000
      Left            =   3480
      List            =   "PortSetting.frx":0010
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "BAUD RATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "COM PORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "PortSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err_handler:
    If Val(cmbPort.Text) = 0 Then
       MsgBox ("Select Port..")
       Exit Sub
    End If
    
    If Val(txt_baudrate.Text) = 0 Then
       MsgBox ("Enter Baud Rate..")
       Exit Sub
    End If
        
        
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    Dim pst_qry As String
    
        
    adocmd_mysql.ActiveConnection = gen_connection_mysql

    
        pst_qry = "delete from mas_wb_settings_finishing where wb_port > 0"
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry
        
        
        pst_qry = "insert into mas_wb_settings_finishing values ( " & Val(cmbPort.Text) & ", " & Val(txt_baudrate.Text) & ")"
        adocmd_mysql.CommandText = pst_qry
        adocmd_mysql.Execute pst_qry
        MsgBox ("Port Settings Saved ...")
    Exit Sub
err_handler:
    If gen_Validation(Err.Number, Err.Description) = 1 Then Resume
    
End Sub

Private Sub Form_Load()
      txt_baudrate.Text = "1200"
      
    Dim adocmd_mysql As New ADODB.Command
    Dim adors As New ADODB.Recordset
    
    Call gen_dbconnection
  
    
    Dim pin_cnt As Integer
    pst_qry = "select * from mas_wb_settings_finishing "
    adocmd_mysql.ActiveConnection = gen_connection_mysql
 ''    cmbPort.Clear

    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
   ''              cmbPort.AddItem (adors("wb_port"))
                 txt_baudrate.Text = adors("wb_baudrate")
                 cmbPort.Text = adors("wb_port")
                 adors.MoveNext
        Next
    End If
    adors.Close

      
      
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If txtPassword.Text = "MIS" Then
          cmdSave.Enabled = True
      Else
          cmdSave.Enabled = False
      End If
    End If
    
End Sub
