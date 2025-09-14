VERSION 5.00
Begin VB.Form frm_itemgroup_master 
   Caption         =   "ITEM GROUP MASTER"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   13800
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   840
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   5055
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Edit"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&New"
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
         TabIndex        =   7
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   9615
      Begin VB.TextBox txt_itemgroup 
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
         TabIndex        =   2
         Top             =   4440
         Width           =   5895
      End
      Begin VB.ComboBox cmb_item_group 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3540
         Left            =   3120
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Material Grp to be Changed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Material Group"
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_itemgroup_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mwt As Integer
Dim adocmd_mysql As New ADODB.Command
Dim adors As New ADODB.Recordset
Dim pst_qry  As String

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
Dim fst_save As String


Private Sub cmd_Edit_Click()
fst_save = "EDIT"
Label2.Visible = True
txt_itemgroup.Visible = True
End Sub

Private Sub cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_New_Click()
Label2.Visible = False
txt_itemgroup.Visible = False
End Sub

Private Sub cmd_Refresh_Click()
Label2.Visible = False
txt_itemgroup.Visible = False
    cmb_item_group.Clear

    pst_qry = "select * from mas_wb_itemgroup order by item_grpname"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    While Not adors.EOF
       cmb_item_group.AddItem adors("item_grpname")
       adors.MoveNext
    Wend
    adors.Close
End Sub

Private Sub cmd_Save_Click()

On Error GoTo err_handler
Dim pst_respo As String
Dim pst_qry As String
Dim no As Integer
If fst_save = "NEW" Then
    If cmb_item_group.Text = "" Then
        MsgBox " Enter Group Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_item_group.SetFocus
        Exit Sub
    End If
    pst_respo = MsgBox("Do You want to save the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        MousePointer = vbDefault
        Exit Sub
    End If
    Dim rs_pay As New ADODB.Recordset
    pst_qry = "select * from mas_wb_itemgroup where item_grpname='" & UCase(cmb_item_group.Text) & "'"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        MsgBox ("Item Group Already Saved...")
        Exit Sub
    End If
        
    no = 1
    pst_qry = "select max(item_grpcode)+1 as endno from mas_wb_itemgroup"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
         no = adors!endno
    End If
    
    pst_qry = "insert into mas_wb_itemgroup values(" & no & ",'" & UCase(cmb_item_group.Text) & "')"
    adocmd_mysql.CommandText = pst_qry
    adocmd_mysql.Execute pst_qry
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"
    cmb_item_group.Text = ""
    Exit Sub
ElseIf fst_save = "EDIT" Then
    If cmb_item_group.Text = "" Then
        
        MsgBox " Enter Group Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_item_group.SetFocus
        Exit Sub
    End If
    If txt_itemgroup.Text = "" Then
        MsgBox " Enter Group Name to be Change as ", vbOKOnly + vbExclamation, "vbInformation "
        txt_itemgroup.SetFocus
        Exit Sub
      End If
    pst_respo = MsgBox("Do You want to Modify the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        MousePointer = vbDefault
        Exit Sub
    End If

    pst_qry = "update mas_wb_itemgroup set item_grpname='" & UCase(txt_itemgroup.Text) & "' where item_grpname='" & UCase(cmb_item_group.Text) & "'"
    adocmd_mysql.CommandText = pst_qry
    adocmd_mysql.Execute pst_qry
    MsgBox "Records Updated", vbOKOnly + vbInformation, "vbInformation"
    cmb_item_group.Text = ""
    txt_itemgroup.Text = ""

End If
cmd_Refresh_Click
Exit Sub
err_handler:

        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub Command1_Click()
    Dim f As Integer
    Dim str1, str2, wt  As String
    str1 = Text1.Text
    
    f = InStrRev(str1, "40")
    If f > 0 Then
       str2 = Mid(str1, 1, f - 1)
    Else
        str2 = str1
    End If
    
    If Len(str2) = 5 Then
       wt = Mid(str2, 3, 3)
    ElseIf Len(str2) = 7 Then
       wt = Mid(str2, 4, 4)
    ElseIf Len(str2) = 9 Then
       wt = Mid(str2, 5, 5)
    ElseIf Len(str2) = 11 Then
       wt = Mid(str2, 6, 6)
      
       
    End If
    MsgBox (wt)


End Sub

Private Sub Form_Load()



fst_save = "NEW"
Label2.Visible = False
txt_itemgroup.Visible = False
    
    Call gen_dbconnection
  
    compcode = 1
    '' fincode = 23
    
    Dim pin_cnt As Integer
    pst_qry = "select * from mas_wb_itemgroup"
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    cmb_item_group.Clear


    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_item_group.AddItem (adors("item_grpname"))
                 cmb_item_group.ItemData(cmb_item_group.NewIndex) = adors("item_grpcode")
                 adors.MoveNext
        Next
    End If
    adors.Close

    
End Sub


