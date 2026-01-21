VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_item_master 
   Caption         =   "ITEM MASTER"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   13935
      Begin MSFlexGridLib.MSFlexGrid flx_item 
         Height          =   5295
         Left            =   360
         TabIndex        =   13
         Top             =   2760
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9340
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmb_item_name 
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
         Left            =   3240
         TabIndex        =   11
         Top             =   480
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
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox txt_itemname 
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
         TabIndex        =   7
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label Label3 
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
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Material"
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
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Material to be Changed"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   5055
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "frm_item_master"
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


Private Sub cmb_item_name_Click()
    pst_qry = "select * from mas_wb_item where item_code = '" & cmb_item_name.ItemData(cmb_item_name.ListIndex) & "'"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
          cmb_item_group.ListIndex = find_index_item_data(cmb_item_group, adors("item_group"))
        Exit Sub
    End If
End Sub


Private Sub cmd_Edit_Click()
fst_save = "EDIT"
Label2.Visible = True
txt_itemname.Visible = True
End Sub

Private Sub cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_New_Click()
Label2.Visible = False
txt_itemname.Visible = False
End Sub

Private Sub cmd_Refresh_Click()
fst_save = "NEW"
Label2.Visible = False
txt_itemname.Visible = False
''    cmb_item_group.Clear
''
''    pst_qry = "select * from mas_wb_itemgroup order by item_grpname"
''    adocmd_mysql.CommandText = pst_qry
''    Set adors = adocmd_mysql.Execute
''    While Not adors.EOF
''       cmb_item_group.AddItem adors("item_grpname")
''       adors.MoveNext
''    Wend
''    adors.Close
End Sub

Private Sub cmd_Save_Click()

On Error GoTo err_handler
Dim pst_respo As String
Dim pst_qry As String
Dim no As Integer
If fst_save = "NEW" Then
    If cmb_item_name.Text = "" Then
        MsgBox " Enter Item Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_item_name.SetFocus
        Exit Sub
    End If
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

    pst_qry = "select * from mas_wb_item where item_name='" & UCase(cmb_item_name.Text) & "'"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        MsgBox ("Item Name Already Saved...")
        Exit Sub
    End If
        
    no = 1
    pst_qry = "select max(item_code)+1 as endno from mas_wb_item"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
         no = adors!endno
    End If
    
    pst_qry = "insert into mas_wb_item values(" & no & ",'" & UCase(cmb_item_name.Text) & "'," & cmb_item_group.ItemData(cmb_item_group.ListIndex) & ")"
    adocmd_mysql.CommandText = pst_qry
    adocmd_mysql.Execute pst_qry
    MsgBox "Record Saved", vbOKOnly + vbInformation, "Information"

    Exit Sub
ElseIf fst_save = "EDIT" Then
    If cmb_item_name.Text = "" Then
        
        MsgBox " Enter Item Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_item_name.SetFocus
        Exit Sub
    End If
    If cmb_item_group.Text = "" Then
        MsgBox " Enter Group Name ", vbOKOnly + vbExclamation, "vbInformation "
        cmb_item_group.SetFocus
        Exit Sub
    End If
        
    If txt_itemname.Text = "" Then
        MsgBox " Enter Item Name to be Change as ", vbOKOnly + vbExclamation, "vbInformation "
        txt_itemname.SetFocus
        Exit Sub
      End If
    pst_respo = MsgBox("Do You want to Modify the record", vbYesNo + vbInformation, "Information")
    If pst_respo = vbNo Then
        MousePointer = vbDefault
        Exit Sub
    End If

    pst_qry = "update mas_wb_item set item_name='" & UCase(txt_itemname.Text) & "' , item_group ='" & cmb_item_group.ItemData(cmb_item_group.ListIndex) & "'  where item_name='" & UCase(cmb_item_name.Text) & "'"
    adocmd_mysql.CommandText = pst_qry
    adocmd_mysql.Execute pst_qry
    MsgBox "Records Updated", vbOKOnly + vbInformation, "vbInformation"
    txt_itemname.Text = ""

End If
cmd_Refresh_Click
Exit Sub
err_handler:

        chk = gen_Validation(Err.Number, Err.Description)
End Sub

Private Sub flx_item_Click()
 Dim fin_selrow%, fin_selcol%
 fin_selrow = flx_item.Row
 fin_selcol = flx_item.Col
'' MsgBox (flx_item.TextMatrix(fin_selrow, 0))
  cmb_item_name.ListIndex = find_index_item_data(cmb_item_name, flx_item.TextMatrix(fin_selrow, 0))
  cmb_item_group.ListIndex = find_index_item_data(cmb_item_group, flx_item.TextMatrix(fin_selrow, 2))
End Sub

Private Sub Form_Load()
fst_save = "NEW"
Label2.Visible = False
txt_itemname.Visible = False
fst_save = "NEW"
    Call gen_dbconnection
    fillgrid
    compcode = 1
    '' fincode = 23
    
    Dim pin_cnt As Integer

    adocmd_mysql.ActiveConnection = gen_connection_mysql
    cmb_item_group.Clear

    pst_qry = "select * from mas_wb_item order by item_name"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
                 cmb_item_name.AddItem (adors("item_name"))
                 cmb_item_name.ItemData(cmb_item_name.NewIndex) = adors("item_code")
                 adors.MoveNext
        Next
    End If
    adors.Close

    pst_qry = "select * from mas_wb_itemgroup order by item_grpname"
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




Function fillgrid()

    With flx_item
        .Clear
        .Cols = 4
        .Rows = 1

        .TextMatrix(0, 0) = "Item code"
        .TextMatrix(0, 1) = "Item Name"
        .TextMatrix(0, 2) = "Group Code"
        .TextMatrix(0, 3) = "Group Name"
                
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 3500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3500
    End With
    
   Dim pin_cnt As Integer
    adocmd_mysql.ActiveConnection = gen_connection_mysql
    pst_qry = "select * from mas_wb_item ,  mas_wb_itemgroup where item_group = item_grpcode"
    adocmd_mysql.CommandText = pst_qry
    Set adors = adocmd_mysql.Execute
    If adors.RecordCount > 0 Then
        For pin_cnt = 1 To adors.RecordCount
            With flx_item
                 .Rows = .Rows + 1
                
                 .TextMatrix(.Rows - 1, 0) = adors("item_code")
                 .TextMatrix(.Rows - 1, 1) = adors("item_name")
                 .TextMatrix(.Rows - 1, 2) = adors("item_grpcode")
                 .TextMatrix(.Rows - 1, 3) = adors("item_grpname")
                 
            End With
            adors.MoveNext
        Next
  End If
End Function

