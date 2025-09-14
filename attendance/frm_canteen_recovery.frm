VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_canteen_recovery 
   Caption         =   "CANTEEN RECOVERY DAILY ENTRY"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   14520
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4560
      TabIndex        =   9
      Top             =   9360
      Width           =   3855
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_canteen_recovery.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_canteen_recovery.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "frm_canteen_recovery.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "frm_canteen_recovery.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "frm_canteen_recovery.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid flx_data 
         Height          =   5175
         Left            =   600
         TabIndex        =   18
         Top             =   3000
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin VB.TextBox txt_others 
         Height          =   405
         Left            =   8040
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txt_empname 
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CommandButton cmd_add 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Food Type"
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
         Height          =   975
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   6135
         Begin VB.ComboBox cmb_food 
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
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.ComboBox cmb_empname 
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
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txt_empcode 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Others"
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
         Index           =   2
         Left            =   8040
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp. Name"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp. Code"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker dt_entry 
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   121569281
      CurrentDate     =   44592
   End
   Begin VB.Label lbl 
      Caption         =   " Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6960
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frm_canteen_recovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_empname_Change()
    If cmb_empname.ListIndex = -1 Then Exit Sub
    txt_empcode.Text = cmb_empname.ItemData(cmb_empname.ListIndex)
    txt_empname.Text = cmb_empname.Text
End Sub

Private Sub cmb_empname_Click()
If cmb_empname.ListIndex = -1 Then Exit Sub
txt_empcode.Text = cmb_empname.ItemData(cmb_empname.ListIndex)
End Sub

Private Sub cmd_add_Click()
    If cmb_food.Text = "" Then
       MsgBox ("Select Food type...")
       Exit Sub
    End If
        
    If txt_empcode.Text = "" Then
       MsgBox ("Enter Employee code ...")
       Exit Sub
    End If
    If txt_empname.Text = "" Then
       MsgBox ("Enter Employee code ...")
       Exit Sub
    End If
    
    sno = 0
    With flx_data
       For i = 1 To flx_data.Rows - 1
           .TextMatrix(i, 0) = i
           If Val(.TextMatrix(i, 0)) > 0 Then sno = Val(.TextMatrix(i, 0))
            If Val(.TextMatrix(i, 1)) = Val(txt_empcode.Text) And Trim(.TextMatrix(i, 3)) = Trim(cmb_food.Text) Then
               MsgBox ("Employee Code Already Selected...")
               Exit Sub
            End If
            
       Next
       
       .Rows = .Rows + 1
       .Row = .Rows - 1

       .TextMatrix(.Row - 1, 1) = txt_empcode.Text
       .TextMatrix(.Row - 1, 2) = txt_empname.Text
       .TextMatrix(.Row - 1, 3) = cmb_food.Text
       .TextMatrix(.Row - 1, 4) = txt_others.Text
        
        txt_empname.Text = ""
        txt_empcode.Text = ""
    End With

End Sub

Private Sub dt_entry_Change()
   editdata
End Sub

Private Sub dt_entry_Click()
editdata
End Sub

Private Sub dt_entry_LostFocus()
editdata
End Sub

Private Sub dt_entry_OLEDragOver(Data As MSComCtl2.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
editdata
End Sub

Private Sub edit_Click()
    editdata
End Sub

Sub editdata()
    fillgrid
    sql = "select * from canteen_recovery a, emp_mas b where cr_empcode = emp_code and cr_date = '" & Format(dt_entry, "MM/dd/yyyy") & "' order by cr_sno "
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
                flx_data.Rows = flx_data.Rows + 1
                flx_data.Row = flx_data.Rows - 1
                flx_data.TextMatrix(flx_data.Row - 1, 0) = payrs.Fields("cr_sno")
                flx_data.TextMatrix(flx_data.Row - 1, 1) = payrs.Fields("cr_empcode")
                flx_data.TextMatrix(flx_data.Row - 1, 2) = payrs.Fields("emp_name")
                flx_data.TextMatrix(flx_data.Row - 1, 4) = payrs.Fields("cr_others")
                If payrs.Fields("cr_foodtype") = "B" Then
                   flx_data.TextMatrix(flx_data.Row - 1, 3) = "BREAKFAST"
                ElseIf payrs.Fields("cr_foodtype") = "L" Then
                   flx_data.TextMatrix(flx_data.Row - 1, 3) = "LUNCH"
                ElseIf payrs.Fields("cr_foodtype") = "D" Then
                   flx_data.TextMatrix(flx_data.Row - 1, 3) = "DINNER"
                                Else
                   flx_data.TextMatrix(flx_data.Row - 1, 3) = "OTHERS"
                End If
        
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
     payrs.Close

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub flx_data_DblClick()
   flex_edit_row = 0
   Dim fin_selrow As Integer
   Dim pst_ans As String
   fin_selrow = flx_data.Row

   timchk = 0
   With flx_data
       pst_ans = MsgBox("Press YES-to DELETE  NO-to CANCEL", vbYesNo, "Confirmation")
       If pst_ans = 6 Then
               If .Rows < 2 Then
                  MsgBox "No rows to remove"
               Else
                  If flx_data.TextMatrix(.Row, 1) <> "" Then
                     flx_data.RemoveItem fin_selrow
                     .Row = flx_data.Rows - 1
                  End If
               End If
        End If
   End With
   For i = 1 To flx_data.Rows - 1
      flx_data.TextMatrix(i, 0) = i
   Next
   
End Sub

Private Sub flx_data_KeyPress(KeyAscii As Integer)
 On Error GoTo err_handler

 Dim fin_selrow%, fin_selcol%
 fin_selrow = flx_data.Row
 fin_selcol = flx_data.Col
 With flx_data
    If fin_selcol = 4 Then
        If KeyAscii <> 13 Then
            KeyAscii = Numeric_Chk(KeyAscii, flx_data.TextMatrix(fin_selrow, fin_selcol), 8, 5, 2)
            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                flx_data.TextMatrix(fin_selrow, fin_selcol) = flx_data.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
            ElseIf KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then .TextMatrix(fin_selrow, fin_selcol) = Mid(.TextMatrix(fin_selrow, fin_selcol), 1, Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              KeyAscii = 0
            End If
         End If
    End If
End With
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If


End Sub

Private Sub Form_Load()
    dt_entry.Value = Now
    fillgrid
    
    sql = "Select * from  bio_empmas where bioemp_status = 'Working' order by bioemp_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_empname.AddItem payrs("bioemp_name")
        cmb_empname.ItemData(cmb_empname.NewIndex) = payrs("bioemp_fpcode")
        payrs.MoveNext
    Wend
    payrs.Close
    
    cmb_food.AddItem "BREAKFAST"
    cmb_food.AddItem "LUNCH"
    cmb_food.AddItem "DINNER"
    cmb_food.AddItem "OTHERS"
End Sub

Private Sub fillgrid()
   With flx_data
     .Redraw = False
     .Clear
     .Rows = 2
     .Cols = 5
     .TextMatrix(0, 0) = "S.No"
     .TextMatrix(0, 1) = "Emp-Code"
     .TextMatrix(0, 2) = "Name"
     .TextMatrix(0, 3) = "Food"
     .TextMatrix(0, 4) = "Others"
     
     .ColWidth(0) = 800
     .ColWidth(1) = 1400
     .ColWidth(2) = 5000
     .ColWidth(3) = 1500
     .ColWidth(4) = 1000
     .Redraw = True
     
   End With

End Sub

Private Sub save_Click()
       
On Error GoTo err_handler
  Me.MousePointer = 11
  paydb.BeginTrans
  
       
       
       sql = "delete from canteen_recovery where cr_date = '" & Format(dt_entry, "MM/dd/yyyy") & "'"
       paydb.Execute sql
       
       
       For i = 1 To flx_data.Rows - 1
           If flx_data.TextMatrix(i, 1) <> "" Then
              sql = "insert into canteen_recovery (cr_date,cr_sno,cr_empcode,cr_foodtype,cr_others) values ('" & Format(dt_entry.Value, "MM/dd/yyyy") & "', " & Val(flx_data.TextMatrix(i, 0)) & ", " & Val(flx_data.TextMatrix(i, 1)) & ", '" & Left(flx_data.TextMatrix(i, 3), 1) & "', " & Val(flx_data.TextMatrix(i, 4)) & ")"
              paydb.Execute sql
           End If
       Next
       paydb.CommitTrans
       MsgBox ("Records are saved..")
       fillgrid

  Me.MousePointer = 1
  Exit Sub
err_handler:
    paydb.RollbackTrans
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub txt_empcode_Change()
   txt_empname.Text = ""
    sql = "Select * from  bio_empmas where bioemp_fpcode =  " & Val(txt_empcode.Text)
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        txt_empname.Text = payrs("bioemp_name")
        payrs.MoveNext
    Wend
    payrs.Close

End Sub

Private Sub txt_empcode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmd_add_Click
    End If
End Sub
