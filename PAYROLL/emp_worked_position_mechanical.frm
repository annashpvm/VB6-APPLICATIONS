VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form emp_worked_position_mechanical 
   Caption         =   "Employee worked Position"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3840
      TabIndex        =   1
      Top             =   9240
      Width           =   3855
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "emp_worked_position_mechanical.frx":0000
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
         Picture         =   "emp_worked_position_mechanical.frx":0442
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
         Picture         =   "emp_worked_position_mechanical.frx":0AAC
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
         Picture         =   "emp_worked_position_mechanical.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "emp_worked_position_mechanical.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx_position 
      Height          =   8250
      Left            =   -240
      TabIndex        =   0
      Top             =   720
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   14552
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   4
      BackColorFixed  =   16776960
      BackColorSel    =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbldept 
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "emp_worked_position_mechanical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub


Private Sub flx_position_KeyPress(KeyAscii As Integer)
    On Error GoTo err_handler

    Dim fin_selrow%, fin_selcol%
    fin_selrow = flx_position.Row
    fin_selcol = flx_position.Col

    With flx_position
        ' Ensure selected cell is valid
        If fin_selrow < 0 Or fin_selcol < 0 Then Exit Sub
        
        If fin_selcol <= 3 Then
            KeyAscii = 0
            Exit Sub
        End If
        
        
        Select Case KeyAscii
            Case 48 To 57 ' ASCII 0–9
                .TextMatrix(fin_selrow, fin_selcol) = .TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)

            Case 8 ' Backspace
                If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
                    .TextMatrix(fin_selrow, fin_selcol) = Left$(.TextMatrix(fin_selrow, fin_selcol), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                End If

            Case 13 ' Enter key
                ' Move to next column
                If fin_selcol < .Cols - 1 Then
                    .Col = fin_selcol + 1
                ElseIf fin_selrow < .Rows - 1 Then
                    ' If at end of row, move to first column of next row
                    .Row = fin_selrow + 1
                    .Col = 0
                End If
                KeyAscii = 0 ' Suppress the default Enter key behavior

            Case Else
                MsgBox "Enter numbers only", vbExclamation
                KeyAscii = 0 ' Cancel invalid input
        End Select
    End With

    Exit Sub

err_handler:
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
End Sub

Private Sub Form_Load()
     fillgrid
     filldata
End Sub

Function fillgrid()
    With flx_position
        .Clear
        .Cols = 13
        .Rows = 1
        .TextMatrix(0, 0) = "Sl.No"
        .TextMatrix(0, 1) = "Emp code"
        .TextMatrix(0, 2) = "Emp Name"
        .TextMatrix(0, 3) = "D.O.J "
        .TextMatrix(0, 4) = "Press "
        .TextMatrix(0, 5) = "Wire "
        .TextMatrix(0, 6) = "Reliver"
        .TextMatrix(0, 7) = "Dryer"
        .TextMatrix(0, 8) = "Ist Asst"
        .TextMatrix(0, 9) = "Reliever"
        .TextMatrix(0, 10) = "Ist Operator"
        .TextMatrix(0, 11) = "Shift Incharge"
        .TextMatrix(0, 12) = "Reliever"
         
        .ColWidth(0) = 700
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1200
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(3) = 4
        Dim i As Integer
        For i = 4 To .Cols - 1
            .ColAlignment(i) = 4
        Next i
        End With
End Function



Function filldata()
  Set payrs = New ADODB.Recordset
sql = "select emp_fpcode,emp_name,emp_doj,dept_name from emp_mas ,pdept_mas where emp_dept = dept_code and emp_company = 1 and emp_status = 'A' and emp_dept = 34"
paydb.CommandTimeout = 300
payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
i = 1

If Not payrs.EOF Then
lbldept.Caption = payrs("dept_name")
   While Not payrs.EOF
             
      '' if ds_fpcode =
      
        With flx_position
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = i
            .TextMatrix(.Rows - 1, 1) = payrs("emp_fpcode")
            .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
            .TextMatrix(.Rows - 1, 3) = payrs("emp_doj")
            i = i + 1
        End With
        
        payrs.MoveNext
    Wend
    payrs.Close
End If

sql = "select * from emp_workposition_history_Production where p_compcode = 1 and p_deptcode = 34"
payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
             For i = 1 To flx_position.Rows - 1
                 If Trim(flx_position.TextMatrix(i, 1)) <> "" Then
                      If Val(flx_position.TextMatrix(i, 1)) = payrs.Fields("p_empcode") Then
                         flx_position.TextMatrix(i, 4) = payrs.Fields("p_press")
                         flx_position.TextMatrix(i, 5) = payrs.Fields("p_wire")
                         flx_position.TextMatrix(i, 6) = payrs.Fields("p_dryer")
                         flx_position.TextMatrix(i, 7) = payrs.Fields("p_ist_asst")
                         flx_position.TextMatrix(i, 8) = payrs.Fields("p_ist_oper")
                         flx_position.TextMatrix(i, 9) = payrs.Fields("p_shiftincharge")
                         flx_position.TextMatrix(i, 10) = 0
                         flx_position.TextMatrix(i, 11) = 0
                         
                         
                      End If
                      
                 End If
             Next
             payrs.MoveNext
        Wend
   End If

End Function



Private Sub refresh_Click()
     fillgrid
     filldata
End Sub

Private Sub SAVE_Click()

Me.MousePointer = 11
Set payrs = New ADODB.Recordset
millcode = 1
  sql = "delete from emp_workposition_history where p_compcode = " & millcode & " and p_deptcode = 3400"
  paydb.Execute sql
  sql = "select * from emp_workposition_history where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 1 To flx_position.Rows - 1
            payrs.AddNew
            payrs.Fields("p_compcode") = millcode
            payrs.Fields("p_deptcode") = 34
            payrs.Fields("p_empcode") = flx_position.TextMatrix(i, 1)
            payrs.Fields("p_press") = Val(flx_position.TextMatrix(i, 4))
            payrs.Fields("p_wire") = Val(flx_position.TextMatrix(i, 5))
            payrs.Fields("p_dryer") = Val(flx_position.TextMatrix(i, 6))
            payrs.Fields("p_ist_asst") = Val(flx_position.TextMatrix(i, 7))
            payrs.Fields("p_ist_oper") = Val(flx_position.TextMatrix(i, 8))
            payrs.Fields("p_shiftincharge") = Val(flx_position.TextMatrix(i, 9))
            
            
            payrs.Update
  Next
  MsgBox ("Records are saved")
  payrs.Close
  fillgrid
  Me.MousePointer = 1
  Exit Sub
err_handler:
   
    Me.MousePointer = 1
    chk = gen_Validation(Err.Number, Err.Description)
    If chk = 1 Then Resume
    
End Sub



