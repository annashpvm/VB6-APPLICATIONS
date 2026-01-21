VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_overtime 
   Caption         =   "WORKERS OVER TIME ENTRY"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   14460
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   14160
      TabIndex        =   18
      Top             =   1800
      Width           =   4335
      Begin VB.ComboBox cmb_dept 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   4095
      End
      Begin VB.OptionButton opt_selective_dept 
         Caption         =   "SELECTIVE DEPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   3735
      End
      Begin VB.OptionButton opt_all_dept 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   11520
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton opt_hrs_all 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opt_hrs_extraonly 
         Caption         =   "EXCESS HOURS ONL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
      Begin VB.OptionButton opt_all 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_cs 
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton opt_worker 
         Caption         =   "WORKER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   4320
      TabIndex        =   5
      Top             =   8880
      Width           =   3855
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "frm_prodn_incentive.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Refresh"
         Height          =   705
         Left            =   2280
         MaskColor       =   &H000000FF&
         Picture         =   "frm_prodn_incentive.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Height          =   705
         Left            =   1560
         MaskColor       =   &H000000FF&
         Picture         =   "frm_prodn_incentive.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit"
         Height          =   705
         Left            =   840
         MaskColor       =   &H000000FF&
         Picture         =   "frm_prodn_incentive.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
         Height          =   705
         Left            =   120
         MaskColor       =   &H000000FF&
         Picture         =   "frm_prodn_incentive.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox cmb_millname 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dt_ot 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   129368065
         CurrentDate     =   44558
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid att_flex 
      Height          =   6570
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   11589
      _Version        =   393216
      Rows            =   4
      Cols            =   8
      FixedCols       =   7
      BackColorFixed  =   16776960
      BackColorSel    =   -2147483624
      BackColorBkg    =   12632256
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   7095
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   13215
   End
End
Attribute VB_Name = "frm_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fst_save As Integer
Dim opt As Integer
Dim datefrom, dateto As Date
Dim emptype As String
Dim getotday As Integer
Dim mmon, myear As Integer

Dim overtime_eligible As Integer

Private Sub cmb_month_Change()
    dt_ot.Value = Now - Day(Now) + 1
End Sub

Private Sub dt_ot_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    fillgrid
    

End Sub

Private Sub dt_ot_Change()
    fillgrid
      
    monthcheck
''    MsgBox (end_date)

End Sub

Function monthcheck()
    mmon = Month(dt_ot.Value)
    myear = Year(dt_ot.Value)
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And myear Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + Str(myear))
    
    If Format(dt_ot.Value, "MM/dd/yyyy") = Format(end_date, "MM/dd/yyyy") Then
''       MsgBox ("OK")
         getotday = 1
    Else
''       MsgBox ("NOT OK")
         getotday = 0
    End If
    
End Function

Private Sub dt_ot_Click()
  '' MsgBox (Month(dt_ot.Value))
End Sub





Private Sub edit_Click()
    fst_save = 1
    
If cmb_millname.Text = "SHVPM" Then
    millcode = 1
End If

millcode = 1

monthcheck

    fillgrid
    filldata
    i = 1

    Set payrs = New ADODB.Recordset
    
    sql = "select *  from bio_worker_daily_pihrs where  w_company = " & millcode & " and w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
    
    If opt_all.Value = True Then
       sql = "select *  from bio_worker_daily_pihrs where w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
    ElseIf opt_worker.Value = True Then
       sql = "select *  from bio_worker_daily_pihrs where w_cat = 'W' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "' order by  w_emp_fpcode"
    ElseIf opt_cs.Value = True Then
       sql = "select *  from bio_worker_daily_pihrs where w_cat = 'C' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "' order by  w_emp_fpcode"
    End If
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    If Not payrs.EOF Then
       While Not payrs.EOF
    
        
             For i = 1 To att_flex.Rows - 1
                 If Trim(att_flex.TextMatrix(i, 3)) <> "" Then
                      If Val(att_flex.TextMatrix(i, 3)) = payrs.Fields("w_emp_fpcode") Then
''                         att_flex.TextMatrix(i, 18) = payrs.Fields("w_tot_ot_hrs")
                         att_flex.TextMatrix(i, 19) = payrs.Fields("w_wo_ot_hrs")
                         If payrs.Fields("w_act_hrs") > 0 Then
''                            att_flex.TextMatrix(i, 20) = payrs.Fields("w_act_hrs")
                         Else
  ''                       att_flex.TextMatrix(i, 20) = 0
                         End If
                         att_flex.TextMatrix(i, 22) = payrs.Fields("w_accepted_hrs")
                         att_flex.TextMatrix(i, 23) = payrs.Fields("w_holiday_ot_hrs")
                         


                      End If
                 End If
             Next
             payrs.MoveNext
        Wend
     Else
        MsgBox ("Details not available for the date ")
     End If
        
    call_ot_days

    payrs.Close
    '' paydb.close

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    overtime_eligible = 1
    getotday = 0
    cmb_millname.AddItem "SHVPM"
    cmb_millname.Text = "SHVPM"
    dt_ot.Value = Now - Day(Now) + 1
    
    Dim payrs As New ADODB.Recordset
    cmb_dept.Clear
    sql = "select * from pdept_mas where left(dept_name,2) != ' '  order by dept_name"
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF()
        cmb_dept.AddItem payrs("dept_name")
        cmb_dept.ItemData(cmb_dept.NewIndex) = payrs("dept_code")
        payrs.MoveNext
    Wend
    payrs.Close
    fillgrid

End Sub

Function fillgrid()
    With att_flex
        .Clear
        .Cols = 30
        .Rows = 1
        .TextMatrix(0, 0) = "Department"
        .TextMatrix(0, 1) = "Dep.code"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Emp Code"
        .TextMatrix(0, 4) = "Seh.Shift"
        .TextMatrix(0, 5) = "Act.Shift"
        .TextMatrix(0, 6) = "Work.HRS"
        .TextMatrix(0, 7) = "Intime1 "
        .TextMatrix(0, 8) = "Outtime1 "
        .TextMatrix(0, 9) = "Intime2 "
        .TextMatrix(0, 10) = "Outtime2 "
        .TextMatrix(0, 11) = "OD Intime "
        .TextMatrix(0, 12) = "OD outtime"
        .TextMatrix(0, 13) = "W.Hrs1"
        .TextMatrix(0, 14) = "W.Hrs2"
        .TextMatrix(0, 15) = "Tot.Hrs"
        .TextMatrix(0, 16) = "CH DATE"
        .TextMatrix(0, 17) = "CH Hrs"
        .TextMatrix(0, 18) = "Bal.Hrs"
        .TextMatrix(0, 19) = "WO OT"
        .TextMatrix(0, 20) = "BAL OT"
        .TextMatrix(0, 21) = "CAT"
        .TextMatrix(0, 22) = "OT-ACCEPTED"
        .TextMatrix(0, 23) = "HOLIDAY-OT"
        .TextMatrix(0, 24) = "Intime3 "
        .TextMatrix(0, 25) = "Outtime3 "
        .TextMatrix(0, 26) = "W.Hrs3"
        .TextMatrix(0, 27) = "M.WO.OT.Hrs"
        .TextMatrix(0, 28) = "P.Hrs"
        .TextMatrix(0, 29) = "OD.Hrs"
        .ColWidth(0) = 1200
        .ColWidth(1) = 0
        .ColWidth(2) = 2500
        .ColWidth(3) = 800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 2000
        .ColWidth(8) = 2000
        .ColWidth(9) = 2000
        .ColWidth(10) = 2000
        .ColWidth(11) = 10
        .ColWidth(12) = 10
        .ColWidth(13) = 900
        .ColWidth(14) = 900
        .ColWidth(15) = 900
        .ColWidth(16) = 1000
        .ColWidth(17) = 800
        .ColWidth(18) = 900
        .ColWidth(19) = 900
        .ColWidth(20) = 900
        .ColWidth(21) = 0
        .ColWidth(22) = 1000
        .ColWidth(23) = 1000
        .ColWidth(24) = 1000
        .ColWidth(25) = 1000
        .ColWidth(26) = 1100
        .ColWidth(27) = 1100
        .ColWidth(28) = 1100
    End With
End Function
Private Sub att_flex_KeyPress(KeyAscii As Integer)

 On Error GoTo err_handler

 Dim fin_selrow%, fin_selcol%
 fin_selrow = att_flex.Row
 fin_selcol = att_flex.Col
 With att_flex
''    If fin_selcol = 21 Then
''        If att_flex.TextMatrix(fin_selrow, fin_selcol) = "Y" Then
''             att_flex.TextMatrix(fin_selrow, fin_selcol) = "N"
''        Else
''             att_flex.TextMatrix(fin_selrow, fin_selcol) = "Y"
''        End If
''    End If
    If fin_selcol = 22 Then

           If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                'allow backspace and the enter keys
                    MsgBox "Enter OT hours in numbers only "
                    If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
                    .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                    End If
                    att_flex.SetFocus
                Else
                
                    att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
                    If Val(att_flex.TextMatrix(fin_selrow, fin_selcol)) > Val(att_flex.TextMatrix(fin_selrow, 20)) Then
                       MsgBox "OT WILL NOT BE GREATER THAN ACTUAL"
                       .TextMatrix(fin_selrow, fin_selcol) = .TextMatrix(fin_selrow, 20)
''                       .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                       att_flex.SetFocus
                    End If
                End If
           End If
           If KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
              .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              End If
           End If
        End If
        
    If fin_selcol = 23 Then

           If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                'allow backspace and the enter keys
                    MsgBox "Enter OT hours in numbers only "
                    If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
                    .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                    End If
                    att_flex.SetFocus
                Else
                
                    att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
                    If Val(att_flex.TextMatrix(fin_selrow, fin_selcol)) > 20 Then
                       MsgBox "OT WILL NOT BE GREATER THAN ACTUAL"
                       .TextMatrix(fin_selrow, fin_selcol) = .TextMatrix(fin_selrow, 20)
                       att_flex.SetFocus
                    End If
                End If
           End If
           If KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
              .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              End If
           End If
        End If
        
        
    If fin_selcol = 19 Then

           If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                'allow backspace and the enter keys
                    MsgBox "Enter OT hours in numbers only "
                    If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
                    .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
                    End If
                    att_flex.SetFocus
                Else
                
''                    att_flex.TextMatrix(fin_selrow, fin_selcol) = att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii)
''                    If Val(att_flex.TextMatrix(fin_selrow, fin_selcol)) > 20 Then
''                       MsgBox "OT WILL NOT BE GREATER THAN ACTUAL"
''                       .TextMatrix(fin_selrow, fin_selcol) = .TextMatrix(fin_selrow, 20)
''                       att_flex.SetFocus
''                    End If
                End If
           End If
           If KeyAscii = 8 Then
              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
              .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
              End If
           End If
        End If
        
''         Select Case att_flex.Col
''        Case 7
''            If att_flex.TextMatrix(fin_selrow, fin_selcol) = "A" Then
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "P"
''            ElseIf att_flex.TextMatrix(fin_selrow, fin_selcol) = "P" Then
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "1/2P"
''            Else
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "A"
''            End If
''
''
''    End Select
''    ElseIf fin_selcol = 6 Then
''            If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 Then
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = UCase(att_flex.TextMatrix(fin_selrow, fin_selcol) & Chr(KeyAscii))
''                Dim Temp As String
''                Dim txt As String
''                Dim I As Integer
''                Dim C As String
''
''                txt = UCase$(att_flex.TextMatrix(fin_selrow, fin_selcol))
''                Temp = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 'You may want to allow a space char too
''                For I = 1 To Len(txt)
''                    C = Mid$(txt, I, 1)
''                    If InStr(Temp, C) = 0 Then
''                    MsgBox "Invalid Shift Press A/B/C/G "
''                    .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
''                    End If
''                    att_flex.SetFocus
''                    Exit Sub
''                Next
''
''           End If
''           If KeyAscii = 8 Then
''              If Len(.TextMatrix(fin_selrow, fin_selcol)) > 0 Then
''              .TextMatrix(fin_selrow, fin_selcol) = Left$((.TextMatrix(fin_selrow, fin_selcol)), Len(.TextMatrix(fin_selrow, fin_selcol)) - 1)
''              End If
''           End If
''    End If
''
End With
Exit Sub
err_handler:
        chk = gen_Validation(Err.Number, Err.Description)
        If chk = 1 Then
            Resume
        End If

End Sub
Private Sub att_flex_EnterCell()
''On Error GoTo err_handler
''
''Dim fin_selrow%, fin_selcol%
'' fin_selrow = att_flex.Row
'' fin_selcol = att_flex.Col
'' With att_flex
''
''    Select Case att_flex.Col
''        Case 7
''            If att_flex.TextMatrix(fin_selrow, fin_selcol) = "A" Then
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "P"
''            ElseIf att_flex.TextMatrix(fin_selrow, fin_selcol) = "P" Then
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "1/2P"
''            Else
''                att_flex.TextMatrix(fin_selrow, fin_selcol) = "A"
''            End If
''
''
''    End Select
''    End With
''    Exit Sub
''err_handler:
''        chk = gen_Validation(Err.Number, Err.Description)
''        If chk = 1 Then
''            Resume
''        End If

End Sub





Private Sub NEW_Click()
  Refresh_Click
  monthcheck
  
End Sub

Function filldata()

If dt_ot >= CDate("11/18/2025") Then
  overtime_eligible = 1
Else
  overtime_eligible = 2
End If

Dim dayfind, dayfind_intime, dayfind_outtime As String
Dim intime, outtime, difftime, in1, in2, out1, out2 As Integer

Dim exhrs, cmin1, cmin2, tmins, thrs, eligible_mins, eligible_hrs, balance_mins, balance_hrs As Double

Dim hrs_to_work, mins_to_work, bal_hrs As Double

''Set paydb = New ADODB.Connection
Set payrs = New ADODB.Recordset
Dim millcode As Integer

If cmb_millname.Text = "SHVPM" Then
    millcode = 1
End If

millcode = 1

''    If opt_all.Value = True Then
''        If opt_all_dept.Value = True Then
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  order by dept_name,emp_name "
''        Else
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and emp_dept = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  order by dept_name,emp_name "
''
''        End If
''     ElseIf opt_worker.Value = True Then
''        If opt_all_dept.Value = True Then
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'W'  order by dept_name,emp_name "
''        Else
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'W' and emp_dept = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  order by dept_name,emp_name "
''
''        End If
''
''     ElseIf opt_cs.Value = True Then
''        If opt_all_dept.Value = True Then
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'C'  order by dept_name,emp_name "
''        Else
''            sql = " select * from bio_device_shiftlogs a, " _
''                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat  ,0 as hrs_to_work from emp_voupay_mast where emp_company = " & millcode & " and emp_cat = 'W'  and emp_status = 'A' " _
''                  & " Union All " _
''                  & " select ca_fpcode as emp_fpcode,ca_compcode as emp_company,ca_empname as  emp_name,ca_dept as emp_dept,'C' as emp_cat ,ca_work_hrs as hrs_to_work   from mas_caemp where ca_dept not in (28,36) and ca_compcode = " & millcode & " and ca_pi_yn ='Y' and ca_status = 'A' " _
''                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'C' and emp_dept = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  order by dept_name,emp_name "
''
''        End If
''
''     End If
     

        If opt_all_dept.Value = True Then
            sql = " select * from bio_device_shiftlogs a, " _
                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat, emp_pi_eligible_yn,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'W'  order by dept_name,emp_fpcode "
        Else
            If cmb_dept.ListIndex = -1 Then
               MsgBox ("Select Department...")
               Exit Function
            End If
            sql = " select * from bio_device_shiftlogs a, " _
                  & " (select emp_fpcode,emp_company,emp_name,emp_dept,emp_cat, emp_pi_eligible_yn ,emp_work_hrs as hrs_to_work from emp_mas   where emp_company = " & millcode & " and emp_cat = 'W' and emp_pi_eligible_yn = 'Y' and emp_status = 'A' " _
                  & " ) b , pdept_mas c where ds_fpcode = emp_fpcode and emp_dept = dept_code  and ds_date = '" & Format(dt_ot, "yyyy/MM/dd") & "' and (ds_sft_hrs+ds_sft_hrs2+ds_sft_hrs3 > 9.59 or (ds_shift = 'WO' and ds_sft_hrs >0) or hrs_to_work > 0)  and emp_cat = 'W' and emp_dept = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  order by dept_name,emp_fpcode "

        End If



Dim i As Integer
Dim odmins, permins As Double

i = 1
paydb.CommandTimeout = 300
payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
If Not payrs.EOF Then
   While Not payrs.EOF
             
      '' if ds_fpcode =
      
        With att_flex
             whrs = 0
             thrs = 0
             cmin1 = 0
             cmin2 = 0
             cmin3 = 0
             tmins = 0
             odmins = 0
             permins = 0
             eligible_mins = 0
             eligible_hrs = 0
             bal_mins = 0
             bal_hrs = 0
             odmins = (Int(payrs("ds_od_hrs")) * 60) + (payrs("ds_od_hrs") - Int(payrs("ds_od_hrs"))) * 100
             permins = (Int(payrs("ds_per_hrs")) * 60) + (payrs("ds_per_hrs") - Int(payrs("ds_per_hrs"))) * 100
''If payrs("emp_fpcode") = 1049 Then
''   MsgBox ("Wait")
''End If

             
             hrs_to_work = 0
             If payrs("ds_sft_hrs1") > 0 Then
                cmin1 = (Int(payrs("ds_sft_hrs1")) * 60) + (payrs("ds_sft_hrs1") - Int(payrs("ds_sft_hrs1"))) * 100
                cmin2 = (Int(payrs("ds_sft_hrs2")) * 60) + (payrs("ds_sft_hrs2") - Int(payrs("ds_sft_hrs2"))) * 100
             End If
             If payrs("ds_sft_hrs3") > 0 Then
                cmin3 = (Int(payrs("ds_sft_hrs3")) * 60) + (payrs("ds_sft_hrs3") - Int(payrs("ds_sft_hrs3"))) * 100
             End If
              
''             If cmin2 > 0 Then
''                cmin1 = cmin1 + 30
''             End If
             
''             If cmin1 + cmin2 + cmin3 + odmins > 0 Then
''                tmins = cmin1 + cmin2 + cmin3 + odmins + permins
''                thrs = Round(Int(tmins / 60) + ((tmins - Int(tmins / 60) * 60) / 100), 2)
''             End If
             
''             If payrs("ds_shift_actual") = "GS" Then
''                mins_to_work = tmins - ((payrs("hrs_to_work") + 1) * 60)
''             Else
''                mins_to_work = tmins - (payrs("hrs_to_work") * 60)
''             End If
             
             If payrs("ds_shift_actual") = "GS" Then

                If cmin1 + cmin2 + cmin3 + odmins > 0 Then
                   If cmin1 + cmin2 + cmin3 > 540 Then
                       tmins = cmin1 + cmin2 + cmin3 + permins
                       eligible_mins = cmin1 + cmin2 + cmin3 + permins - 60
                   Else
                       tmins = cmin1 + cmin2 + cmin3 + odmins + permins
                       eligible_mins = cmin1 + cmin2 + cmin3 + odmins + permins - 60
                   End If
                   thrs = Round(Int(tmins / 60) + ((tmins - Int(tmins / 60) * 60) / 100), 2)
                  
                   eligible_hrs = Round(Int(eligible_mins / 60) + ((eligible_mins - Int(eligible_mins / 60) * 60) / 100), 2)
                   mins_to_work = (9 * 60)
                End If
                
                
             Else
                
                If cmin1 + cmin2 + cmin3 + odmins > 0 Then
                   If cmin1 + cmin2 + cmin3 > 480 Then
                       tmins = cmin1 + cmin2 + cmin3 + permins
                       eligible_mins = cmin1 + cmin2 + cmin3 + permins
                   Else
                       tmins = cmin1 + cmin2 + cmin3 + odmins + permins
                       eligible_mins = cmin1 + cmin2 + cmin3 + odmins + permins
                   End If
                   thrs = Round(Int(tmins / 60) + ((tmins - Int(tmins / 60) * 60) / 100), 2)
                   
                   eligible_hrs = Round(Int(eligible_mins / 60) + ((eligible_mins - Int(eligible_mins / 60) * 60) / 100), 2)
                   mins_to_work = (8 * 60)
                End If
                
             End If
             
             balance_mins = eligible_mins - 480 '' 480 = 8 * 60
             If balance_mins < 0 Then balance_mins = 0
             
             
             If mins_to_work < 0 Then mins_to_work = 0
             
             bal_hrs = Round(Int(balance_mins / 60) + ((balance_mins - Int(balance_mins / 60) * 60) / 100), 2)
            
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = payrs("dept_name")
                    .TextMatrix(.Rows - 1, 1) = payrs("dept_code")
                    .TextMatrix(.Rows - 1, 2) = payrs("emp_name")
                    .TextMatrix(.Rows - 1, 3) = payrs("emp_fpcode")
                    .TextMatrix(.Rows - 1, 4) = payrs("ds_shift")
                    .TextMatrix(.Rows - 1, 21) = payrs("emp_cat")
                    .TextMatrix(.Rows - 1, 5) = Format(payrs("ds_shift_actual"), "#0.00")
''                    .TextMatrix(.Rows - 1, 6) = Format(payrs("hrs_to_work"), "#0.00")
                    
                         If payrs("ds_shift_actual") = "GS" Then
                            .TextMatrix(.Rows - 1, 6) = Format(9, "#0.00")
                         Else
                            .TextMatrix(.Rows - 1, 6) = Format(8, "#0.00")
                         End If
                         
                         

                    If IsNull(payrs("ds_shift_in")) Or payrs("ds_shift_in") = "01/01/1900" Then
                       .TextMatrix(.Rows - 1, 7) = ""
                       .TextMatrix(.Rows - 1, 8) = ""
                    Else
                       .TextMatrix(.Rows - 1, 7) = payrs("ds_shift_in")
                       .TextMatrix(.Rows - 1, 8) = payrs("ds_shift_out")
                        .TextMatrix(.Rows - 1, 13) = payrs("ds_sft_hrs1")
                    End If
                     
                    If IsNull(payrs("ds_shift_in2")) Or payrs("ds_shift_in2") = "01/01/1900" Then
                        .TextMatrix(.Rows - 1, 9) = ""
                        .TextMatrix(.Rows - 1, 10) = ""
                    Else
                        .TextMatrix(.Rows - 1, 9) = payrs("ds_shift_in2")
                        .TextMatrix(.Rows - 1, 10) = payrs("ds_shift_out2")
                        .TextMatrix(.Rows - 1, 14) = payrs("ds_sft_hrs2")
                    End If
                    
                                             
                         If IsNull(payrs("ds_shift_in3")) Or payrs("ds_shift_in3") = "01/01/1900" Then
                            .TextMatrix(.Rows - 1, 24) = ""
                            .TextMatrix(.Rows - 1, 25) = ""
                            .TextMatrix(.Rows - 1, 26) = ""
                         Else
                            .TextMatrix(.Rows - 1, 24) = payrs("ds_shift_in3")
                            .TextMatrix(.Rows - 1, 25) = payrs("ds_shift_out3")
                            .TextMatrix(.Rows - 1, 26) = payrs("ds_sft_hrs3")
                         End If
                         .TextMatrix(.Rows - 1, 28) = payrs("ds_per_hrs")
                         .TextMatrix(.Rows - 1, 29) = payrs("ds_od_hrs")
                    If eligible_hrs > 0 Then .TextMatrix(.Rows - 1, 15) = Format(eligible_hrs, "#0.00")
                    
                    
''                    If Left(payrs("ds_shift"), 2) = "WO" Or Left(payrs("ds_shift"), 1) = "H" Then
''                         If thrs > 0 Then .TextMatrix(.Rows - 1, 18) = Format(thrs, "#0.00")
''                    Else
''                        If thrs > 0 Then .TextMatrix(.Rows - 1, 18) = Format(thrs - 8, "#0.00")
''''                        If bal_hrs > 0 Then .TextMatrix(.Rows - 1, 18) = Format(bal_hrs, "#0.00")
''                    End If
                      If bal_hrs > 0 Then .TextMatrix(.Rows - 1, 18) = Format(bal_hrs, "#0.00")
                      If bal_hrs > 0 Then .TextMatrix(.Rows - 1, 22) = Format(bal_hrs, "#0.00")
                                             
                     
''                     MsgBox (.TextMatrix(.Rows - 1, 18))
                    
                    
                    i = i + 1
        End With
        
        payrs.MoveNext
    Wend
    
    
    
    
    payrs.Close
    
    

    
    
    sql = "select * from bio_emp_chleave where empch_worked_date = '" & Format(dt_ot, "yyyy/MM/dd") & "'"
    
    payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
       For i = 1 To att_flex.Rows - 1
       

          thrs = Val(att_flex.TextMatrix(i, 18))
          If att_flex.TextMatrix(i, 3) = payrs("empch_fpcode") Then
             att_flex.TextMatrix(i, 16) = payrs("empch_ch_date")
             If payrs("emp_ch_period") = "H" Then
                att_flex.TextMatrix(i, 18) = 4
                att_flex.TextMatrix(i, 19) = Format(Round(thrs - 4, 0), "#0.00")
''                att_flex.TextMatrix(i, 17) = Format(Round(thrs - 12, 0), "#0.00")
             Else
                att_flex.TextMatrix(i, 18) = 8
                att_flex.TextMatrix(i, 19) = Format(Round(thrs - 8, 0), "#0.00")
''                att_flex.TextMatrix(i, 17) = Format(Round(thrs - 16, 0), "#0.00")
             End If
          End If
       Next
    
    
    
       payrs.MoveNext
    Wend
    
 Else
    MsgBox ("Details not available for the month")
 End If
 
    
For i = 1 To att_flex.Rows - 1
''
''If Val(att_flex.TextMatrix(i, 3)) = 2079 Then
''   MsgBox ("Wait")
''End If


''   If Trim(att_flex.TextMatrix(i, 5)) = "WO" Then
''
''      If Val(att_flex.TextMatrix(i, 18)) >= 8 Then
''         att_flex.TextMatrix(i, 19) = "8.00"
''      Else
''         att_flex.TextMatrix(i, 19) = att_flex.TextMatrix(i, 18)
''      End If
''   End If
   
   If Val(att_flex.TextMatrix(i, 18)) - Val(att_flex.TextMatrix(i, 19)) >= overtime_eligible Then
      att_flex.TextMatrix(i, 20) = Format(Round(Val(att_flex.TextMatrix(i, 18)) - Val(att_flex.TextMatrix(i, 19)), 0), "#0.00")
   Else
      att_flex.TextMatrix(i, 20) = 0
   End If
   
   If Val(att_flex.TextMatrix(i, 18)) >= overtime_eligible Then
      att_flex.TextMatrix(i, 18) = Format(Round(Val(att_flex.TextMatrix(i, 18)), 0), "#0.00")
   Else
      att_flex.TextMatrix(i, 18) = 0
   End If
   If Val(att_flex.TextMatrix(i, 19)) > 0 Then
      att_flex.TextMatrix(i, 19) = Format(Round(Val(att_flex.TextMatrix(i, 19)), 0), "#0.00")
   Else
      att_flex.TextMatrix(i, 19) = 0
   End If
   If Val(att_flex.TextMatrix(i, 20)) >= overtime_eligible Then
      att_flex.TextMatrix(i, 20) = Format(Round(Val(att_flex.TextMatrix(i, 20)), 0), "#0.00")
   Else
      att_flex.TextMatrix(i, 20) = 0
   End If
   If Val(att_flex.TextMatrix(i, 22)) >= overtime_eligible Then

      att_flex.TextMatrix(i, 22) = Format(Round(Val(att_flex.TextMatrix(i, 22)), 0), "#0.00")
   Else
      att_flex.TextMatrix(i, 22) = 0
   End If
Next
 
 payrs.Close
 '' paydb.close


End Function

Function call_ot_days()

    If getotday = 1 Then
        Dim payrs2 As New ADODB.Recordset
        sql = "select * from bio_attendlogs where a_year = " & myear & " and a_month = " & mmon
        payrs2.Open sql, paydb, adOpenDynamic, adLockOptimistic
        While Not payrs2.EOF
           For i = 1 To att_flex.Rows - 1
              If att_flex.TextMatrix(i, 3) = payrs2("a_fpcode") Then
                 If payrs2("a_ot_days") > 0 Then att_flex.TextMatrix(i, 27) = payrs2("a_ot_days") * 8
              End If
           Next
           payrs2.MoveNext
        Wend
        payrs2.Close
   End If


End Function

Private Sub opt_cs_Click()
    emptype = "C"
End Sub

Private Sub opt_worker_Click()
    emptype = "W"
End Sub

Private Sub Refresh_Click()
fst_save = 0

opt = 0

''If cmb_millname.Text = "" Then
''   MsgBox "Select Mill "
''   cmb_millname.SetFocus
''   Exit Sub
''End If
''


''Set paydb = New ADODB.Connection
Set payrs = New ADODB.Recordset
''paydb.Open pay
Dim millcode As Integer
If cmb_millname.Text = "SHVPM" Then
    millcode = 1
End If
    millcode = 1
If opt_selective_dept.Value = True And cmb_dept.Text = "" Then
   MsgBox ("Select Department...")
   fillgrid
   Exit Sub
End If

''If opt_all.Value = True Then
''   sql = "select *  from bio_worker_daily_pihrs where w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
''ElseIf opt_worker.Value = True Then
''   sql = "select *  from bio_worker_daily_pihrs where w_cat = 'W' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
''ElseIf opt_cs.Value = True Then
''   sql = "select *  from bio_worker_daily_pihrs where w_cat = 'C' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
''
''End If
If opt_selective_dept.Value = True And cmb_dept.Text <> "" Then
   sql = "select *  from bio_worker_daily_pihrs where w_deptcode = " & cmb_dept.ItemData(cmb_dept.ListIndex) & "  and  w_cat = 'W' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
Else
   sql = "select *  from bio_worker_daily_pihrs where  w_cat = 'W' and w_company = " & millcode & " and  w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "'"
End If
payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
If Not payrs.EOF Then
 MsgBox "Data already available,Please EDIT"
 Exit Sub
End If
payrs.Close
'' paydb.close

fillgrid
filldata

call_ot_days

End Sub

Private Sub save_Click()

datefrom = dt_ot.Value - Day(dt_ot.Value) + 1
dateto = dt_ot.Value
If Day(dt_ot.Value) <> 1 Then dateto = dt_ot.Value - 1

Dim qry, cat As String
Dim maxhrs, tot_ot_hrs As Double
On Error GoTo err_handler
   If att_flex.Rows < 2 Then
      MsgBox (" Details not available ")
      Exit Sub
   End If

'''
'''  For i = 1 To att_flex.Rows - 1
'''      If Val(att_flex.TextMatrix(i, 19)) + Val(att_flex.TextMatrix(i, 22)) > 0 Then
'''         If att_flex.TextMatrix(i, 21) = "W" Then
'''            cat = "W"
'''''            If att_flex.TextMatrix(i, 3) = "4024" Then
'''''               MsgBox ("WAIT")
'''''            End If
'''            qry = "select * from emp_mas where emp_fpcode = " & Val(att_flex.TextMatrix(i, 3)) & " and emp_pi_max_yn = 'Y'"
'''            maxhrs = 0
'''            payrs.Open qry, paydb, adOpenDynamic, adLockOptimistic
'''            If Not payrs.EOF Then
'''               maxhrs = payrs("emp_pi_max_hrs")
'''            End If
'''            payrs.Close
'''         Else
'''           cat = "C"
'''''           If att_flex.TextMatrix(i, 3) = "4517" Then
'''''               MsgBox ("WAIT")
'''''            End If
'''            qry = "select * from mas_caemp where ca_fpcode = " & Val(att_flex.TextMatrix(i, 3)) & " and ca_pi_max_yn = 'Y'"
'''            maxhrs = 0
'''            payrs.Open qry, paydb, adOpenDynamic, adLockOptimistic
'''            If Not payrs.EOF Then
'''               maxhrs = payrs("ca_pi_max_hrs")
'''            End If
'''            payrs.Close
'''         End If
'''            If maxhrs > 0 Then
'''                qry = "select sum(w_accepted_hrs) as accept_hrs from bio_worker_daily_pihrs where w_emp_fpcode = " & att_flex.TextMatrix(i, 3) & "  and w_cat = '" & cat & "' and w_date between  '" & Format(datefrom, "MM/dd/yyyy") & "'  and '" & Format(dateto, "MM/dd/yyyy") & "'"
'''
'''                payrs.Open qry, paydb, adOpenDynamic, adLockOptimistic
'''                If Not payrs.EOF Then
'''                   tot_ot_hrs = IIf(IsNull(payrs("accept_hrs")), 0, payrs("accept_hrs"))
'''                End If
'''                payrs.Close
'''                If Val(att_flex.TextMatrix(i, 18)) + tot_ot_hrs > maxhrs Then
'''                   MsgBox ("Upto date OT Hours  = (" + Str(Val(att_flex.TextMatrix(i, 18)) + tot_ot_hrs) + ") is Exceed then allowed of " + Str(maxhrs) + " Hours for " + att_flex.TextMatrix(i, 2))
'''                   Exit Sub
'''                End If
'''            End If
'''      End If
'''  Next





  Me.MousePointer = 11
  Set payrs = New ADODB.Recordset
If cmb_millname.Text = "SHVPM" Then
    millcode = 1
End If
    millcode = 1
For i = 1 To att_flex.Rows - 1
    sql = "delete from bio_worker_daily_pihrs where w_company = " & millcode & " and w_date = '" & Format(dt_ot.Value, "MM/dd/yyyy") & "' and w_emp_fpcode = " & Val(att_flex.TextMatrix(i, 3))
    paydb.Execute sql
Next
  sql = "select * from bio_worker_daily_pihrs where 1=2"
  payrs.Open sql, paydb, adOpenDynamic, adLockOptimistic
  For i = 1 To att_flex.Rows - 1
      If Val(att_flex.TextMatrix(i, 19)) + Val(att_flex.TextMatrix(i, 22)) + Val(att_flex.TextMatrix(i, 23)) + Val(att_flex.TextMatrix(i, 27)) > 0 Then
            payrs.AddNew
            payrs.Fields("w_company") = millcode
            payrs.Fields("w_date") = Format(dt_ot.Value, "dd/MM/yyyy")
            payrs.Fields("w_cat") = att_flex.TextMatrix(i, 21)
            payrs.Fields("w_emp_fpcode") = att_flex.TextMatrix(i, 3)
            payrs.Fields("w_deptcode") = Val(att_flex.TextMatrix(i, 1))
            payrs.Fields("w_tot_ot_hrs") = Val(att_flex.TextMatrix(i, 18))
            
''            payrs.Fields("w_wo_ot_hrs") = Val(att_flex.TextMatrix(i, 19))
            
            payrs.Fields("w_wo_ot_hrs") = 0
            payrs.Fields("w_act_hrs") = Val(att_flex.TextMatrix(i, 20))
            payrs.Fields("w_accepted_hrs") = Val(att_flex.TextMatrix(i, 22))
            payrs.Fields("w_holiday_ot_hrs") = Val(att_flex.TextMatrix(i, 23))
            payrs.Fields("w_woot_days_hrs") = Val(att_flex.TextMatrix(i, 27))
            
            
            payrs.Update
      End If
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


