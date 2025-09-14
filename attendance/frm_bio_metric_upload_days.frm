VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_bio_metric_upload_days 
   Caption         =   "ATTENDANCE UPLOAD FROM BIO-METRIC SYSTEM - forthe days"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE UNREGISTED  EMPOYEE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   11880
      TabIndex        =   23
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "PROCESS TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2280
      TabIndex        =   20
      Top             =   2040
      Width           =   8775
      Begin VB.OptionButton optFinal 
         Caption         =   "Final Process"
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
         Left            =   960
         TabIndex        =   22
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton optRegular 
         Caption         =   "Regular Process"
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
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATA IMPORT FROM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2400
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   8775
      Begin VB.OptionButton opt_local 
         Caption         =   "Local D Drive"
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
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opt_server 
         Caption         =   "Server(10.0.0.252)"
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
         Left            =   4440
         TabIndex        =   18
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9720
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd_process 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&IMPORT - LOGS from biometric"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   10680
      MaskColor       =   &H000000FF&
      Picture         =   "frm_bio_metric_upload_days.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Frame frame_month 
      Height          =   1095
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   8295
      Begin VB.ComboBox cmb_year 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmb_month 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
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
         Height          =   330
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "YEAR"
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
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5400
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
      Begin VB.CommandButton cmd_upload 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Process"
         Height          =   705
         Left            =   360
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_days.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         Height          =   705
         Left            =   1440
         MaskColor       =   &H000000FF&
         Picture         =   "frm_bio_metric_upload_days.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   4320
      TabIndex        =   1
      Top             =   4920
      Width           =   3735
      Begin MSComCtl2.DTPicker st_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130351105
         CurrentDate     =   44565
      End
      Begin MSComCtl2.DTPicker end_date 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130351105
         CurrentDate     =   44565
      End
      Begin VB.Label Label10 
         Caption         =   "Report From Date"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "End Date"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker monthend_date 
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130351105
      CurrentDate     =   44565
   End
   Begin VB.Label lbl_emp 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE ATTENDANCE UPLOAD FROM BIO-METRIC SYSTEMS"
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
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "frm_bio_metric_upload_days"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProcessType As String
Dim mdays As Integer

Private Sub cmb_month_Change()
   find_dates
End Sub

Private Sub cmb_month_Click()
    find_dates
End Sub

Private Sub cmb_year_Change()
find_dates
End Sub

Private Sub cmb_year_Click()
find_dates
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_process_Click()
   pst_ans = MsgBox("Are u sure want to Update Attendance Logs from Bio-metric machines ...  ", vbYesNo)
   If pst_ans = vbNo Then Exit Sub
      
   all_chk = 1
   upload_biometric_data
   
   MsgBox ("Data uploaded from biometic - Completed ..")
End Sub

Private Sub cmd_upload_Click()
       
''   all_chk = 1
''   update_leave_od_data
''   cmd_exit.SetFocus
''
''
   pst_qry = "update bio_devicelogs set ad_upd ='N' where ad_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date + 1, "MM/dd/yyyy") & "'"
   paydb.Execute pst_qry
   
   all_chk = 1
   update_leave_od_data
   cmd_exit.SetFocus
   
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim absent As Single
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
''            If payrs.Fields("a_fpcode") = 3221 Then
''               MsgBox ("Wait")
''            End If
            absent = 0
            For i = 1 To 31
                dayfind = "a_day" & i
                If payrs.Fields(dayfind) = "A" Then
                    absent = absent + 1
                Else
''                    If absent >= 4 And payrs.Fields("a_present") > 0 Then
''''                       MsgBox (payrs.Fields("a_fpcode"))
''                       If payrs.Fields(dayfind) = "WO" Then
''                           payrs.Fields(dayfind) = "A"
''                           payrs.Update
''                       End If
''                    End If
                    If absent >= 4 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = "A"
                           payrs.Update
                    End If
                    absent = 0
                End If
            Next
            payrs.MoveNext
        Wend
    End If
    payrs.Close
    
    Dim present, hop, wop, cl, sl, h, ch, layoff, wo, pl, hope, el, woh, ml, totdays, eligible_days As Single
    sql = "select * from bio_attendlogs ,emp_mas  where a_fpcode = emp_fpcode and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
       
''             If payrs.Fields("a_fpcode") = 3221 Then
''               MsgBox ("Wait")
''            End If
            
            present = 0
            absent = 0
            hop = 0
            wop = 0
            cl = 0
            sl = 0
            h = 0
            ch = 0
            layoff = 0
            wo = 0
            pl = 0
            hope = 0
            el = 0
            woh = 0
            ml = 0
            For i = 1 To 31
                dayfind = "a_day" & i
                If payrs.Fields(dayfind) = "P" Or payrs.Fields(dayfind) = "OD" Or payrs.Fields(dayfind) = "P(OD)" Or payrs.Fields(dayfind) = "½P(OD)" Or payrs.Fields(dayfind) = "A(OD)" Then
                    present = present + 1
                ElseIf payrs.Fields(dayfind) = "A" Then
                    absent = absent + 1
                ElseIf payrs.Fields(dayfind) = "L" Or payrs.Fields(dayfind) = "PLP" Then
                    pl = pl + 1
                ElseIf payrs.Fields(dayfind) = "½L" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½LP" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "EL" Or payrs.Fields(dayfind) = "CL" Or payrs.Fields(dayfind) = "CL½P" Or payrs.Fields(dayfind) = "CLP" Then
                    el = el + 1
                ElseIf payrs.Fields(dayfind) = "½CL" Then
                    absent = absent + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½CLP" Or payrs.Fields(dayfind) = "½CL½P" Or payrs.Fields(dayfind) = "½EL" Then
                    present = present + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.HP" Then
                    present = present + 0.5
                    ch = ch + 0.5
                ElseIf payrs.Fields(dayfind) = "SL" Or payrs.Fields(dayfind) = "SLP" Then
                    sl = sl + 1
                ElseIf payrs.Fields(dayfind) = "½SLP" Then
                    sl = sl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "H" Then
                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "HP" Or payrs.Fields(dayfind) = "HP(OD)" Then
                    hop = hop + 1
                    h = h + 1

                ElseIf payrs.Fields(dayfind) = "HPE" Then
                    hope = hope + 1
                    h = h + 1
                    hop = hop + 1
                
                ElseIf payrs.Fields(dayfind) = "½HP" Then
''                    hop = hop + 0.5
                    hop = hop + 0.5
                    h = h + 1
                
                ElseIf payrs.Fields(dayfind) = "½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "Layoff" Or payrs.Fields(dayfind) = "LayoffP" Or payrs.Fields(dayfind) = "LAYOFF" Or payrs.Fields(dayfind) = "½PLAYOFF" Then
                    layoff = layoff + 1
                ElseIf payrs.Fields(dayfind) = "C.H" Or payrs.Fields(dayfind) = "C.H½P" Or payrs.Fields(dayfind) = "C.HP" Or payrs.Fields(dayfind) = "C.HP(OD)" Then
                    ch = ch + 1
                ElseIf payrs.Fields(dayfind) = "HOP" Or payrs.Fields(dayfind) = "H½P(OD)" Then
                    hop = hop + 1
                ElseIf payrs.Fields(dayfind) = "WOP" Or payrs.Fields(dayfind) = "WOP(OD)" Or payrs.Fields(dayfind) = "WO(OD)" Then
                    wop = wop + 1
                ElseIf payrs.Fields(dayfind) = "WO" Or payrs.Fields(dayfind) = "WOL" Then
                    wo = wo + 1
                ElseIf payrs.Fields(dayfind) = "WO½P" Then
                    wop = wop + 0.5
                    wo = wo + 0.5
                ElseIf payrs.Fields(dayfind) = "WOH" Then
                    woh = woh + 1
                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "WOHP" Then
                    woh = woh + 1
''                    h = h + 1
                    wop = wop + 1
                ElseIf payrs.Fields(dayfind) = "WOHPE" Then
                    woh = woh + 1
                    h = h + 1
                    wop = wop + 1
''                    hope = hope + 1


''                ElseIf payrs.Fields(dayfind) = "½C.H" Then
''                    ch = ch + 0.5
''                    present = present + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½C.H" Then
                    ch = ch + 0.5
                    el = el + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½L" Then
                    pl = pl + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "P½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PC.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PC.H½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½EL" Then
                    el = el + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½L" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½C.H" Then
                    pl = pl + 0.5
                    ch = ch + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½A" Or payrs.Fields(dayfind) = "½A½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½A" Or payrs.Fields(dayfind) = "½A½EL" Then
                    el = el + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½A" Or payrs.Fields(dayfind) = "½A½L" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.H½A" Or payrs.Fields(dayfind) = "½A½C.H" Then
                    ch = ch + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.H½P" Then
                    present = present + 0.5
                    ch = ch + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½A½OD" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½OD" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½OD" Then
                    el = el + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½ML" Or payrs.Fields(dayfind) = "½P½ML" Or payrs.Fields(dayfind) = "½ML½P" Then
                    ml = ml + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "ML" Then
                    ml = ml + 1
                End If
            Next
            
            payrs("a_present") = present
            payrs("a_hop") = hop
            payrs("a_hpe") = hope
            payrs("a_wop") = wop
            payrs("a_el") = el
            payrs("a_pl") = pl
            payrs("a_ml") = ml
            payrs("a_holiday") = h
            payrs("a_ch") = ch
            payrs("a_layoff") = layoff
            payrs("a_absent") = absent
            payrs("a_wo") = wo + woh
            payrs("a_woh") = woh
            
            payrs("a_month_days") = mdays
            
            If payrs("emp_cat") = "W" Then
''               totdays = present + wop + wo + wop + ch + ml
               totdays = present + wop + wo + wop + ch + ml + hop
            Else
               totdays = present + wop + wo + ch + ml
            End If
            
            
            payrs("a_month_days") = mdays
            
''            totdays = present + wop + wo + wop + ch
            
            payrs("a_total_days") = totdays
            
            
            
            If totdays >= mdays Then
                payrs("a_eligible_days") = mdays
                payrs("a_salary_days") = mdays + h
                If payrs("emp_cat") = "W" Then
                   payrs("a_ot_days") = totdays - mdays
                Else
                   payrs("a_ot_days") = 0
                End If
           Else
            
                payrs("a_eligible_days") = totdays
                payrs("a_salary_days") = totdays + h
                payrs("a_ot_days") = 0
            End If
            
            payrs.Update
            payrs.MoveNext
        Wend
    End If
    payrs.Close


 '' For reversing Holiday present if avail CH
    Dim holidaych As Single
    sql = "select * ,case when emp_ch_period = 'F' then 1 else 0.5 end as Holidaych from bio_device_shiftlogs a, bio_emp_chleave b where  ds_fpcode =  empch_fpcode and ds_shift = 'H' and ds_date = empch_worked_date and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' order by ds_date"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          idate = payrs!ds_date
          id = payrs!ds_fpcode
          holidaych = payrs!holidaych
          pst_qry = "update bio_attendlogs set a_hop  = a_hop - " & holidaych & "  where a_fpcode = " & id & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " "
          paydb.Execute pst_qry
          payrs.MoveNext
    Wend
    payrs.Close
    pst_qry = "update bio_attendlogs set a_hop  = 0 where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_hop < 0 "
    paydb.Execute pst_qry
    
    MsgBox ("updated.")
End Sub

Private Sub Command2_Click()
                For idate = st_date To end_date
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (10031, 10031,  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
                   paydb.Execute pst_qry
                   
                Next
                
End Sub

Private Sub Form_Load()
    ProcessType = "R"
    st_date.Value = Now
    end_date.Value = Now
    monthend_date.Value = Now
        all_chk = 0
    With cmb_month
        .AddItem "January"
        .ItemData(.NewIndex) = 1
        .AddItem "February"
        .ItemData(.NewIndex) = 2
        .AddItem "March"
        .ItemData(.NewIndex) = 3
        .AddItem "April"
        .ItemData(.NewIndex) = 4
        .AddItem "May"
        .ItemData(.NewIndex) = 5
        .AddItem "June"
        .ItemData(.NewIndex) = 6
        .AddItem "July"
        .ItemData(.NewIndex) = 7
        .AddItem "August"
        .ItemData(.NewIndex) = 8
        .AddItem "September"
        .ItemData(.NewIndex) = 9
        .AddItem "October"
        .ItemData(.NewIndex) = 10
        .AddItem "November"
        .ItemData(.NewIndex) = 11
        .AddItem "December"
        .ItemData(.NewIndex) = 12
    End With
    
    With cmb_year
      .AddItem Left(fyear, 4)
      .AddItem Mid(fyear, 6, 4)
      If Year(Date) = Int(Left(fyear, 4)) Then
         cmb_year.Text = Left(fyear, 4)
      Else
          cmb_year.Text = Mid(fyear, 6, 4)
      End If
    End With
    cmb_month.ListIndex = Month(Date) - 1
    find_dates
End Sub


Public Sub update_leave_od_data()
On Error GoTo err_handler
      
    Dim emp_err(1000) As Integer
    
    Dim id, fcode As Integer
    Dim dlogdate As Date
    
    Dim dev_log(100) As Long
    
    Dim log_details As String
    
    
    
    stime = TimeValue(Now)
    
    
    Dim sft, sft_bt, sft_et, sft_begin_dur, sft_end_dur As String
    
    Dim tmins As Double
    ''ProgressBar1.Visible = True
    
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    
    paydb.Open pay


''                 Set paydb2 = New ADODB.Connection
''                 Set payrs2 = New ADODB.Recordset
''                 paydb2.Open pay
''
''
''          GoTo 1000
''          Exit Sub
            
    paydb.CommandTimeout = 300
    sql = "update bio_device_shiftlogs set ds_shift_original = '', ds_shift = 'GS',ds_shift_actual = 'GS',ds_status = '', ds_no_of_punches = 0, ds_shift_in = 0 , ds_shift_out = 0,ds_shift_in2 = 0 , ds_shift_out2 = 0,ds_shift_in3 = 0 , ds_shift_out3 = 0,ds_shift_in4 = 0 , ds_shift_out4 = 0, ds_shift_in5 = 0 , ds_shift_out5 = 0 ,ds_shift_in6 = 0 , ds_shift_out6 = 0, ds_per_hrs = 0,ds_od_hrs = 0, ds_sft_hrs = 0,ds_sft_hrs1 = 0,ds_sft_hrs2 = 0  ,ds_sft_hrs3 = 0,ds_sft_hrs4 = 0,ds_sft_hrs5 = 0  ,ds_sft_hrs6 = 0   where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute sql


''    GoTo 100

'''''''''''--------------------------------
    
''    GoTo 500
''    Exit Sub



      sql = "update bio_device_shiftlogs set  ds_shift =  emps_shift  , ds_shift_original  =  emps_shift from bio_device_shiftlogs a, bio_shift_schedule b where emps_date = ds_date and  emps_fpcode = ds_fpcode and  emps_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' "
      paydb.Execute sql
  

pst_qry = "Update b SET  ds_status = 'WO', ds_shift = 'WO',  ds_shift_begintime = '',  ds_shift_endtime = '',    ds_begin_duration = '',    ds_end_duration = '' FROM bio_device_shiftlogs b JOIN emp_mas e ON b.ds_fpcode = e.emp_fpcode " _
       & " Where   e.emp_status = 'A' AND e.emp_company = 1 AND b.ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' AND '" & Format(monthend_date.Value, "MM/dd/yyyy") & "' AND b.ds_date >= e.emp_doj AND   DATEPART(dw, b.ds_date) =  Case UPPER(e.emp_holiday) WHEN 'SUNDAY' THEN 1 WHEN 'MONDAY' THEN 2 WHEN 'TUESDAY' THEN 3 WHEN 'WEDNESDAY' THEN 4 WHEN 'THURSDAY' THEN 5 WHEN 'FRIDAY' THEN 6 WHEN 'SATURDAY' THEN 7 ELSE 0 End"
paydb.Execute pst_qry


        Set paydb2 = New ADODB.Connection
        Set payrs2 = New ADODB.Recordset
        paydb2.Open pay
        
        Dim i, j As Integer
        Dim dt_log As Integer
        Dim inpunch_chk, outpunch_chk, incount, outcount   As Integer


       Dim rs_set As New ADODB.Recordset
       Dim newqry As String
       
    
''    end_date = end_date - 1


''    For idate = st_date - 1 To end_date
        
''        If idate = "06/11/2020" Then
''           MsgBox ("TEst")
''        End If
        
''        sql = "select * from bio_devicelogs where ad_fpcode = 1018 and ad_date between '11/01/2020' and '11/30/2020' order by ad_date,ad_logdate"
        
        sql = "select * from bio_device_shiftlogs where ds_fpcode = 1018 and  ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
''        sql = "select * from bio_device_shiftlogs where  ds_date  = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
        newqry = "select * from bio_device_shiftlogs where ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  order by ds_fpcode,ds_date"
        
        newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where ds_fpcode = ad_fpcode and ds_date = ad_date and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 102 and ds_fpcode < 20000  group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
        '' newqry = "select ds_fpcode,ds_date from bio_device_shiftlogs,bio_devicelogs where  ds_fpcode = 1590  and  ds_fpcode = ad_fpcode and ds_date = ad_date and ds_date  between  '" & Format(st_date.Value, "MM/dd/yyyy") & "' and '" & Format(end_date.Value, "MM/dd/yyyy") & "' and ds_fpcode > 0  group by ds_fpcode,ds_date order by ds_fpcode,ds_date"
        
        rs_set.Open newqry, paydb, 1, 2
        While Not rs_set.EOF
              id = rs_set!ds_fpcode
              idate = rs_set!ds_date
              sql = "select * from bio_device_shiftlogs where ds_fpcode = " & id & " and ds_date = '" & Format(idate, "MM/dd/yyyy") & "' "
              payrs.Open sql, paydb, 1, 2
              While Not payrs.EOF
              dt_log = 1
              i = 1

              id = payrs!ds_fpcode
              sft_from_date = payrs!ds_date
''              If sft_from_date = "04/01/2022" Then
''                 MsgBox ("Wait")
''              End If

''              If id = 3203 Then
''                 MsgBox ("Wait")
''              End If
''


''              pst_qry = "select ad_date,gtime,ad_logslno from " _
''                  & " (select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) > 15  and ad_upd = 'N'  group by ad_date,ad_logslno " _
''                  & " Union All " _
''                  & " select ad_date,ad_logslno,min(ad_logdate) as gtime from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "' and datepart(hh,ad_logdate) < 12  and ad_upd = 'N' group by ad_date,ad_logslno ) a group by  ad_date,gtime,ad_logslno  order by ad_date,gtime,ad_logslno"
              inpunch_chk = 0
              outpunch_chk = 0
              incount = 1
              outcount = 1
              
              
              pst_qry = "select * from bio_devicelogs where ad_fpcode =  " & id & "  and ad_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "'  and ad_upd = 'N' order by ad_date,ad_logdate"
              payrs2.Open pst_qry, paydb2, 1, 2
              While Not payrs2.EOF
                    If Trim(payrs2!ad_punch) = "in" Then
                             If incount = 1 Then
                                 payrs("ds_shift_in") = payrs2!ad_logdate
                                 incount = incount + 1
                                 inpunch_chk = 1
                             ElseIf incount = 2 Then
                                 payrs("ds_shift_in2") = payrs2!ad_logdate
                                 incount = incount + 1
                                 inpunch_chk = 1
                             ElseIf incount = 3 Then
                                 payrs("ds_shift_in3") = payrs2!ad_logdate
                                  incount = incount + 1
                                  inpunch_chk = 1
                             ElseIf incount = 4 Then
                                 payrs("ds_shift_in4") = payrs2!ad_logdate
                                 incount = incount + 1
                                 inpunch_chk = 1
                             ElseIf incount = 5 Then
                                 payrs("ds_shift_in5") = payrs2!ad_logdate
                                 incount = incount + 1
                                 inpunch_chk = 1
                             ElseIf incount = 6 Then
                                 payrs("ds_shift_in6") = payrs2!ad_logdate
                                 incount = incount + 1
                                 inpunch_chk = 1
                             End If
''                             inpunch_chk = 1
                    ElseIf Trim(payrs2!ad_punch) = "out" Then
                             If outcount = 1 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in") Then
                                    If payrs("ds_shift_in") <> "01/01/1900" Then
                                        payrs("ds_shift_out") = payrs2!ad_logdate
                                        payrs("ds_no_of_punches") = outcount
                                        outcount = outcount + 1
                                        outpunch_chk = 1
                                    End If
                                 Else
   ''                                 samepunch = samepunch + 1
                                 End If
                             ElseIf outcount = 2 Then
                                  If payrs2!ad_logdate > payrs("ds_shift_in2") Then '' And payrs("ds_shift_in2") <> "01/01/1900" Then
                                     payrs("ds_shift_out2") = payrs2!ad_logdate
                                     payrs("ds_no_of_punches") = outcount
                                     outcount = outcount + 1
                                     outpunch_chk = 1
                                  End If
''                                 If payrs2!ad_logdate > payrs("ds_shift_in2") And payrs("ds_shift_in2") <> "01/01/1900" Then
''                                    payrs("ds_shift_in2") = 0
''                                    payrs("ds_no_of_punches") = outcount - 1
''                                    outpunch_chk = 1
''                                    dev_log(2) = 0
''                                    incount = incount - 1
''                                 End If
                             ElseIf outcount = 3 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in3") And payrs("ds_shift_in3") <> "01/01/1900" Then
                                    payrs("ds_shift_out3") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    outpunch_chk = 1
                                    
                                 End If
                             ElseIf outcount = 4 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in4") And payrs("ds_shift_in4") <> "01/01/1900" Then
                                    payrs("ds_shift_out4") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    outpunch_chk = 1
                                    
                                 End If
                             ElseIf outcount = 5 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in5") And payrs("ds_shift_in5") <> "01/01/1900" Then
                                    payrs("ds_shift_out5") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    outpunch_chk = 1
                                    
                                 End If
                             ElseIf incount = 6 Then
                                 If payrs2!ad_logdate > payrs("ds_shift_in6") And payrs("ds_shift_in6") <> "01/01/1900" Then
                                    payrs("ds_shift_out6") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    outpunch_chk = 1
                                 End If
                             End If
                             
                    End If
                    If inpunch_chk = 1 Or outpunch_chk = 1 Then
                       dev_log(dt_log) = payrs2!ad_logslno
                       dt_log = dt_log + 1
                    End If
                    payrs2.MoveNext
              Wend
              payrs2.Close
              
              
              If incount > outcount Then
                    pst_qry = "select * from bio_devicelogs where ad_fpcode =  " & id & " and ad_punch = 'out'  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "'  and ad_upd = 'N' and datepart(hh,ad_logdate) <= 9   order by ad_date,ad_logdate"
                    pst_qry = "select * from bio_devicelogs where ad_fpcode =  " & id & " and ad_punch = 'out'  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "'  and ad_upd = 'N' and datepart(hh,ad_logdate) <= 13   order by ad_date,ad_logdate"
                    ''pst_qry = "select * from bio_devicelogs where ad_fpcode =  " & id & " and ad_punch = 'out'  and ad_date = '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "'  and ad_upd = 'N' and datepart(hh,ad_logdate) <= 23   order by ad_date,ad_logdate"
                    
                    payrs2.Open pst_qry, paydb2, 1, 2
                    While Not payrs2.EOF

                          If payrs2!ad_punch = "out" Then
                             If outcount = 1 Then
                                  If payrs("ds_shift_in") <> "01/01/1900" Then
                                     payrs("ds_shift_out") = payrs2!ad_logdate
                                     payrs("ds_no_of_punches") = outcount
                                     outcount = outcount + 1
                                     dev_log(dt_log) = payrs2!ad_logslno
                                     dt_log = dt_log + 1
                                  End If
                             ElseIf outcount = 2 Then
                                 If payrs("ds_shift_in2") <> "01/01/1900" Then
                                    payrs("ds_shift_out2") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    dev_log(dt_log) = payrs2!ad_logslno
                                    dt_log = dt_log + 1
                                 End If
                             ElseIf outcount = 3 Then
                                 If payrs("ds_shift_in3") <> "01/01/1900" Then
                                    payrs("ds_shift_out3") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    dev_log(dt_log) = payrs2!ad_logslno
                                    dt_log = dt_log + 1
                                 End If
                             ElseIf outcount = 4 Then
                                 If payrs("ds_shift_in4") <> "01/01/1900" Then
                                    payrs("ds_shift_out4") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                     dev_log(dt_log) = payrs2!ad_logslno
                                     dt_log = dt_log + 1
                                  End If
                                    
                             ElseIf outcount = 5 Then
                                 If payrs("ds_shift_in5") <> "01/01/1900" Then
                                    payrs("ds_shift_out5") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                    dev_log(dt_log) = payrs2!ad_logslno
                                    dt_log = dt_log + 1
                                  End If
                                    
                             ElseIf incount = 6 Then
                                 If payrs("ds_shift_in6") <> "01/01/1900" Then
                                    payrs("ds_shift_out6") = payrs2!ad_logdate
                                    payrs("ds_no_of_punches") = outcount
                                    outcount = outcount + 1
                                     dev_log(dt_log) = payrs2!ad_logslno
                                     dt_log = dt_log + 1
                                  End If
                             
                             End If
                             
'                             dev_log(dt_log) = payrs2!ad_logslno
 '                            dt_log = dt_log + 1
                          End If
                          payrs2.MoveNext
                    Wend
                    payrs2.Close
              
              End If

              If inpunch_chk > 0 And outpunch_chk > 0 Then payrs.Update
              ''payrs.Update
    
              log_details = "(0"
              For j = 1 To dt_log - 1
                  log_details = log_details + "," + Str(dev_log(j))
              Next
              log_details = log_details + ")"
''              pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and ad_logslno in " & log_details
              pst_qry = "update bio_devicelogs set ad_upd  = 'Y' where  ad_fpcode =  " & id & "  and  ad_upd  = 'N' and ad_logslno in " & log_details & " and ad_date <= '" & Format(sft_from_date + 1, "MM/dd/yyyy") & "'"
              paydb2.Execute pst_qry
              payrs.MoveNext
        Wend
        payrs.Close
        rs_set.MoveNext
      Wend
        rs_set.Close
        
''01/03/2022
1500:
''for updaing missing punches - modified on 03/01/2022
        Dim logdate As Date
        Dim punch As String
        sql = "select * from bio_devicelogs where ad_date >= '" & Format(st_date + 1, "MM/dd/yyyy") & "' and ad_date <= '" & Format(end_date, "MM/dd/yyyy") & "' and ad_upd = 'N' "
        payrs.Open sql, paydb, 1, 2
        While Not payrs.EOF
             id = payrs!ad_fpcode
             sft_from_date = payrs!ad_date
             logdate = payrs!ad_logdate
             punch = Trim(payrs!ad_punch)
''             If id = 3254 Then
''               MsgBox (id)
''             End If
             
             
             
             pst_qry = "select * from bio_device_shiftlogs where ds_fpcode = " & id & " and ds_date = '" & Format(sft_from_date, "MM/dd/yyyy") & "'"
             payrs2.Open pst_qry, paydb2, 1, 2
             If Not payrs2.EOF Then
                  If punch = "in" Then
                      If payrs2("ds_shift_in") = "01/01/1900" Then
                         payrs2("ds_shift_in") = logdate
                         payrs2.Update

                      End If
                  End If
                  If punch = "out" Then
                      If payrs2("ds_shift_out") = "01/01/1900" Then
                         payrs2("ds_shift_out") = logdate
                         payrs2.Update
                      End If
                  End If
                  payrs2.Close
              Else
                  payrs2.Close
              End If
             payrs.MoveNext
             
        Wend
        payrs.Close




        
        
150:
    paydb.CommandTimeout = 300
''    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday_empwise b where emp_decholi_date = ds_date and  emp_decholi_fpcode = ds_fpcode and  emp_decholi_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
    
''    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday b where emp_dec_holiday = ds_date  and emp_dec_holiday Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
    
    sql = "update bio_device_shiftlogs set ds_shift = 'H',ds_shift_actual = 'H',ds_status = 'H',ds_shift_begintime = '00:00',ds_shift_endtime = '00:00',ds_begin_duration  = '',ds_end_duration = '' ,ds_sft_hrs = 0  from bio_device_shiftlogs a, emp_dec_holiday b , emp_mas c   where emp_dec_holiday = ds_date  and emp_dec_holiday Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_date >= emp_doj and  ds_fpcode = emp_fpcode"
    
    paydb.Execute sql
    
    
'' Removed on 01/09/2025
''pst_qry = "select * from emp_mas where emp_company in (1) and emp_status = 'A' "
''    payrs.Open pst_qry, paydb, 1, 2
''    While Not payrs.EOF
''         If payrs("emp_holiday") = "SUNDAY" Then
''                woday = 1
''         ElseIf payrs("emp_holiday") = "MONDAY" Then
''                woday = 2
''         ElseIf payrs("emp_holiday") = "TUESDAY" Then
''                woday = 3
''         ElseIf payrs("emp_holiday") = "WEDNESDAY" Then
''                woday = 4
''         ElseIf payrs("emp_holiday") = "THURSDAY" Then
''                woday = 5
''         ElseIf payrs("emp_holiday") = "FRIDAY" Then
''                woday = 6
''         ElseIf payrs("emp_holiday") = "SATURDAY" Then
''                woday = 7
''         Else
''                woday = 0
''         End If
''         If payrs("emp_cat") = "W" Then
''            sql = "update bio_device_shiftlogs set  ds_status = 'WOH', ds_shift = 'WOH',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = ''  from  bio_device_shiftlogs , emp_mas   where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(monthend_date.Value, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift = 'H' and ds_fpcode = emp_fpcode and ds_date >= emp_doj and datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''            sql = "update bio_device_shiftlogs set  ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = ''  from  bio_device_shiftlogs , emp_mas   where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(monthend_date.Value, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_shift <> 'WOH' and ds_fpcode = emp_fpcode and ds_date >= emp_doj and  datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''
''         Else
''            sql = "update bio_device_shiftlogs set  ds_status = 'WO', ds_shift = 'WO',ds_shift_begintime = '',ds_shift_endtime = '',ds_begin_duration  = '',ds_end_duration = ''  from  bio_device_shiftlogs , emp_mas    where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(monthend_date.Value, "MM/dd/yyyy") & "' and ds_fpcode = " & payrs("emp_fpcode") & " and ds_fpcode = emp_fpcode and ds_date >= emp_doj and datepart(dw,ds_date) =  " & woday
''            paydb.Execute sql
''         End If
''
''         payrs.MoveNext
''    Wend
''    payrs.Close




    Dim emp As Integer
    Dim skip_empcode As String
    skip_empcode = "("
    emp = 0
    i = 0
    
100:
    
        sql = "update bio_device_shiftlogs  set ds_sft_hrs1=  (datediff(minute,ds_shift_in, ds_shift_out)/60)+convert(decimal,(datediff(minute,ds_shift_in, ds_shift_out) -(datediff(minute,ds_shift_in, ds_shift_out)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out > ds_shift_in and ds_shift_in > 0 "
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs2 = (datediff(minute,ds_shift_in2, ds_shift_out2)/60)+convert(decimal,(datediff(minute,ds_shift_in2, ds_shift_out2) -(datediff(minute,ds_shift_in2, ds_shift_out2)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out2 > ds_shift_in2 and ds_shift_in2 > 0 "
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs3 = (datediff(minute,ds_shift_in3, ds_shift_out3)/60)+convert(decimal,(datediff(minute,ds_shift_in3, ds_shift_out3) -(datediff(minute,ds_shift_in3, ds_shift_out3)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out3 > ds_shift_in3 and ds_shift_in3 > 0 "
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs4 = (datediff(minute,ds_shift_in4, ds_shift_out4)/60)+convert(decimal,(datediff(minute,ds_shift_in4, ds_shift_out4) -(datediff(minute,ds_shift_in4, ds_shift_out4)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out4 > ds_shift_in4 and ds_shift_in4 > 0 "
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs5 = (datediff(minute,ds_shift_in5, ds_shift_out5)/60)+convert(decimal,(datediff(minute,ds_shift_in5, ds_shift_out5) -(datediff(minute,ds_shift_in5, ds_shift_out5)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out5 > ds_shift_in5 and ds_shift_in5 > 0 "
        paydb.Execute sql
        sql = "update bio_device_shiftlogs  set ds_sft_hrs6 = (datediff(minute,ds_shift_in6, ds_shift_out6)/60)+convert(decimal,(datediff(minute,ds_shift_in6, ds_shift_out6) -(datediff(minute,ds_shift_in6, ds_shift_out6)/60 * 60)))/100 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift_out6 > ds_shift_in6 and ds_shift_in6 > 0 "
        paydb.Execute sql

    
    
    Dim firsttime, endtime As Double
    Dim per_totmins, per_hrs, per_mins, per_time As Double
 ''for updating Permission entries
    pst_qry = "select * from bio_emp_permissions where empp_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!empp_fpcode
          idate = payrs!empp_date
          firsttime = Int(Val(payrs!empp_fromtime)) * 60 + (Val(payrs!empp_fromtime) - Int(Val(payrs!empp_fromtime))) * 100
          endtime = Int(Val(payrs!empp_totime)) * 60 + (Val(payrs!empp_totime) - Int(Val(payrs!empp_totime))) * 100
        
          per_totmins = endtime - firsttime
          per_hrs = Int(Round(per_totmins) / 60)
          per_mins = Round(per_totmins) - (per_hrs * 60)
          per_time = per_hrs + (per_mins / 100)
          
           sql = "select * from bio_device_shiftlogs where ds_date = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode = " & id & "  and ds_shift_in is not null "
           payrs2.Open sql, paydb, 1, 2
           While Not payrs2.EOF
                 payrs2("ds_per_hrs") = IIf(per_time > 0, per_time, 0)
                 payrs2.Update
                 payrs2.MoveNext
          Wend
          payrs2.Close
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    
    Dim fday_od As Integer
    Dim od_totmins, od_hrs, od_mins, od_time As Double

    pst_qry = "select * from bio_emp_oddetails where empod_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!empod_fpcode
          idate = payrs!empod_date
          
          firsttime = Int(Val(payrs!empod_fromtime)) * 60 + (Val(payrs!empod_fromtime) - Int(Val(payrs!empod_fromtime))) * 100
          endtime = Int(Val(payrs!empod_totime)) * 60 + (Val(payrs!empod_totime) - Int(Val(payrs!empod_totime))) * 100
          
''          If endtime - firsttime >= 540 Then
''             fday_od = 1
''          Else
''             fday_od = 0
''          End If
          
          od_totmins = endtime - firsttime
          od_hrs = Int(Round(od_totmins) / 60)
          od_mins = Round(od_totmins) - Int((od_hrs * 60))
          od_time = od_hrs + (od_mins / 100)

          
           sql = "select * from bio_device_shiftlogs where ds_date = '" & Format(idate, "MM/dd/yyyy") & "' and ds_fpcode = " & id & "  and ds_shift_in is not null "
           payrs2.Open sql, paydb, 1, 2
           While Not payrs2.EOF
                 payrs2("ds_od_hrs") = IIf(od_time > 0, od_time, 0)
                 payrs2.Update
                 payrs2.MoveNext
           Wend
          payrs2.Close
          payrs.MoveNext
    Wend
    payrs.Close
 
 
    Dim wmins1, wmins2, wmins3, wmins4, wmins5, wmins6, permins, odmins, totmins As Double
    sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_sft_hrs1+ds_sft_hrs2+ds_sft_hrs3+ds_sft_hrs4+ds_sft_hrs5+ds_sft_hrs6+ds_per_hrs+ds_od_hrs >0"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          
          wmins1 = Int(Val(payrs!ds_sft_hrs1)) * 60 + (Val(payrs!ds_sft_hrs1) - Int(Val(payrs!ds_sft_hrs1))) * 100
          wmins2 = Int(Val(payrs!ds_sft_hrs2)) * 60 + (Val(payrs!ds_sft_hrs2) - Int(Val(payrs!ds_sft_hrs2))) * 100
          wmins3 = Int(Val(payrs!ds_sft_hrs3)) * 60 + (Val(payrs!ds_sft_hrs3) - Int(Val(payrs!ds_sft_hrs3))) * 100
          wmins4 = Int(Val(payrs!ds_sft_hrs4)) * 60 + (Val(payrs!ds_sft_hrs4) - Int(Val(payrs!ds_sft_hrs4))) * 100
          wmins5 = Int(Val(payrs!ds_sft_hrs5)) * 60 + (Val(payrs!ds_sft_hrs5) - Int(Val(payrs!ds_sft_hrs5))) * 100
          wmins6 = Int(Val(payrs!ds_sft_hrs6)) * 60 + (Val(payrs!ds_sft_hrs6) - Int(Val(payrs!ds_sft_hrs6))) * 100
          permins = Int(Val(payrs!ds_per_hrs)) * 60 + (Val(payrs!ds_per_hrs) - Int(Val(payrs!ds_per_hrs))) * 100
          odmins = Int(Val(payrs!ds_od_hrs)) * 60 + (Val(payrs!ds_od_hrs) - Int(Val(payrs!ds_od_hrs))) * 100
          totmins = wmins1 + wmins2 + wmins3 + wmins4 + wmins5 + wmins6 + permins + odmins
          tot_hrs = Int(totmins / 60)
          tot_mins = totmins - (tot_hrs * 60)
          tot_time = tot_hrs + (tot_mins / 100)
          
          
          payrs("ds_sft_hrs") = IIf(tot_time > 0, tot_time, 0)
          payrs.Update
    
          payrs.MoveNext
    Wend
    payrs.Close
 
 

    
    
    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE ds_sft_hrs = 0 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO'  and ds_shift <> 'WOH' "
    paydb.Execute sql

    sql = "update bio_device_shiftlogs set ds_status = 'A' WHERE  ds_sft_hrs < 3 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' "
    paydb.Execute sql

    sql = "update bio_device_shiftlogs set ds_status = '' from  bio_device_shiftlogs , emp_mas  WHERE ds_sft_hrs = 0 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(Now - 1, "MM/dd/yyyy") & "' and ds_shift <> 'WO'  and ds_shift <> 'WOH' and ds_fpcode = emp_fpcode  and ds_date < emp_doj"
    paydb.Execute sql
    
    
    
    
210:

    
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'A SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 5 and DATEpart(hour,ds_shift_in) <= 6 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'B SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 13 and DATEpart(hour,ds_shift_in) <= 16 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'C SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 21 and DATEpart(hour,ds_shift_in) <= 22 "
    paydb.Execute sql
    
''FOR 7.00 AM to 03.00 PM
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00AM-03.00PM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 07 and DATEpart(hour,ds_shift_in) <= 07"
    paydb.Execute sql
''FOR 8.00 AM to 05.00 PM
''    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00AM-04.00PM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 08 and DATEpart(hour,ds_shift_in) <= 08"
''    paydb.Execute sql
    
''    sql = "update bio_device_shiftlogs set ds_shift_actual = '07to8AM-03to4PM'  from bio_device_shiftlogs , bio_empmas  WHERE bioemp_team = 'WORKER' and bioemp_fpcode = ds_fpcode and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 07 and DATEpart(hour,ds_shift_in) <= 08"
''    paydb.Execute sql



    
    sql = "update bio_device_shiftlogs set ds_shift_actual = '06.00PM-06.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 17 and DATEpart(hour,ds_shift_in) <= 18"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00PM-07.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 18 and DATEpart(hour,ds_shift_in) <= 19"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00PM-08.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 19 and DATEpart(hour,ds_shift_in) <= 20"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'A SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 5 and DATEpart(hour,ds_shift_in) <= 6 and   DATEpart(hour,ds_shift_out) >= 13  and  DATEpart(hour,ds_shift_out) <= 16 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'B SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 13 and DATEpart(hour,ds_shift_in) <= 16 and  DATEpart(hour,ds_shift_out) >= 21  and  DATEpart(hour,ds_shift_out) <= 23 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = 'C SHIFT'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 21 and DATEpart(hour,ds_shift_in) <= 22 and  DATEpart(hour,ds_shift_out) >= 5  and  DATEpart(hour,ds_shift_out) <= 10 "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '06.00PM-06.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 17 and DATEpart(hour,ds_shift_in) <= 18 and  DATEpart(hour,ds_shift_out) >= 6  and  DATEpart(hour,ds_shift_out) <= 7"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '07.00PM-07.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 18 and DATEpart(hour,ds_shift_in) <= 19 and  DATEpart(hour,ds_shift_out) >= 7  and  DATEpart(hour,ds_shift_out) <= 8"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_shift_actual = '08.00PM-08.00AM'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 19 and DATEpart(hour,ds_shift_in) <= 20 and  DATEpart(hour,ds_shift_out) >= 8  and  DATEpart(hour,ds_shift_out) <= 9"
    paydb.Execute sql

    sql = "update bio_device_shiftlogs set ds_shift_actual = 'Unshift'  from bio_device_shiftlogs WHERE  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "' and DATEpart(hour,ds_shift_in) >= 15 and DATEpart(hour,ds_shift_in) <= 20 and  DATEpart(hour,ds_shift_out) >= 2  and  DATEpart(hour,ds_shift_out) <= 4 "
    paydb.Execute sql


''End



    sql = "update bio_device_shiftlogs set ds_status = 'H' WHERE   ds_shift = 'H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute sql

''FOR A,B & C SHIFT - COMMON FOR STAFF & WORKER - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    
''FOR STAFF - 1/2 DAY PRESENT - for SINGLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    
    
''FOR STAFF - 1/2 DAY PRESENT - for DOUBLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F'"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F'"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M'"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 2 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M'"
    paydb.Execute sql
        
    
    
''FOR UNSHIFT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift')"
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and   ds_sft_hrs >3.30  and  ds_sft_hrs < 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift')"
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs > 7.44  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') "
    sql = "update bio_device_shiftlogs set ds_status = 'P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and   ds_sft_hrs > 7.44  and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual  in ('Unshift') "
    
    paydb.Execute sql
    
''FOR STAFF - 1/2 DAY PRESENT - for DOUBLE IN / OUT punches
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'F' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 8.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') and bioemp_gender = 'M' "
    paydb.Execute sql
    
    
''FOR DECLARE HOLIDAY - STAFF - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs > 4 and  ds_sft_hrs < 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 7.40 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H' "
    paydb.Execute sql
    
    
''FOR WORKER - 1/2 DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = '½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and   ds_sft_hrs >3.30  and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift <> 'WO' and ds_shift <> 'WOH'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WO½P'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >3.30  and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_shift = 'WO' and ds_shift <> 'WOH'  "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = '½HP'  from bio_device_shiftlogs , bio_empmas   WHERE  ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >3.30 and  ds_sft_hrs < 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'HP' from bio_device_shiftlogs , bio_empmas    WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.46 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'H'"
    paydb.Execute sql


''    sql = "update bio_device_shiftlogs set ds_status = 'P' WHERE ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH'"
''    paydb.Execute sql
''    sql = "update bio_device_shiftlogs set ds_status = 'WOP' WHERE ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'"
''    paydb.Execute sql
''    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' WHERE ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'"
''    paydb.Execute sql
''    sql = "update bio_device_shiftlogs set ds_status = 'P' WHERE ds_sft_hrs >= 7.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO'"
''    paydb.Execute sql
    
    
''FOR B & C SHIFT - COMMON FOR STAFF & WORKER - FULL DAY PRESENT
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P'  and ds_shift_actual in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    

    '' for  STAFF - FEMALE
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches = 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches > 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches > 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'F' "
    paydb.Execute sql
    
    
    '' for  STAFF - MALE
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches = 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches = 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches = 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and  ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_no_of_punches > 1   and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.00 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH'  and ds_no_of_punches > 1 and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team = 'STAFF' and ds_sft_hrs >= 8.30 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_no_of_punches > 1  and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')  and bioemp_gender = 'M' "
    paydb.Execute sql
    
    
    
    
    
    
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H'  and  ds_shift <> 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and  ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO' and ds_status = 'P' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT')"
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'P' from bio_device_shiftlogs , bio_empmas WHERE ds_fpcode = bioemp_fpcode and bioemp_team <> 'STAFF' and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and  ds_shift <> 'H' and  ds_shift <> 'WOH'  and  ds_shift <> 'WO' and ds_shift_actual not in ('A SHIFT','B SHIFT','C SHIFT') "
    paydb.Execute sql
    
    
    
    
    sql = "update bio_device_shiftlogs set ds_status = 'WOHP' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.45 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_status = 'WOH' "
    paydb.Execute sql
    

    
'' for OD Assignment
    sql = "update bio_device_shiftlogs set ds_status = 'P(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and  ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4   and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
    paydb.Execute sql
    sql = "update bio_device_shiftlogs set ds_status = 'WOP(OD)' from bio_device_shiftlogs , bio_empmas  WHERE ds_fpcode = bioemp_fpcode and ds_sft_hrs >= 7.49 and  ds_od_hrs >= 4 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and  ds_shift = 'WO'"
    paydb.Execute sql
    

    
    Dim leave, leavetype As String
 ''for updating leave entries
    pst_qry = "select * from bio_empleave where emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    pst_qry = "select * from bio_empleave , bio_device_shiftlogs  where ds_fpcode = emp_fpcode and ds_date= emp_leave_date and emp_leave_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'  order by emp_leave_date"

    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!emp_fpcode
''          If id = 1211 Then
''            If payrs("ds_date") = "2019-02-01" Then
''              MsgBox ("TEst")
''            End If
''          End If
''          leavetype = payrs!emp_leave_type
''          If payrs!emp_leave_no = 0 Then
''             leave = payrs!emp_leave_type
''          Else
''             leave = IIf(payrs!emp_leave_period = "F", "EL", "½EL")
''          End If
''          If Format(payrs!emp_leave_date, "MM/dd/yyyy") = "10/06/2015" And id = 1002 Then
''              MsgBox ("Wait")
''          End If
''
          If payrs!emp_leave_no > 0 Then
               If payrs!emp_leave_type = "EL" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "EL", "½L")
               ElseIf payrs!emp_leave_type = "L" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "L", "½L")
               ElseIf payrs!emp_leave_type = "L" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "L", "½L")
                  
                  
               ElseIf payrs!emp_leave_type = "ML" Then
                  leavetype = IIf(payrs!emp_leave_period = "F", "ML", "½ML")
               ElseIf payrs!emp_leave_type = "LAYOFF" Then
                  leavetype = "LAYOFF"
               Else
                  leavetype = payrs!emp_leave_type
               End If
          Else
               If UCase(payrs!emp_leave_type) = "LAYOFF" Then
                   leavetype = payrs!emp_leave_type
               ElseIf payrs!emp_leave_type = "CL" Then
                   leavetype = "EL"
               ElseIf payrs!emp_leave_type = "½CL" Then
                   leavetype = "½EL"
               Else
                  leavetype = payrs!emp_leave_type
               End If
          End If
''          sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          If leavetype = "½L" And payrs!ds_status = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½L' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf leavetype = "½EL" And payrs!ds_status = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½P" And leavetype = "½C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And leavetype = "½C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = 'C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And (leavetype = "½EL" Or leavetype = "½CL") Then
             sql = "update bio_device_shiftlogs set ds_status = '½C.H½EL' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And leavetype = "½L" Then
             sql = "update bio_device_shiftlogs set ds_status = '½C.H½L' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" And leavetype = "½L" Then
             sql = "update bio_device_shiftlogs set ds_status = '½C.H½L' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
             
          ElseIf payrs!ds_status = "P" And leavetype <> "SA" Then
             sql = "update bio_device_shiftlogs set ds_status = 'P' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "P" And leavetype = "SA" Then
             sql = "update bio_device_shiftlogs set ds_status = 'SA' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
             
          ElseIf payrs!ds_status <> "A" And payrs!ds_status <> "" Then
             sql = "update bio_device_shiftlogs set ds_status = ds_status+'" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½L" Then
             sql = "update bio_device_shiftlogs set ds_status = '½L½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½P" Then
             sql = "update bio_device_shiftlogs set ds_status = '½P½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "A" And leavetype = "½L" Then
             sql = "update bio_device_shiftlogs set ds_status = '½L½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "" And leavetype = "½L" Then
             sql = "update bio_device_shiftlogs set ds_status = '½L½A' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
                       
          Else
             sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!emp_leave_date, "MM/dd/yyyy") & "'"
          End If
          paydb2.Execute sql
          payrs.MoveNext
    Wend
    payrs.Close
 
 
 
 ''for updating CH. leave entries
    pst_qry = "select * from bio_emp_chleave where empch_ch_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"
    pst_qry = "select * from bio_emp_chleave , bio_device_shiftlogs  where ds_fpcode = empch_fpcode and ds_date= empch_ch_date and empch_ch_date between '" & Format(st_date, "MM/dd/yyyy") & "' and '" & Format(end_date, "MM/dd/yyyy") & "'"

    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          id = payrs!empch_fpcode
''          If id = 1429 Then
''             MsgBox ("Wait")
''          End If
          leavetype = IIf(payrs!emp_ch_period = "F", "C.H", "½C.H")
          If payrs!ds_status = "A" Then
             sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = 'C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "C.H" Then
             sql = "update bio_device_shiftlogs set ds_status = 'C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½L½A" Then
             sql = "update bio_device_shiftlogs set ds_status = '½L½C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          ElseIf payrs!ds_status = "½EL½A" Then
             sql = "update bio_device_shiftlogs set ds_status = '½EL½C.H' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          Else
             sql = "update bio_device_shiftlogs set ds_status =  ds_status + '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          End If

''          sql = "update bio_device_shiftlogs set ds_status = '" & leavetype & "' where ds_fpcode =  '" & id & "' and ds_date =  '" & Format(payrs!empch_ch_date, "MM/dd/yyyy") & "'"
          paydb.Execute sql
          payrs.MoveNext
    Wend
    payrs.Close
 
''for 1/2 Present only
  sql = "update bio_device_shiftlogs set ds_status = '½P½A' from bio_device_shiftlogs where ds_Status = '½P' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
  paydb.Execute sql
 
 
 ''for 1/2CH and 1/2 ABS
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½A' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs = 0"
  paydb.Execute sql
 
 ''for 1/2CH and 1/2 PRESENT
  sql = "update bio_device_shiftlogs set ds_status = '½C.H½P' from bio_device_shiftlogs where ds_Status = '½C.H' and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ds_sft_hrs > 3.4"
  paydb.Execute sql
 
 
 
 
1000:
 

    ''for updating bio_attendlogs_daily
        Dim aday As String
        sql = "select * from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' order by ds_date"
        payrs.Open sql, paydb, 1, 2
        While Not payrs.EOF
              If Month(payrs!ds_date) = cmb_month.ItemData(cmb_month.ListIndex) Then
                 aday = Trim(Str(Day(payrs!ds_date)))
                 pst_qry = "update bio_attendlogs set a_day" & aday & " = '" & payrs!ds_status & "',a_in_day" & aday & " = '" & Format(payrs!ds_shift_in, "MM/dd/yyyy HH:MM:SS") & "' ,a_out_day" & aday & " = '" & Format(payrs!ds_shift_out, "MM/dd/yyyy HH:MM:SS") & "' where a_fpcode = " & payrs!ds_fpcode & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
                 paydb.Execute pst_qry

               End If
            payrs.MoveNext
        Wend
        payrs.Close
    'end


 
700:


'' NEW ADDITION ON 03-02-2022
''-------------
   Dim dayfind, dayfind_intime, dayfind_outtime As String
    
    
    Dim blankdata As Integer
    Dim absent, leavedata As Single
    
If ProcessType = "F" Then
    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
''            If payrs.Fields("a_fpcode") = 3326 Then
''               MsgBox ("Wait")
''            End If
            absent = 0
            leavedata = 0
            blankdata = 0
            
            For i = 1 To 31
                dayfind = "a_day" & i
                If payrs.Fields(dayfind) = "A" Then
                    absent = absent + 1
                ElseIf payrs.Fields(dayfind) = "L" Then
                    leavedata = leavedata + 1
                ElseIf payrs.Fields(dayfind) = "" Then
                    blankdata = blankdata + 1
                ElseIf payrs.Fields(dayfind) = "P" Then
                    blankdata = 0
                    absent = 0
                    leavedata = 0
                Else
''                    If absent >= 4 And payrs.Fields("a_present") > 0 Then
''                       If payrs.Fields(dayfind) = "WO" Then
''                           payrs.Fields(dayfind) = "A"
''                           payrs.Update
''                       End If
''                    End If

                    If absent >= 3 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = "A"
                           payrs.Update
''                           absent = 0
                    End If
                    If (blankdata + absent + leavedata) >= 4 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = ""
                           blankdata = 0
                           leavedata = 0
''                           absent = 0
                           payrs.Update
                    End If
                    
'' Addition on 12/07/2024
''Start
                    If absent >= 3 And payrs.Fields(dayfind) = "H" Then
                           payrs.Fields(dayfind) = "A"
                           payrs.Update
''                           absent = 0
                    End If
                    If (blankdata + absent) >= 3 And payrs.Fields(dayfind) = "H" Then
                           payrs.Fields(dayfind) = "A"
                           blankdata = 0
''                           absent = 0
                           payrs.Update
                    End If
                    If (blankdata + absent + leavedata) >= 4 And payrs.Fields(dayfind) = "WO" Then
                           payrs.Fields(dayfind) = ""
                           blankdata = 0
                           ''absent = 0
                           leavedata = 0
                           payrs.Update
                    End If
''End
''                    absent = 0
                End If
            Next
            
''            Dim presentdata As Integer
''                                blankdata = 0
''                    absent = 0
''                    leavedata = 0
''                    presentdata = 0
''
''            For i = 1 To 31
''                dayfind = "a_day" & i
''                If payrs.Fields(dayfind) = "A" Then
''                    absent = absent + 1
''                ElseIf payrs.Fields(dayfind) = "L" Then
''                    leavedata = leavedata + 1
''                ElseIf payrs.Fields(dayfind) = "" Then
''                    blankdata = blankdata + 1
''                ElseIf payrs.Fields(dayfind) = "P" Then
''                    presentdata = presentdata + 1
''                End If
''
''                    If (blankdata + absent + leavedata) >= 3 And payrs.Fields(dayfind) = "WO" Then
''                           payrs.Fields(dayfind) = "A"
''                           blankdata = 0
''                           leavedata = 0
''                           presentdata = 0
''                           payrs.Update
''                    End If
''
''
''            Next
            
            payrs.MoveNext
        Wend
    End If
    payrs.Close
End If
''--------------



    
    Dim present, hop, wop, cl, sl, h, ch, layoff, wo, pl, hope, el, woh, ml, totdays, eligible_days, emer_leave  As Single
 
''    sql = "select * from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    sql = "select * from bio_attendlogs ,emp_mas  where a_fpcode = emp_fpcode and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
    payrs.Open sql, paydb, 1, 2
    If Not payrs.EOF Then
       While Not payrs.EOF
''
''             If payrs.Fields("a_fpcode") = 2091 Then
''               MsgBox ("Wait")
''            End If
''
            present = 0
            absent = 0
            hop = 0
            wop = 0
            cl = 0
            sl = 0
            h = 0
            ch = 0
            layoff = 0
            wo = 0
            pl = 0
            hope = 0
            el = 0
            woh = 0
            ml = 0
            emer_leave = 0
            For i = 1 To 31
                dayfind = "a_day" & i
                If payrs.Fields(dayfind) = "P" Or payrs.Fields(dayfind) = "OD" Or payrs.Fields(dayfind) = "P(OD)" Or payrs.Fields(dayfind) = "½P(OD)" Or payrs.Fields(dayfind) = "A(OD)" Then
                    present = present + 1
                ElseIf payrs.Fields(dayfind) = "A" Then
                    absent = absent + 1
                ElseIf payrs.Fields(dayfind) = "L" Or payrs.Fields(dayfind) = "PLP" Or payrs.Fields(dayfind) = "PL" Then
                    pl = pl + 1
                ElseIf payrs.Fields(dayfind) = "½L" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½LP" Or payrs.Fields(dayfind) = "½PL" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "EL" Or payrs.Fields(dayfind) = "CL" Or payrs.Fields(dayfind) = "CL½P" Or payrs.Fields(dayfind) = "CLP" Then
                    el = el + 1
                ElseIf payrs.Fields(dayfind) = "½CL" Then
                    absent = absent + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½CLP" Or payrs.Fields(dayfind) = "½CL½P" Or payrs.Fields(dayfind) = "½EL" Then
                    present = present + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.HP" Then
                    present = present + 0.5
                    ch = ch + 0.5
                ElseIf payrs.Fields(dayfind) = "SL" Or payrs.Fields(dayfind) = "SLP" Then
                    sl = sl + 1
                ElseIf payrs.Fields(dayfind) = "½SLP" Then
                    sl = sl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "H" Then
                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "HP" Or payrs.Fields(dayfind) = "HP(OD)" Then
                    hop = hop + 1
                    h = h + 1

                ElseIf payrs.Fields(dayfind) = "HPE" Then
                    hope = hope + 1
                    h = h + 1
                    hop = hop + 1
                
                ElseIf payrs.Fields(dayfind) = "½HP" Then
''                    hop = hop + 0.5
                    hop = hop + 0.5
                    h = h + 1
                
                ElseIf payrs.Fields(dayfind) = "½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "Layoff" Or payrs.Fields(dayfind) = "LayoffP" Or payrs.Fields(dayfind) = "LAYOFF" Or payrs.Fields(dayfind) = "½PLAYOFF" Then
                    layoff = layoff + 1
                ElseIf payrs.Fields(dayfind) = "C.H" Or payrs.Fields(dayfind) = "C.H½P" Or payrs.Fields(dayfind) = "C.HP" Or payrs.Fields(dayfind) = "C.HP(OD)" Then
                    ch = ch + 1
                ElseIf payrs.Fields(dayfind) = "HOP" Or payrs.Fields(dayfind) = "H½P(OD)" Then
                    hop = hop + 1
                ElseIf payrs.Fields(dayfind) = "WOP" Or payrs.Fields(dayfind) = "WOP(OD)" Or payrs.Fields(dayfind) = "WO(OD)" Then
                    wop = wop + 1
                ElseIf payrs.Fields(dayfind) = "WO" Then
                    wo = wo + 1
                ElseIf payrs.Fields(dayfind) = "WO½P" Then
                    wop = wop + 0.5
                    wo = wo + 0.5
                ElseIf payrs.Fields(dayfind) = "WOH" Then
                    woh = woh + 1
                    wo = wo + 1
''                    h = h + 1
                ElseIf payrs.Fields(dayfind) = "WOHP" Then
                    woh = woh + 1
''                    h = h + 1
                    wop = wop + 1
''                ElseIf payrs.Fields(dayfind) = "WOHPE" Then
''                    woh = woh + 1
''                    h = h + 1
''                    wop = wop + 1
''                    hope = hope + 1


''                ElseIf payrs.Fields(dayfind) = "½C.H" Then
''                    ch = ch + 0.5
''                    present = present + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½C.H" Then
                    ch = ch + 0.5
                    el = el + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½EL½L" Then
                    pl = pl + 0.5
                    el = el + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "P½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PC.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½PC.H½C.H" Then
                    ch = ch + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½EL" Then
                    el = el + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½L" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½C.H" Then
                    pl = pl + 0.5
                    ch = ch + 0.5
                ElseIf payrs.Fields(dayfind) = "½P½A" Or payrs.Fields(dayfind) = "½A½P" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½A" Or payrs.Fields(dayfind) = "½A½EL" Then
                    el = el + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½A" Or payrs.Fields(dayfind) = "½A½L" Then
                    pl = pl + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.H½A" Or payrs.Fields(dayfind) = "½A½C.H" Then
                    ch = ch + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½C.H½P" Or payrs.Fields(dayfind) = "½P½C.H" Then
                    present = present + 0.5
                    ch = ch + 0.5
                
                ElseIf payrs.Fields(dayfind) = "½A½OD" Then
                    present = present + 0.5
                    absent = absent + 0.5
                ElseIf payrs.Fields(dayfind) = "½L½OD" Then
                    pl = pl + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½EL½OD" Then
                    el = el + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "½ML" Or payrs.Fields(dayfind) = "½P½ML" Or payrs.Fields(dayfind) = "½ML½P" Then
                    ml = ml + 0.5
                    present = present + 0.5
                ElseIf payrs.Fields(dayfind) = "ML" Then
                    ml = ml + 1
                ElseIf payrs.Fields(dayfind) = "EM.L" Then
                    emer_leave = emer_leave + 1
                ElseIf payrs.Fields(dayfind) = "½HP½L" Or payrs.Fields(dayfind) = "½L½HP" Then
                    pl = pl + 0.5
                    hop = hop + 0.5
                    h = h + 0.5
                End If
            
 
 Next
            
            payrs("a_present") = present
            payrs("a_hop") = hop
            payrs("a_hpe") = hope
            payrs("a_wop") = wop
            payrs("a_el") = el
            payrs("a_pl") = pl
            payrs("a_ml") = ml
            payrs("a_holiday") = h
            payrs("a_ch") = ch
            payrs("a_layoff") = layoff
            payrs("a_absent") = absent
            payrs("a_wo") = wo
            payrs("a_woh") = woh
            payrs("a_emer_leave_days") = emer_leave
            
            payrs("a_month_days") = mdays
            
            



            If payrs("emp_cat") = "W" Then
''               totdays = present + wop + wo + wop + ch + ml + woh + hop + wohp
               totdays = present + wop + wo + wop + ch + ml + hop
            Else
               totdays = present + wop + wo + ch + ml + woh
            End If
                        
            totdays = totdays + h
            
            payrs("a_total_days") = totdays
            
            
            If totdays >= mdays Then
                payrs("a_eligible_days") = mdays
                payrs("a_salary_days") = mdays
            If payrs("emp_cat") = "W" Then
                   payrs("a_ot_days") = totdays - mdays
                Else
                   payrs("a_ot_days") = 0
                End If
           Else
            
                payrs("a_eligible_days") = totdays
                payrs("a_salary_days") = totdays
                payrs("a_ot_days") = 0
            End If
            
            payrs.Update
            payrs.MoveNext
        Wend
    End If
    payrs.Close

 '' For reversing Holiday present if avail CH
    Dim holidaych As Single
    sql = "select * ,case when emp_ch_period = 'F' then 1 else 0.5 end as Holidaych from bio_device_shiftlogs a, bio_emp_chleave b where  ds_fpcode =  empch_fpcode and ds_shift = 'H' and ds_date = empch_worked_date and ds_date Between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' order by ds_date"
    payrs.Open sql, paydb, 1, 2
    While Not payrs.EOF
          idate = payrs!ds_date
          id = payrs!ds_fpcode
          holidaych = payrs!holidaych
          pst_qry = "update bio_attendlogs set a_hop  = a_hop - " & holidaych & "  where a_fpcode = " & id & " and a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " "
          paydb2.Execute pst_qry
          payrs.MoveNext
    Wend
    payrs.Close
''    pst_qry = "update bio_attendlogs set a_hop  = 0 where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_hop < 0 "
''   paydb2.Execute pst_qry
    
'end
 
 
    Dim entime2 As Date
    endtime2 = TimeValue(Now)
 
    MsgBox ("Updation completed... Process Start by " + Str(stime) + " ...  end by " + Str(endtime2))
    
 




 
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000 and ds_fpcode < 20000  and  ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in > ds_shift_out and ds_sft_hrs <2 and ds_date < '" & Format(end_date, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          If Format(payrs!ds_date, "MM/dd/yyyy") <> Format(Now, "MM/dd/yyyy") Then
             MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          End If
          payrs.MoveNext
    Wend
    payrs.Close


    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b  where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000 and ds_fpcode < 20000 and ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in = 0  and ds_shift_OUT > 0 and ds_sft_hrs <2  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic
    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close

    
    pst_qry = "select * from bio_device_shiftlogs  a , bio_empmas b    where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000  and ds_fpcode < 20000   and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in2 > 0  and ds_shift_out2 = 0 and ds_sft_hrs <2 and  ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    pst_qry = "select * from bio_device_shiftlogs  a , bio_empmas b    where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000 and ds_fpcode < 20000   and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in2 > 0  and ds_shift_out2 = 0 and ds_sft_hrs >24 and  ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b   where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000 and ds_fpcode < 20000 and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in3 > 0  and ds_shift_out3 = 0 and ds_sft_hrs < 3  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          i = i + 1
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    pst_qry = "select * from bio_device_shiftlogs a , bio_empmas b   where  ds_fpcode = bioemp_fpcode and ds_fpcode >=1000  and ds_fpcode < 20000 and   ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'  and ds_shift_in3 > 0  and ds_shift_out3 = 0 and ds_sft_hrs > 24  and ds_date < '" & Format(end_date.Value, "MM/dd/YYYY") & "'"
    payrs.Open pst_qry, paydb, adOpenDynamic, adLockOptimistic

    While Not payrs.EOF
          i = i + 1
          MsgBox ("Problem in " + Str(payrs!bioemp_fpcode) + " " + payrs!bioemp_name + " in the date of " + Format(payrs!ds_date, "dd/MM/yyyy"))
          payrs.MoveNext
    Wend
    payrs.Close
    
    
    
    skip_empcode = skip_empcode + ")"

 

    ''MsgBox ("Updated...")
    Exit Sub
err_handler:
       MsgBox (sql)
       MsgBox ("Problem in the ID :  " & id & "  in the date of  " & sft_from_date)
      chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume

End Sub



Public Sub upload_biometric_data()
On Error GoTo err_handler
    
    Dim idcode As Long
    Dim id, fcode As Integer
    Dim dlogdate As Date
    
    Dim dev_log(100) As Long
    
    Dim log_details As String
    
    
    
    Dim stime, etime As Date
    
    stime = TimeValue(Now)
    
    
    Dim sft, sft_bt, sft_et, sft_begin_dur, sft_end_dur As String
    
    ''ProgressBar1.Visible = True
    
''    If cmb_month.Text = "" Then
''       MsgBox ("Select Month...")
''       Exit Sub
''    End If
''    If cmb_year.Text = "" Then
''       MsgBox ("Select Year...")
''       Exit Sub
''    End If
''    find_dates
    
''    ProgressBar1.Value = ProgressBar1.Min
    
    Dim tablename, tablename2 As String
    tablename = "devicelogs_" + Trim(Str(cmb_month.ItemData(cmb_month.ListIndex))) + "_" + Trim(Str(cmb_year.Text))
    
    tablename2 = "devicelogs_" + Trim(Str(Month(end_date.Value))) + "_" + Trim(Str(Year(end_date.Value)))
    
    Dim dsnmdb As String
    Dim mdbrs As New ADODB.Recordset
    
    Set paydb = New ADODB.Connection
    Set payrs = New ADODB.Recordset
    Set payrsnew = New ADODB.Recordset
    
    paydb.Open pay

  ''  dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.11.31\eSSL\eTimeTrackLite\eTimeTrackLite1.mdb"


''      dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\eTimeTrackLite1.mdb"
     dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.0.75\d\attendance\att2000.mdb"
     
     
''     If opt_local.Value = True Then
''        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\attendance\att2000.mdb"
''     Else
''        dsnmdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\attendance\att2000.mdb"
''     End If

   
    
    Dim woday As String


  ''  pst_qry = "delete from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "  and a_month = " & cmb_month.ItemData(cmb_month.ListIndex)
  ''   paydb.Execute pst_qry


''    pst_qry = "delete from bio_devicelogs where ad_logdate between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ad_auto = 'A'"
    pst_qry = "delete from bio_devicelogs where ad_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' and ad_auto = 'A'"
    paydb.Execute pst_qry


    pst_qry = "delete from bio_device_shiftlogs where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "'"
    paydb.Execute pst_qry


'''---select MSACESS MDB FILE

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode and userid = '8007' and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid"

''    mdb_qry = "Select * from devicelogs as a, employees as b where a.deviceid <> 1 and a.userid = b.EmployeeCode  and logdate between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date, "MM/dd/yyyy") & "# order by devicelogid "

    
    Dim logid As Long
    Dim io As String
    io = ""
    mdb_qry = "Select *  from checkinout as a, userinfo as b where  a.userid = b.userid  and checktime between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by sersorid "
    mdb_qry = "Select a.*,b.*, a.userid as id  from checkinout as a, userinfo as b where  a.userid = b.userid  and checktime between #" & Format(st_date, "MM/dd/yyyy") & "# and  #" & Format(end_date + 1, "MM/dd/yyyy") & "# order by sensorid,checktime"
    Dim logtype As String
    mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
    logid = 1
    While Not mdbrs.EOF
'         ProgressBar1.Value = ProgressBar1.Value + 1
          logtype = "A"
''         If mdbrs!Deviceid = 1 Then
''            logtype = "M"
''         End If
''         If mdbrs!sensorid = "104" Or mdbrs!sensorid = "106" Then
''            io = "in"
''         ElseIf mdbrs!sensorid = "105" Or mdbrs!sensorid = "107" Then
''            io = "out"
''         End If

         If mdbrs!sensorid = "104" Then
            io = "in"
         ElseIf mdbrs!sensorid = "106" Then
''            If mdbrs!checktype = "I" Then
''                io = "in"
''             Else
                io = "in"
''      End If
         ElseIf mdbrs!sensorid = "105" Then
            io = "out"
         ElseIf mdbrs!sensorid = "107" Then
            io = "out"
         End If


         pst_qry = "insert into  bio_devicelogs (ad_fpcode,ad_empid,ad_logslno,ad_date,ad_logdate,ad_auto,ad_punch) values (  " & mdbrs!Badgenumber & ", " & mdbrs!id & ", " & logid & " , '" & Format(mdbrs!checktime, "MM/dd/yyyy") & "' , '" & Format(mdbrs!checktime, "MM/dd/yyyy HH:MM:SS") & "' , '" & logtype & "','" & io & "')"
         logid = logid + 1
         paydb.Execute pst_qry
         mdbrs.MoveNext
    Wend
    mdbrs.Close

    Dim pst_qrychk As String
    
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 1767"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '1767','1767', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10018"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10018','10018', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10015"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10015','10015', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 1302"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '1302','1302', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10023"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10023','10023', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10014"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10014','10014', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10005"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10005','10005', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10006"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10006','10006', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
       payrs.Close
    
    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = 10010"
    payrs.Open pst_qrychk, paydb, 1, 2
    If payrs!nos = 0 Then
       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  '10010','10010', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
       paydb.Execute pst_qry
    End If
    payrs.Close
    
    
    
    Dim idate, sft_from_date, sft_end_date As Date

   '' If rec_found = 0 Then
        mdb_qry = "Select * from employees where employeecode <> '0' and employeeid = 2819 and Status = 'Working' "
        mdb_qry = "Select * from employees where employeecode <> '0' and Status = 'Working' "
        mdb_qry = "Select * from USERINFO"
        
  ''                 pst_qry = "insert into  bio_empmas (bioemp_company,bioemp_id,bioemp_fpcode,bioemp_name) values (  '1',  " & mdbrs!userid & ", " & mdbrs!badgenumber & ", '" & mdbrs!Name & "' )"
  
        mdbrs.Open mdb_qry, dsnmdb, adOpenDynamic, adLockOptimistic
        While Not mdbrs.EOF
             If Val(mdbrs!userid) > 0 Then
             
''                If Val(mdbrs!Badgenumber) = 3202 Then
''                   MsgBox (mdbrs!Badgenumber)
''                End If
             
             
''              pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!userid & "," & mdbrs!badgenumber & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
''              paydb.Execute pst_qry

                  
                idcode = mdbrs!Badgenumber

''                If idcode = 2071 Then
''                   MsgBox (idcode)
''                End If
                pst_qry1 = "Select count(*) as noofrec from emp_mas where emp_fpcode = " & idcode & " and emp_status = 'A'"
                
                
                payrsnew.Open pst_qry1, paydb, 1, 2
                If payrsnew!noofrec > 0 Then
                    pst_qrychk = "select count(*) as nos from bio_attendlogs where a_year = " & Val(cmb_year.Text) & "   and a_month = " & cmb_month.ItemData(cmb_month.ListIndex) & " and a_fpcode = " & mdbrs!Badgenumber & ""
                    payrs.Open pst_qrychk, paydb, 1, 2
                    If payrs!nos = 0 Then
                       pst_qry = "insert into bio_attendlogs (a_bioid,a_fpcode,a_month,a_year) values (  " & mdbrs!userid & "," & mdbrs!Badgenumber & ", " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                       paydb.Execute pst_qry
                    End If
                    payrs.Close
                    
                    For idate = st_date To end_date
                       pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  " & mdbrs!Badgenumber & ", " & mdbrs!userid & ",  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & "   )"
                       paydb.Execute pst_qry
                       
                    Next
                End If
                payrsnew.Close
             End If
             mdbrs.MoveNext
        Wend
        mdbrs.Close
    
    
                For idate = st_date To end_date
                  
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '1767', '1767',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10018', '10018',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10015', '10015',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '1302', '1302',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10023', '10023',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10014', '10014',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10005', '10005',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10006', '10006',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   pst_qry = "insert into  bio_device_shiftlogs (ds_fpcode,ds_empid,ds_date,ds_month,ds_year) values (  '10010', '10010',  '" & Format(idate, "MM/dd/yyyy") & "', " & cmb_month.ItemData(cmb_month.ListIndex) & " , " & Val(cmb_year.Text) & " )"
                   paydb.Execute pst_qry
                   
                Next
    
''For Adding data in the bio_device_shiftlogs for Outstation working staffs

    
    
    
    
        Dim sftfound As Integer
        sql = "update bio_device_shiftlogs set ds_shift = 'GS',ds_shift_begintime = '08:00',ds_shift_endtime = '17:00',ds_begin_duration  = '60',ds_end_duration = '420' ,ds_shift_in = 0,ds_shift_out = 0,ds_shift_in2 = 0,ds_shift_out2 =0 ,ds_shift_in3 = 0,ds_shift_out3 = 0,ds_shift_in4 = 0,ds_shift_out4 = 0,ds_shift_in5 = 0 ,ds_shift_out5 = 0,ds_shift_in6 = 0,ds_shift_out6 = 0,ds_sft_hrs = 0 ,ds_sft_hrs2 = 0 ,ds_sft_hrs3 = 0 ,ds_sft_hrs4 = 0 ,ds_sft_hrs5 = 0 ,ds_sft_hrs6 = 0 where ds_date between '" & Format(st_date, "MM/dd/yyyy") & "' and  '" & Format(end_date, "MM/dd/yyyy") & "' "
        paydb.Execute sql
        If all_chk = 1 Then MsgBox ("Attendance Logs are uploaded...")
        Exit Sub
''    End If

err_handler:
     chk = gen_Validation(Err.Number, Err.Description)
    '' paydb.RollbackTrans
     Me.MousePointer = 1
  '  chk = gen_Validation(Err.Number, Err.Description)
      If chk = 1 Then Resume
    
    

End Sub


Public Sub find_dates()
    If cmb_month.ListIndex = -1 Then Exit Sub
    If cmb_year.Text = "" Then Exit Sub
    Dim d1 As Date
    mmon = cmb_month.ItemData(cmb_month.ListIndex)
    If mmon = 1 Or mmon = 3 Or mmon = 5 Or mmon = 7 Or mmon = 8 Or mmon = 10 Or mmon = 12 Then
        mdays = 31
    ElseIf mmon = 4 Or mmon = 6 Or mmon = 9 Or mmon = 11 Then
        mdays = 30
    ElseIf mmon = 2 And Val(cmb_year.Text) Mod 4 = 0 Then
        mdays = 29
    Else
        mdays = 28
    End If
    
    end_date = DateValue(Str(mmon) + "/" + Str(mdays) + "/" + cmb_year.Text) + 1
    st_date = end_date - Day(end_date - 1)
    monthend_date.Value = end_date - 1


    If end_date.Value > Now Then
       end_date.Value = Now + 1
    End If
End Sub

Private Sub optFinal_Click()
    ProcessType = "F"
End Sub

Private Sub optRegular_Click()
    ProcessType = "R"
End Sub
